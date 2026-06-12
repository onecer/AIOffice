using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>The M4 shared find/replace contract on docx.</summary>
public sealed class ReplaceTests : WordTestBase
{
    /// <summary>Rebuilds one body paragraph as explicit runs, to model Word's run fragmentation.</summary>
    private static void SplitIntoRuns(string file, int paragraphIndex, params (string Text, bool Bold)[] runs)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: true);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ElementAt(paragraphIndex - 1);
        paragraph.RemoveAllChildren<Run>();
        foreach (var (text, bold) in runs)
        {
            var run = new Run();
            if (bold)
            {
                run.RunProperties = new RunProperties(new Bold());
            }

            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
        }
    }

    private static JsonNode FirstOp(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["data"]!["ops"]!.AsArray()[0]!;

    private static JsonArray? MetaWarnings(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"] as JsonArray;

    [Fact]
    public void Replace_matches_text_split_across_runs_and_keeps_first_affected_run_formatting()
    {
        var file = CreateDoc();
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"placeholder"}}]""");
        SplitIntoRuns(file, 1, ("Hello ", false), ("wor", true), ("ld!", false));

        var envelope = Edit(file, """[{"op":"replace","path":"/body","props":{"find":"world","replace":"earth"}}]""");

        var op = FirstOp(envelope);
        Assert.Equal(1, op["replacements"]!.GetValue<int>());
        Assert.Equal("/body/p[1]", op["locations"]!.AsArray()[0]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
            Assert.Equal("Hello earth!", paragraph.InnerText);

            // The replacement run wears the formatting of the first affected run ("wor", bold).
            var replaced = paragraph.Elements<Run>().Single(r => r.InnerText == "earth");
            Assert.NotNull(replaced.RunProperties?.Bold);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Replace_counts_every_match_and_reports_paragraph_locations()
    {
        var file = CreateDoc();
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"alpha beta alpha"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"alpha"}}
            ]
            """);

        var envelope = Edit(file, """[{"op":"replace","path":"/body","props":{"find":"alpha","replace":"omega"}}]""");

        var op = FirstOp(envelope);
        Assert.Equal(3, op["replacements"]!.GetValue<int>());
        var locations = op["locations"]!.AsArray().Select(l => l!.GetValue<string>()).ToList();
        Assert.Equal(new[] { "/body/p[1]", "/body/p[1]", "/body/p[2]" }, locations);
        Assert.Equal(new[] { "omega beta omega", "omega" }, BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Replace_is_case_insensitive_by_default_and_matchCase_narrows_it()
    {
        var file = CreateDoc();
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Cat cat CAT"}}]""");

        var loose = Edit(file, """[{"op":"replace","path":"/body","props":{"find":"cat","replace":"dog"}}]""");
        Assert.Equal(3, FirstOp(loose)["replacements"]!.GetValue<int>());
        Assert.Equal(new[] { "dog dog dog" }, BodyTexts(file));

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Cat cat CAT"}}]""");
        var strict = Edit(file, """[{"op":"replace","path":"/body","props":{"find":"cat","replace":"dog","matchCase":true}}]""");
        Assert.Equal(1, FirstOp(strict)["replacements"]!.GetValue<int>());
        Assert.Equal(new[] { "Cat dog CAT" }, BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void WholeWord_skips_matches_inside_words()
    {
        var file = CreateDoc();
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"cat concatenate cat."}}]""");

        var envelope = Edit(file, """[{"op":"replace","path":"/body","props":{"find":"cat","replace":"dog","wholeWord":true}}]""");

        Assert.Equal(2, FirstOp(envelope)["replacements"]!.GetValue<int>());
        Assert.Equal(new[] { "dog concatenate dog." }, BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Regex_replace_supports_group_substitution()
    {
        var file = CreateDoc();
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Results for Q1-2025 and Q3-2026."}}]""");

        var envelope = Edit(
            file,
            """[{"op":"replace","path":"/body","props":{"find":"(Q[1-4])-([0-9]{4})","replace":"$2 $1","regex":true}}]""");

        Assert.Equal(2, FirstOp(envelope)["replacements"]!.GetValue<int>());
        Assert.Equal(new[] { "Results for 2025 Q1 and 2026 Q3." }, BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Invalid_regex_is_invalid_args_with_suggestion()
    {
        var file = CreateDoc(title: "Patterns");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"replace","path":"/body","props":{"find":"([unclosed","replace":"x","regex":true}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("regex", ex.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void No_match_is_ok_true_with_find_no_match_warning()
    {
        var file = CreateDoc(title: "Nothing here");

        var envelope = Edit(file, """[{"op":"replace","path":"/body","props":{"find":"absent text","replace":"x"}}]""");

        Assert.True(envelope.IsOk);
        Assert.Equal(0, FirstOp(envelope)["replacements"]!.GetValue<int>());
        Assert.Empty(FirstOp(envelope)["locations"]!.AsArray());
        var warning = Assert.Single(MetaWarnings(envelope)!);
        Assert.Equal("find_no_match", warning!["code"]!.GetValue<string>());
    }

    [Fact]
    public void Empty_replacement_deletes_the_matched_text()
    {
        var file = CreateDoc();
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"remove me please"}}]""");

        Edit(file, """[{"op":"replace","path":"/body","props":{"find":" me","replace":""}}]""");

        Assert.Equal(new[] { "remove please" }, BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Scope_limits_replacement_to_one_paragraph_header_or_footer()
    {
        var file = CreateDoc();
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"draft body"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"draft tail"}},
              {"op":"add","path":"/header[1]","type":"header","props":{"text":"draft header"}},
              {"op":"add","path":"/footer[1]","type":"footer","props":{"text":"draft footer"}}
            ]
            """);

        // Single paragraph scope: only p[2] changes.
        Edit(file, """[{"op":"replace","path":"/body/p[2]","props":{"find":"draft","replace":"final"}}]""");
        Assert.Equal(new[] { "draft body", "final tail" }, BodyTexts(file));

        // Header scope: the body stays untouched.
        var headerOp = FirstOp(Edit(file, """[{"op":"replace","path":"/header[1]","props":{"find":"draft","replace":"FINAL"}}]"""));
        Assert.Equal(1, headerOp["replacements"]!.GetValue<int>());
        Assert.Equal("/header[1]/p[1]", headerOp["locations"]!.AsArray()[0]!.GetValue<string>());

        var footerOp = FirstOp(Edit(file, """[{"op":"replace","path":"/footer[1]","props":{"find":"draft","replace":"FINAL"}}]"""));
        Assert.Equal(1, footerOp["replacements"]!.GetValue<int>());

        Assert.Equal(new[] { "draft body", "final tail" }, BodyTexts(file));
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Equal("FINAL header", doc.MainDocumentPart!.HeaderParts.Single().Header!.InnerText);
            Assert.Equal("FINAL footer", doc.MainDocumentPart.FooterParts.Single().Footer!.InnerText);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_replace_produces_del_ins_revision_pairs_that_accept_cleanly()
    {
        var file = CreateDoc();
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"The old wording stands"}}]""");

        var envelope = Edit(
            file,
            """[{"op":"replace","path":"/body","props":{"find":"old","replace":"new"}}]""",
            new JsonObject { ["track"] = true, ["author"] = "Reviewer" });

        Assert.True(FirstOp(envelope)["tracked"]!.GetValue<bool>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
            var del = Assert.Single(paragraph.Descendants<DeletedRun>());
            Assert.Equal("old", del.InnerText);
            Assert.Single(del.Descendants<DeletedText>()); // w:t became w:delText inside w:del
            var ins = Assert.Single(paragraph.Descendants<InsertedRun>());
            Assert.Equal("new", ins.InnerText);
            Assert.Contains(del.GetAttributes(), a => a.LocalName == "author" && a.Value == "Reviewer");
        }

        // Pending state reads as old+new; accepting settles on the replacement.
        var revisions = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "revisions" })));
        Assert.Equal(2, revisions["count"]!.GetValue<int>());

        Edit(file, """[{"op":"accept","path":"/body"}]""");
        Assert.Equal(new[] { "The new wording stands" }, BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_replace_rejected_restores_the_original_text()
    {
        var file = CreateDoc();
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"keep this phrase"}}]""");

        Edit(
            file,
            """[{"op":"replace","path":"/body","props":{"find":"this","replace":"that"}}]""",
            new JsonObject { ["track"] = true });
        Edit(file, """[{"op":"reject","path":"/body"}]""");

        Assert.Equal(new[] { "keep this phrase" }, BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_replace_outside_body_is_unsupported()
    {
        var file = CreateDoc(title: "Headed");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"draft"}}]""");

        var ex = Assert.Throws<AiofficeException>(() => Edit(
            file,
            """[{"op":"replace","path":"/header[1]","props":{"find":"draft","replace":"final"}}]""",
            new JsonObject { ["track"] = true }));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    [Fact]
    public void Replace_requires_find()
    {
        var file = CreateDoc(title: "Empty find");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"replace","path":"/body","props":{"replace":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("find", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Replace_scope_must_be_a_container()
    {
        var file = CreateDoc(title: "Scoped");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"replace","path":"/body/p[1]/run[1]","props":{"find":"a","replace":"b"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("/body/p[1]", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Replace_inside_a_table_cell_scope_works()
    {
        var file = CreateDoc();
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":2}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"total: 100"}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"total: 200"}}
            ]
            """);

        var envelope = Edit(file, """[{"op":"replace","path":"/body/table[1]/tr[1]/tc[1]","props":{"find":"total","replace":"sum"}}]""");

        Assert.Equal(1, FirstOp(envelope)["replacements"]!.GetValue<int>());
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var cells = doc.MainDocumentPart!.Document!.Body!.Descendants<TableCell>().ToList();
            Assert.Equal("sum: 100", cells[0].InnerText);
            Assert.Equal("total: 200", cells[1].InnerText);
        }

        AssertValidatesClean(file);
    }
}
