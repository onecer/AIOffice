using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class TrackedChangesTests : WordTestBase
{
    private static JsonObject Track(string? author = null)
    {
        var args = new JsonObject { ["track"] = true };
        if (author is not null)
        {
            args["author"] = author;
        }

        return args;
    }

    private JsonArray Revisions(string file) =>
        Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "revisions" })))["revisions"]!.AsArray();

    // ------------------------------------------------------------ producing

    [Fact]
    public void Tracked_set_text_produces_delete_plus_insert_revisions()
    {
        var file = CreateDoc(title: "Hello");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"World"}}]""", Track());

        var revisions = Revisions(file);
        Assert.Equal(2, revisions.Count);
        Assert.Equal("delete", revisions[0]!["kind"]!.GetValue<string>());
        Assert.Equal("Hello", revisions[0]!["text"]!.GetValue<string>());
        Assert.Equal("insert", revisions[1]!["kind"]!.GetValue<string>());
        Assert.Equal("World", revisions[1]!["text"]!.GetValue<string>());
        Assert.All(revisions, r => Assert.Equal("AIOffice", r!["author"]!.GetValue<string>()));
        Assert.All(revisions, r => Assert.Equal("/body/p[1]", r!["at"]!.GetValue<string>()));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Revision_ids_are_sequential_and_paths_canonical()
    {
        var file = CreateDoc(title: "One");
        Edit(file, """[{"op":"add","path":"/body","props":{"text":"Two"}}]""");

        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"ONE"}},
              {"op":"set","path":"/body/p[3]","props":{"text":"TWO"}}
            ]
            """, Track());

        var revisions = Revisions(file);
        var ids = revisions.Select(r => r!["id"]!.GetValue<int>()).ToList();
        Assert.Equal(ids.OrderBy(i => i), ids);
        Assert.Equal(ids.Distinct().Count(), ids.Count);
        Assert.All(revisions, r => Assert.Equal(
            $"/revision[@id={r!["id"]!.GetValue<int>()}]",
            r!["path"]!.GetValue<string>()));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Author_resolution_prefers_op_props_then_batch_then_default()
    {
        var file = CreateDoc(title: "Attribution");

        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"by op author","author":"Bob"}},
              {"op":"add","path":"/body","props":{"text":"by batch author"}}
            ]
            """, Track(author: "Alice"));

        var revisions = Revisions(file);
        Assert.Contains(revisions, r => r!["author"]!.GetValue<string>() == "Bob");
        Assert.Contains(revisions, r => r!["author"]!.GetValue<string>() == "Alice");
        Assert.DoesNotContain(revisions, r => r!["author"]!.GetValue<string>() == "AIOffice");
        AssertValidatesClean(file);
    }

    [Fact]
    public void Untracked_edits_produce_no_revisions()
    {
        var file = CreateDoc(title: "Plain");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"still plain"}}]""");

        Assert.Empty(Revisions(file));
    }

    // -------------------------------------------------------- accept/reject

    [Fact]
    public void Accept_all_in_scope_applies_the_text_change()
    {
        var file = CreateDoc(title: "Hello");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"World"}}]""", Track());

        var envelope = Edit(file, """[{"op":"accept","path":"/body"}]""");

        Assert.Equal(2, Data(envelope)["ops"]!.AsArray()[0]!["applied"]!.GetValue<int>());
        Assert.Equal("World", BodyTexts(file)[0]);
        Assert.Empty(Revisions(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Reject_all_in_scope_restores_the_original_text()
    {
        var file = CreateDoc(title: "Hello");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"World"}}]""", Track());

        Edit(file, """[{"op":"reject","path":"/body"}]""");

        Assert.Equal("Hello", BodyTexts(file)[0]);
        Assert.Empty(Revisions(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Accept_single_revision_by_id_leaves_the_other_pending()
    {
        var file = CreateDoc(title: "Hello");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"World"}}]""", Track());
        var insertId = Revisions(file).First(r => r!["kind"]!.GetValue<string>() == "insert")!["id"]!.GetValue<int>();

        Edit(file, $$"""[{"op":"accept","path":"/revision[@id={{insertId}}]"}]""");

        var remaining = Revisions(file);
        var only = Assert.Single(remaining);
        Assert.Equal("delete", only!["kind"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_added_paragraph_accepts_to_plain_content_and_rejects_to_nothing()
    {
        var fileAccept = CreateDoc("accept.docx", title: "Base");
        Edit(fileAccept, """[{"op":"add","path":"/body","props":{"text":"Inserted paragraph"}}]""", Track());
        Assert.NotEmpty(Revisions(fileAccept));

        Edit(fileAccept, """[{"op":"accept","path":"/body"}]""");
        Assert.Contains("Inserted paragraph", BodyTexts(fileAccept));
        Assert.Empty(Revisions(fileAccept));
        AssertValidatesClean(fileAccept);

        var fileReject = CreateDoc("reject.docx", title: "Base");
        var before = BodyTexts(fileReject);
        Edit(fileReject, """[{"op":"add","path":"/body","props":{"text":"Inserted paragraph"}}]""", Track());

        Edit(fileReject, """[{"op":"reject","path":"/body"}]""");
        Assert.Equal(before, BodyTexts(fileReject));
        Assert.Empty(Revisions(fileReject));
        AssertValidatesClean(fileReject);
    }

    [Fact]
    public void Tracked_removed_paragraph_accepts_to_gone_and_rejects_to_restored()
    {
        var fileAccept = CreateDoc("accept.docx", title: "Keep");
        Edit(fileAccept, """[{"op":"add","path":"/body","props":{"text":"Doomed"}}]""");
        Edit(fileAccept, """[{"op":"remove","path":"/body/p[3]"}]""", Track());
        Assert.Contains("Doomed", BodyTexts(fileAccept)); // still visible while pending

        Edit(fileAccept, """[{"op":"accept","path":"/body"}]""");
        Assert.DoesNotContain("Doomed", BodyTexts(fileAccept));
        Assert.Empty(Revisions(fileAccept));
        AssertValidatesClean(fileAccept);

        var fileReject = CreateDoc("reject.docx", title: "Keep");
        Edit(fileReject, """[{"op":"add","path":"/body","props":{"text":"Saved"}}]""");
        Edit(fileReject, """[{"op":"remove","path":"/body/p[3]"}]""", Track());

        Edit(fileReject, """[{"op":"reject","path":"/body"}]""");
        Assert.Contains("Saved", BodyTexts(fileReject));
        Assert.Empty(Revisions(fileReject));
        AssertValidatesClean(fileReject);
    }

    [Fact]
    public void Scope_accept_on_one_paragraph_leaves_other_paragraphs_pending()
    {
        var file = CreateDoc(title: "First");
        Edit(file, """[{"op":"add","path":"/body","props":{"text":"Second"}}]""");
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"FIRST"}},
              {"op":"set","path":"/body/p[3]","props":{"text":"SECOND"}}
            ]
            """, Track());

        Edit(file, """[{"op":"accept","path":"/body/p[1]"}]""");

        Assert.Equal("FIRST", BodyTexts(file)[0]);
        var remaining = Revisions(file);
        Assert.Equal(2, remaining.Count);
        Assert.All(remaining, r => Assert.Equal("/body/p[3]", r!["at"]!.GetValue<string>()));
        AssertValidatesClean(file);
    }

    // --------------------------------------------------------------- honesty

    [Fact]
    public void Tracked_formatting_change_is_unsupported_until_m3()
    {
        var file = CreateDoc(title: "Styled");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"bold":true}}]""", Track()));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("track", ex.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Plants a Word-authored w:rPrChange: run is now bold, was previously unformatted.</summary>
    private static void PlantRunFormatRevision(string file, int id = 901)
    {
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: true);
        var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First();
        var rPr = run.RunProperties ??= new RunProperties();
        rPr.Bold = new Bold(); // the NEW formatting
        rPr.AppendChild(new RunPropertiesChange
        {
            Id = id.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Author = "Word User",
            Date = DateTime.UtcNow,
            PreviousRunProperties = new PreviousRunProperties(), // the OLD formatting: none
        });
    }

    [Fact]
    public void Format_revision_reads_as_format_and_accept_keeps_new_formatting()
    {
        var file = CreateDoc(title: "Formatted");
        PlantRunFormatRevision(file);

        var revisions = Revisions(file);
        var format = Assert.Single(revisions);
        Assert.Equal("format", format!["kind"]!.GetValue<string>());
        Assert.Equal(901, format["id"]!.GetValue<int>());

        Edit(file, """[{"op":"accept","path":"/revision[@id=901]"}]""");

        Assert.Empty(Revisions(file));
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false))
        {
            var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First();
            Assert.NotNull(run.RunProperties?.Bold); // accept keeps the new formatting
            Assert.Empty(doc.MainDocumentPart.Document.Body.Descendants<RunPropertiesChange>());
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Format_revision_reject_restores_the_previous_run_properties()
    {
        var file = CreateDoc(title: "Formatted");
        PlantRunFormatRevision(file);

        Edit(file, """[{"op":"reject","path":"/revision[@id=901]"}]""");

        Assert.Empty(Revisions(file));
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false))
        {
            var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First();
            Assert.Null(run.RunProperties?.Bold); // reject restores the old (unformatted) state
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Paragraph_format_revision_reject_restores_previous_paragraph_properties()
    {
        var file = CreateDoc(title: "Centered");

        // Word-authored w:pPrChange: paragraph is now centered; the marker's inner
        // pPr holds the COMPLETE previous properties (Heading1, no jc).
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: true))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
            var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();
            pPr.Justification = new Justification { Val = JustificationValues.Center };
            pPr.AppendChild(new ParagraphPropertiesChange
            {
                Id = "902",
                Author = "Word User",
                Date = DateTime.UtcNow,
                ParagraphPropertiesExtended = new ParagraphPropertiesExtended(
                    new ParagraphStyleId { Val = "Heading1" }),
            });
        }

        Edit(file, """[{"op":"reject","path":"/revision[@id=902]"}]""");

        Assert.Empty(Revisions(file));
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
            Assert.Null(paragraph.ParagraphProperties?.Justification); // the centering never happened
            Assert.Equal("Heading1", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value); // old props restored
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Scope_accept_resolves_format_revisions_too()
    {
        var file = CreateDoc(title: "Formatted");
        PlantRunFormatRevision(file);

        var envelope = Edit(file, """[{"op":"accept","path":"/body"}]""");

        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal(1, summary["applied"]!.GetValue<int>());
        Assert.Empty(Revisions(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_table_add_is_unsupported()
    {
        var file = CreateDoc(title: "Tables");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}}]""", Track()));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("track", ex.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Tracked_edit_outside_body_is_unsupported()
    {
        var file = CreateDoc(title: "Heads");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"H"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/header[1]/p[1]","props":{"text":"tracked?"}}]""", Track()));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    // ------------------------------------------------------------ addressing

    [Fact]
    public void Get_revision_by_id_and_by_position_agree()
    {
        var file = CreateDoc(title: "Hello");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"World"}}]""", Track());
        var firstId = Revisions(file)[0]!["id"]!.GetValue<int>();

        var byId = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = $"/revision[@id={firstId}]" })));
        var byPosition = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/revision[1]" })));

        Assert.Equal("revision", byId["type"]!.GetValue<string>());
        Assert.Equal(byId["properties"]!["id"]!.GetValue<int>(), byPosition["properties"]!["id"]!.GetValue<int>());
        Assert.Equal("delete", byId["properties"]!["kind"]!.GetValue<string>());
    }

    [Fact]
    public void Unknown_revision_id_is_invalid_path_with_candidates()
    {
        var file = CreateDoc(title: "Hello");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"World"}}]""", Track());

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"accept","path":"/revision[@id=999]"}]"""));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.All(ex.Candidates!, c => Assert.StartsWith("/revision[@id=", c, StringComparison.Ordinal));
    }

    [Fact]
    public void Remove_op_on_a_revision_points_to_accept_reject()
    {
        var file = CreateDoc(title: "Hello");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"World"}}]""", Track());

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"remove","path":"/revision[1]"}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("accept", ex.Suggestion, StringComparison.Ordinal);
    }

    // ------------------------------------------------------- reopen contract

    [Fact]
    public void Tracked_changes_survive_reopen_with_delText_markup()
    {
        var file = CreateDoc(title: "Persist me");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Persisted"}}]""", Track());

        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var del = Assert.Single(body.Descendants<DeletedRun>());
        Assert.Single(body.Descendants<InsertedRun>());
        Assert.Equal("Persist me", string.Concat(del.Descendants<DeletedText>().Select(t => t.Text)));
        Assert.Empty(del.Descendants<Text>()); // w:t must become w:delText inside w:del
    }
}
