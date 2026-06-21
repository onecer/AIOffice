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
    public void Tracked_formatting_set_in_a_header_is_still_unsupported()
    {
        var file = CreateDoc(title: "Heads");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"H"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/header[1]/p[1]","props":{"bold":true}}]""", Track()));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    [Fact]
    public void Tracked_formatting_set_on_a_table_cell_is_still_unsupported()
    {
        var file = CreateDoc(title: "Cells");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"bold":true}}]""", Track()));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("table cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ------------------------------------------ authored formatting revisions

    [Fact]
    public void Tracked_text_set_stays_on_the_byte_stable_del_ins_path()
    {
        // The text-only branch is first-in-code and untouched: a tracked text set
        // must still produce ONLY w:del + w:ins with no formatting markers, and
        // the body XML must match a control edit modulo the wall-clock w:date.
        var baseline = CreateDoc("baseline.docx", title: "Original");
        var widened = CreateDoc("widened.docx", title: "Original");

        Edit(baseline, """[{"op":"set","path":"/body/p[1]","props":{"text":"New"}}]""", Track());
        Edit(widened, """[{"op":"set","path":"/body/p[1]","props":{"text":"New"}}]""", Track());

        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(widened, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Single(body.Descendants<DeletedRun>());
            Assert.Single(body.Descendants<InsertedRun>());
            Assert.Empty(body.Descendants<RunPropertiesChange>());
            Assert.Empty(body.Descendants<ParagraphPropertiesChange>());
        }

        // Identical document XML once the per-edit timestamps are normalized.
        Assert.Equal(BodyXmlSansDate(baseline), BodyXmlSansDate(widened));
        AssertValidatesClean(widened);
    }

    /// <summary>The main document XML with every w:date attribute stripped (so two edits compare).</summary>
    private static string BodyXmlSansDate(string file)
    {
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false);
        var xml = doc.MainDocumentPart!.Document!.OuterXml;
        return System.Text.RegularExpressions.Regex.Replace(xml, "w:date=\"[^\"]*\"", string.Empty);
    }

    [Fact]
    public void Tracked_bold_set_authors_an_rPrChange_with_previous_state()
    {
        var file = CreateDoc(title: "Plain");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"bold":true}}]""", Track(author: "Reviewer"));

        var revision = Assert.Single(Revisions(file));
        Assert.Equal("format", revision!["kind"]!.GetValue<string>());
        Assert.Equal("Reviewer", revision["author"]!.GetValue<string>());
        Assert.True(revision["id"]!.GetValue<int>() > 0);
        Assert.NotNull(revision["date"]!.GetValue<string>());

        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false))
        {
            var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First();
            Assert.NotNull(run.RunProperties?.Bold); // live run carries w:b
            var change = Assert.Single(run.RunProperties!.Elements<RunPropertiesChange>());
            Assert.NotNull(change.Id);
            Assert.Equal("Reviewer", change.Author?.Value);
            Assert.NotNull(change.Date);
            // PreviousRunProperties holds the prior (unbold) state.
            Assert.Null(change.PreviousRunProperties?.GetFirstChild<Bold>());
        }

        // Reopen round-trips the authored change cleanly.
        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_text_and_formatting_together_produce_del_ins_and_rPrChange_on_the_inserted_run()
    {
        var file = CreateDoc(title: "Before");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"New","bold":true,"fontSize":14}}]""", Track());

        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Single(body.Descendants<DeletedRun>()); // old text deleted
        var ins = Assert.Single(body.Descendants<InsertedRun>());
        var insertedRun = ins.Elements<Run>().First();

        Assert.NotNull(insertedRun.RunProperties?.Bold);
        Assert.Equal("28", insertedRun.RunProperties?.FontSize?.Val?.Value); // 14pt -> 28 half-points
        // The rPrChange lives on the INSERTED run, snapshotting the pre-format state.
        var change = Assert.Single(insertedRun.RunProperties!.Elements<RunPropertiesChange>());
        Assert.Null(change.PreviousRunProperties?.GetFirstChild<Bold>());
        Assert.Null(change.PreviousRunProperties?.GetFirstChild<FontSize>());

        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_style_set_authors_a_pPrChange_reading_as_format()
    {
        var file = CreateDoc(title: "Body text");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"style":"Heading1"}}]""", Track());

        var revision = Assert.Single(Revisions(file));
        Assert.Equal("format", revision!["kind"]!.GetValue<string>());

        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
            var pPr = paragraph.ParagraphProperties!;
            Assert.Equal("Heading1", pPr.ParagraphStyleId?.Val?.Value); // live style applied
            var change = Assert.Single(pPr.Elements<ParagraphPropertiesChange>());
            Assert.NotNull(change.ParagraphPropertiesExtended); // snapshot present
            Assert.NotNull(change.Id);
            Assert.NotNull(change.Author);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Authored_rPrChange_accepts_to_new_and_rejects_to_previous()
    {
        var fileAccept = CreateDoc("accept.docx", title: "Plain");
        Edit(fileAccept, """[{"op":"set","path":"/body/p[1]","props":{"bold":true}}]""", Track());
        var id = Revisions(fileAccept).First()!["id"]!.GetValue<int>();

        Edit(fileAccept, $$"""[{"op":"accept","path":"/revision[@id={{id}}]"}]""");
        Assert.Empty(Revisions(fileAccept));
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(fileAccept, isEditable: false))
        {
            var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First();
            Assert.NotNull(run.RunProperties?.Bold); // accept keeps the new bold
            Assert.Empty(doc.MainDocumentPart.Document.Body.Descendants<RunPropertiesChange>());
        }

        AssertValidatesClean(fileAccept);

        var fileReject = CreateDoc("reject.docx", title: "Plain");
        Edit(fileReject, """[{"op":"set","path":"/body/p[1]","props":{"bold":true}}]""", Track());
        var rejectId = Revisions(fileReject).First()!["id"]!.GetValue<int>();

        Edit(fileReject, $$"""[{"op":"reject","path":"/revision[@id={{rejectId}}]"}]""");
        Assert.Empty(Revisions(fileReject));
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(fileReject, isEditable: false))
        {
            var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First();
            Assert.Null(run.RunProperties?.Bold); // reject restores the unbold previous state
        }

        AssertValidatesClean(fileReject);
    }

    [Fact]
    public void Authored_pPrChange_rejects_to_previous_paragraph_props()
    {
        // The title paragraph is Heading1; add a plain (unstyled) paragraph and
        // track a style change on it so reject restores the no-style state.
        var file = CreateDoc(title: "Normal");
        Edit(file, """[{"op":"add","path":"/body","props":{"text":"Plain"}}]""");

        Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"style":"Heading1"}}]""", Track());
        var id = Revisions(file).First()!["id"]!.GetValue<int>();

        Edit(file, $$"""[{"op":"reject","path":"/revision[@id={{id}}]"}]""");

        Assert.Empty(Revisions(file));
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ElementAt(1);
            // The added paragraph had no explicit style; reject restores that (no Heading1).
            Assert.Null(paragraph.ParagraphProperties?.ParagraphStyleId);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Authored_pPrChange_accepts_to_new_paragraph_style()
    {
        // Accept of an authored w:pPrChange must drop the marker and KEEP the new style.
        var file = CreateDoc(title: "Normal");
        Edit(file, """[{"op":"add","path":"/body","props":{"text":"Plain"}}]""");

        Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"style":"Heading1"}}]""", Track());
        var id = Revisions(file).First()!["id"]!.GetValue<int>();

        Edit(file, $$"""[{"op":"accept","path":"/revision[@id={{id}}]"}]""");

        Assert.Empty(Revisions(file));
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ElementAt(1);
            Assert.Equal("Heading1", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value); // accept keeps the new style
            Assert.Empty(paragraph.Descendants<ParagraphPropertiesChange>()); // marker gone
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Untracked_formatting_set_produces_no_revision()
    {
        var file = CreateDoc(title: "Plain");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"bold":true}}]""");

        Assert.Empty(Revisions(file));
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false);
        Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<RunPropertiesChange>());
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

    // ---------------------------------------------- tracked footnote/endnote

    [Fact]
    public void Tracked_footnote_add_reports_one_insert_and_no_spurious_revision()
    {
        var file = CreateDoc(title: "Body stays clean");

        Edit(
            file,
            """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"Note."}}]""",
            Track());

        // Exactly one insert entry, anchored at the body paragraph; the
        // pre-existing title text produces NO revision of its own.
        var revision = Assert.Single(Revisions(file));
        Assert.Equal("insert", revision!["kind"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", revision["at"]!.GetValue<string>());
        Assert.Null(revision["mark"]); // not a paragraph-mark insertion
        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_endnote_add_reports_one_insert_and_no_spurious_revision()
    {
        var file = CreateDoc(title: "Body stays clean");

        Edit(
            file,
            """[{"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"Note."}}]""",
            Track());

        var revision = Assert.Single(Revisions(file));
        Assert.Equal("insert", revision!["kind"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", revision["at"]!.GetValue<string>());
        Assert.Null(revision["mark"]);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_footnote_add_accepts_to_kept_reference_and_rejects_to_gone()
    {
        var fileAccept = CreateDoc("accept.docx", title: "Keep");
        Edit(fileAccept, """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"Kept."}}]""", Track());
        Edit(fileAccept, """[{"op":"accept","path":"/body"}]""");

        Assert.Empty(Revisions(fileAccept));
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(fileAccept, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<InsertedRun>());
            // Accept keeps the reference run and the note part content.
            Assert.Single(body.Descendants<FootnoteReference>());
            Assert.Contains("Kept.", doc.MainDocumentPart.FootnotesPart!.Footnotes!.InnerText, StringComparison.Ordinal);
        }

        AssertValidatesClean(fileAccept);

        var fileReject = CreateDoc("reject.docx", title: "Keep");
        Edit(fileReject, """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"Dropped."}}]""", Track());
        Edit(fileReject, """[{"op":"reject","path":"/body"}]""");

        Assert.Empty(Revisions(fileReject));
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(fileReject, isEditable: false))
        {
            // Reject drops the reference run; the orphan note part is benign.
            Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<FootnoteReference>());
        }

        AssertValidatesClean(fileReject); // still validates clean despite the orphan note
    }

    [Fact]
    public void Tracked_endnote_add_rejects_to_gone_and_validates_clean()
    {
        var file = CreateDoc(title: "Keep");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"Dropped."}}]""", Track());
        Edit(file, """[{"op":"reject","path":"/body"}]""");

        Assert.Empty(Revisions(file));
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<EndnoteReference>());
        }

        AssertValidatesClean(file);
    }
}
