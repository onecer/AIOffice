using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class FootnoteTests : WordTestBase
{
    [Fact]
    public void Add_footnote_creates_part_with_separator_defaults_and_reference_run()
    {
        var file = CreateDoc(title: "Cited claim");

        var envelope = Edit(file, """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"Source: annual report."}}]""");
        Assert.Equal("/footnote[@id=1]", Data(envelope)["ops"]!.AsArray()[0]!["path"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var footnotes = doc.MainDocumentPart!.FootnotesPart!.Footnotes!;
            var all = footnotes.Elements<Footnote>().ToList();
            Assert.Contains(all, f => f.Type?.Value == FootnoteEndnoteValues.Separator && f.Id?.Value == -1);
            Assert.Contains(all, f => f.Type?.Value == FootnoteEndnoteValues.ContinuationSeparator && f.Id?.Value == 0);

            var note = Assert.Single(all, f => f.Type is null || f.Type.Value == FootnoteEndnoteValues.Normal);
            Assert.Equal(1, note.Id!.Value);
            Assert.Contains("Source: annual report.", note.InnerText, StringComparison.Ordinal);

            // The reference run sits at the end of the paragraph, superscripted.
            var paragraph = doc.MainDocumentPart.Document!.Body!.Elements<Paragraph>().First();
            var reference = Assert.Single(paragraph.Descendants<FootnoteReference>());
            Assert.Equal(1, reference.Id!.Value);
            var run = Assert.IsType<Run>(reference.Parent);
            Assert.NotNull(run.RunProperties?.VerticalTextAlignment);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Footnote_ids_increment_and_get_resolves_them()
    {
        var file = CreateDoc(title: "Twice noted");
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"first"}},
              {"op":"add","path":"/body/p[2]","type":"footnote","props":{"text":"second"}}
            ]
            """);

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/footnote[@id=2]" })));

        Assert.Equal("footnote", got["type"]!.GetValue<string>());
        Assert.Equal("second", got["properties"]!["text"]!.GetValue<string>());
        Assert.Equal("/body/p[2]", got["properties"]!["anchorPath"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Text_view_appends_markers_and_a_footnote_section()
    {
        var file = CreateDoc(title: "Marked claim");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"the fine print"}}]""");

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "text" })));

        Assert.Equal("Marked claim[^1]", data["lines"]!.AsArray()[0]!["text"]!.GetValue<string>());
        var footnote = Assert.Single(data["footnotes"]!.AsArray())!;
        Assert.Equal(1, footnote["id"]!.GetValue<int>());
        Assert.Equal("the fine print", footnote["text"]!.GetValue<string>());
    }

    [Fact]
    public void Html_render_emits_sup_reference_and_footer_list()
    {
        var file = CreateDoc(title: "Rendered claim");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"see appendix"}}]""");

        var html = Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();

        Assert.Contains("""<sup data-aio-path="/footnote[@id=1]">1</sup>""", html, StringComparison.Ordinal);
        Assert.Contains("<section class=\"footnotes\">", html, StringComparison.Ordinal);
        Assert.Contains("""<li data-aio-path="/footnote[@id=1]">see appendix</li>""", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Remove_footnote_clears_part_entry_and_reference()
    {
        var file = CreateDoc(title: "Fleeting");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"temp"}}]""");

        Edit(file, """[{"op":"remove","path":"/footnote[@id=1]"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<FootnoteReference>());
            Assert.DoesNotContain(
                doc.MainDocumentPart.FootnotesPart!.Footnotes!.Elements<Footnote>(),
                f => f.Type is null || f.Type.Value == FootnoteEndnoteValues.Normal);
        }

        AssertValidatesClean(file);
    }

    // Endnotes_are_unsupported_naming_footnotes was removed in M4: endnotes are
    // a real feature now (same surface as footnotes) — see EndnoteTests.

    [Fact]
    public void Footnote_text_is_required()
    {
        var file = CreateDoc(title: "Silent");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("text", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Footnote_outside_body_is_invalid_args()
    {
        var file = CreateDoc(title: "Headed");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"H"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/header[1]/p[1]","type":"footnote","props":{"text":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Unknown_footnote_id_is_invalid_path_with_candidates()
    {
        var file = CreateDoc(title: "Missing");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"only"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"remove","path":"/footnote[@id=9]"}]"""));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/footnote[@id=1]", ex.Candidates!);
    }

    // -------------------------------------------------------------- tracked

    [Fact]
    public void Untracked_footnote_add_is_byte_stable_no_track_markup()
    {
        // The untracked append-Run path is first-in-code and timestamp-free. The
        // resulting document/footnotes XML is deterministic, so the same op on two
        // identical docs yields byte-identical PARTS (the OPC zip carries wall-clock
        // entry timestamps, so we compare part bytes, not the whole package) — proof
        // the v1.15 behaviour is unchanged by the new session-threaded gate.
        var a = CreateDoc("a.docx", title: "Same input");
        var b = CreateDoc("b.docx", title: "Same input");
        const string op = """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"Footnote body."}}]""";

        var envelope = Edit(a, op);
        Edit(b, op);

        // Existing return shape unchanged: {op,type,path,anchor}, no "tracked".
        var added = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("add", added["op"]!.GetValue<string>());
        Assert.Equal("footnote", added["type"]!.GetValue<string>());
        Assert.Equal("/footnote[@id=1]", added["path"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", added["anchor"]!.GetValue<string>());
        Assert.Null(added["tracked"]);

        Assert.Equal(DocumentPartXml(a), DocumentPartXml(b));
        Assert.Equal(FootnotesPartXml(a), FootnotesPartXml(b));

        // And no tracked markup leaked into the untracked package.
        using var doc = WordprocessingDocument.Open(a, isEditable: false);
        Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<InsertedRun>());
    }

    private static string DocumentPartXml(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        return doc.MainDocumentPart!.Document!.OuterXml;
    }

    private static string FootnotesPartXml(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        return doc.MainDocumentPart!.FootnotesPart!.Footnotes!.OuterXml;
    }

    [Fact]
    public void Tracked_footnote_add_wraps_only_the_reference_run_in_an_insertion()
    {
        var file = CreateDoc(title: "Pre-existing body text");

        // Pre-seed a tracked formatting set so the revision counter is already past 1
        // (the rPrChange takes id 1) while leaving the title run live and in place.
        // The footnote's own @w:id space restarts at 1, so the resulting w:ins @w:id
        // (2) is provably NOT the footnote @w:id (1): the two id spaces are independent.
        Edit(
            file,
            """[{"op":"set","path":"/body/p[1]","props":{"bold":true}}]""",
            new JsonObject { ["track"] = true });

        Edit(
            file,
            """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"Tracked note."}}]""",
            new JsonObject { ["track"] = true });

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();

        // Exactly one w:ins, and it wraps the FootnoteReference run only.
        var ins = Assert.Single(paragraph.Descendants<InsertedRun>());
        var reference = Assert.Single(ins.Descendants<FootnoteReference>());
        Assert.Equal(1, reference.Id!.Value);

        // w:ins carries id (NextRevisionId) + author + date.
        var (insId, author, date) = TrackAttributes(ins);
        Assert.True(insId > 1); // advanced past the seed's rPrChange
        Assert.Equal("AIOffice", author);
        Assert.False(string.IsNullOrEmpty(date));

        // Distinct id spaces: the w:ins @w:id is NOT the footnote @w:id.
        Assert.NotEqual(reference.Id!.Value, insId);

        // The author's pre-existing title run is NOT inside the w:ins, and the
        // paragraph mark is NOT flagged inserted.
        Assert.Contains(
            paragraph.Elements<Run>(),
            r => r.InnerText.Contains("Pre-existing body text", StringComparison.Ordinal));
        Assert.Empty(ParagraphMarkInsertions(paragraph));

        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_footnote_add_outside_body_stays_refused()
    {
        var file = CreateDoc(title: "Headed");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"H"}}]""");

        var ex = Assert.Throws<AiofficeException>(() => Edit(
            file,
            """[{"op":"add","path":"/header[1]/p[1]","type":"footnote","props":{"text":"x"}}]""",
            new JsonObject { ["track"] = true }));

        // The /body guard inside ApplyAddFootnote refuses it before any w:ins.
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Tracked_footnote_removal_is_structural_not_a_tracked_deletion()
    {
        // Removing a footnote drops the part entry + every reference run; it is a
        // structural delete that ignores track (a w:del cannot host a note-part
        // removal). This pre-existing behaviour is untouched: no w:del is authored.
        var file = CreateDoc(title: "Noted");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"only"}}]""");

        Edit(file, """[{"op":"remove","path":"/footnote[@id=1]"}]""", new JsonObject { ["track"] = true });

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Empty(body.Descendants<FootnoteReference>());
        Assert.Empty(body.Descendants<DeletedRun>()); // not a tracked deletion
        AssertValidatesClean(file);
    }

    /// <summary>w:id / w:author / w:date off a track-change element.</summary>
    internal static (int Id, string? Author, string? Date) TrackAttributes(OpenXmlElement element)
    {
        string? id = null, author = null, date = null;
        foreach (var attribute in element.GetAttributes())
        {
            switch (attribute.LocalName)
            {
                case "id": id = attribute.Value; break;
                case "author": author = attribute.Value; break;
                case "date": date = attribute.Value; break;
                default: break;
            }
        }

        return (int.TryParse(id, out var n) ? n : 0, author, date);
    }

    /// <summary>Any w:ins flagged on the paragraph mark's run properties.</summary>
    internal static IEnumerable<Inserted> ParagraphMarkInsertions(Paragraph paragraph) =>
        paragraph.ParagraphProperties?.ParagraphMarkRunProperties?.Elements<Inserted>() ?? [];
}
