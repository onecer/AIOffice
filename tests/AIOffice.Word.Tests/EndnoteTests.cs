using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// Endnotes share the footnote surface since M4 (the M3 unsupported_feature
/// refusal and its test are gone — see FootnoteTests for the removal note).
/// </summary>
public sealed class EndnoteTests : WordTestBase
{
    [Fact]
    public void Add_endnote_creates_part_with_separator_defaults_and_reference_run()
    {
        var file = CreateDoc(title: "Cited claim");

        var envelope = Edit(file, """[{"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"See the appendix."}}]""");
        Assert.Equal("/endnote[@id=1]", Data(envelope)["ops"]!.AsArray()[0]!["path"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var endnotes = doc.MainDocumentPart!.EndnotesPart!.Endnotes!;
            var all = endnotes.Elements<Endnote>().ToList();
            Assert.Contains(all, e => e.Type?.Value == FootnoteEndnoteValues.Separator && e.Id?.Value == -1);
            Assert.Contains(all, e => e.Type?.Value == FootnoteEndnoteValues.ContinuationSeparator && e.Id?.Value == 0);

            var note = Assert.Single(all, e => e.Type is null || e.Type.Value == FootnoteEndnoteValues.Normal);
            Assert.Equal(1, note.Id!.Value);
            Assert.Contains("See the appendix.", note.InnerText, StringComparison.Ordinal);

            var paragraph = doc.MainDocumentPart.Document!.Body!.Elements<Paragraph>().First();
            var reference = Assert.Single(paragraph.Descendants<EndnoteReference>());
            Assert.Equal(1, reference.Id!.Value);
            var run = Assert.IsType<Run>(reference.Parent);
            Assert.NotNull(run.RunProperties?.VerticalTextAlignment);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Endnote_ids_increment_and_get_resolves_them()
    {
        var file = CreateDoc(title: "Twice noted");
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"first"}},
              {"op":"add","path":"/body/p[2]","type":"endnote","props":{"text":"second"}}
            ]
            """);

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/endnote[@id=2]" })));

        Assert.Equal("endnote", got["type"]!.GetValue<string>());
        Assert.Equal("second", got["properties"]!["text"]!.GetValue<string>());
        Assert.Equal("/body/p[2]", got["properties"]!["anchorPath"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Text_view_appends_markers_and_an_endnote_section()
    {
        var file = CreateDoc(title: "Marked claim");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"the long version"}}]""");

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "text" })));

        Assert.Equal("Marked claim[^e1]", data["lines"]!.AsArray()[0]!["text"]!.GetValue<string>());
        var endnote = Assert.Single(data["endnotes"]!.AsArray())!;
        Assert.Equal(1, endnote["id"]!.GetValue<int>());
        Assert.Equal("the long version", endnote["text"]!.GetValue<string>());
    }

    [Fact]
    public void Html_render_emits_sup_reference_and_endnote_section_after_body()
    {
        var file = CreateDoc(title: "Rendered claim");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"see chapter 9"}}]""");

        var html = Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();

        Assert.Contains("""<sup data-aio-path="/endnote[@id=1]">e1</sup>""", html, StringComparison.Ordinal);
        Assert.Contains("<section class=\"endnotes\">", html, StringComparison.Ordinal);
        Assert.Contains("""<li data-aio-path="/endnote[@id=1]">see chapter 9</li>""", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Footnotes_and_endnotes_coexist()
    {
        var file = CreateDoc(title: "Both kinds");
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"page-level"}},
              {"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"document-level"}}
            ]
            """);

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "text" })));
        Assert.Equal("Both kinds[^1][^e1]", data["lines"]!.AsArray()[0]!["text"]!.GetValue<string>());
        Assert.Single(data["footnotes"]!.AsArray());
        Assert.Single(data["endnotes"]!.AsArray());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_endnote_clears_part_entry_and_reference()
    {
        var file = CreateDoc(title: "Fleeting");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"temp"}}]""");

        Edit(file, """[{"op":"remove","path":"/endnote[@id=1]"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<EndnoteReference>());
            Assert.DoesNotContain(
                doc.MainDocumentPart.EndnotesPart!.Endnotes!.Elements<Endnote>(),
                e => e.Type is null || e.Type.Value == FootnoteEndnoteValues.Normal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Endnote_text_is_required()
    {
        var file = CreateDoc(title: "Silent");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"endnote","props":{}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("text", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Endnote_outside_body_is_invalid_args()
    {
        var file = CreateDoc(title: "Headed");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"H"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/header[1]/p[1]","type":"endnote","props":{"text":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Unknown_endnote_id_is_invalid_path_with_candidates()
    {
        var file = CreateDoc(title: "Missing");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"only"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"remove","path":"/endnote[@id=9]"}]"""));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/endnote[@id=1]", ex.Candidates!);
    }

    [Fact]
    public void Tracked_endnote_add_is_unsupported()
    {
        var file = CreateDoc(title: "Tracked");

        var ex = Assert.Throws<AiofficeException>(() => Edit(
            file,
            """[{"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"x"}}]""",
            new JsonObject { ["track"] = true }));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }
}
