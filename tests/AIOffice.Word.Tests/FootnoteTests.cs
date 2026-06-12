using System.Text.Json.Nodes;
using AIOffice.Core;
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

    [Fact]
    public void Endnotes_are_unsupported_naming_footnotes()
    {
        var file = CreateDoc(title: "Ends");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"endnote","props":{"text":"x"}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("footnote", ex.Suggestion, StringComparison.Ordinal);
    }

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
}
