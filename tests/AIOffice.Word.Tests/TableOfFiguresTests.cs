using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class TableOfFiguresTests : WordTestBase
{
    private static JsonNode FirstOp(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["data"]!["ops"]!.AsArray()[0]!;

    private static List<string> WarningCodes(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"] is JsonArray warnings
            ? [.. warnings.Select(w => w!["code"]!.GetValue<string>())]
            : [];

    /// <summary>A doc with two Figure captions and one Table caption.</summary>
    private string CaptionedDoc()
    {
        var file = CreateDoc(title: "Figures");
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"[chart 1]"}},
              {"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"Revenue","position":"after"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"[chart 2]"}},
              {"op":"add","path":"/body/p[3]","type":"caption","props":{"label":"Figure","text":"Costs","position":"after"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"[grid]"}},
              {"op":"add","path":"/body/p[5]","type":"caption","props":{"label":"Table","text":"Totals","position":"after"}}
            ]
            """);
        return file;
    }

    [Fact]
    public void Add_table_of_figures_lists_figure_captions_with_field_and_warning()
    {
        var file = CaptionedDoc();

        var envelope = Edit(file,
            """[{"op":"add","path":"/body","type":"tableOfFigures","props":{"label":"Figure","title":"Figures","position":"before /body/p[1]"}}]""");

        var op = FirstOp(envelope);
        Assert.Equal("/tableOfFigures[1]", op["path"]!.GetValue<string>());
        Assert.Equal("Figure", op["label"]!.GetValue<string>());
        Assert.Equal(2, op["entries"]!.GetValue<int>());
        Assert.Contains("figures_cached", WarningCodes(envelope));

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sdt = Assert.Single(body.Elements<SdtBlock>(), s =>
                s.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value
                    == "Table of Figures");

            // The TOF field uses the figure caption switch.
            var instruction = Assert.Single(sdt.Descendants<FieldCode>());
            Assert.Contains("TOC \\h \\z \\c \"Figure\"", instruction.Text, StringComparison.Ordinal);

            // begin + separate + end frame, exactly like the TOC.
            var fieldTypes = sdt.Descendants<FieldChar>().Select(f => f.FieldCharType!.Value).ToList();
            Assert.Equal(new[] { FieldCharValues.Begin, FieldCharValues.Separate, FieldCharValues.End }, fieldTypes);

            // One hyperlinked entry per Figure caption (the Table caption is excluded).
            var links = sdt.Descendants<Hyperlink>().ToList();
            Assert.Equal(2, links.Count);
            Assert.Contains("Figure 1: Revenue", links[0].InnerText, StringComparison.Ordinal);
            Assert.Contains("Figure 2: Costs", links[1].InnerText, StringComparison.Ordinal);

            // Each entry hyperlinks to a _Toc bookmark living on its caption paragraph.
            foreach (var link in links)
            {
                var anchor = link.Anchor!.Value!;
                Assert.StartsWith("_Toc", anchor, StringComparison.Ordinal);
                var start = Assert.Single(body.Descendants<BookmarkStart>(), b => b.Name?.Value == anchor);
                Assert.Equal("Caption", start.Ancestors<Paragraph>().First().ParagraphProperties?.ParagraphStyleId?.Val?.Value);
            }
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_table_of_figures_reports_label_and_entry_count()
    {
        var file = CaptionedDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"tableOfFigures","props":{"label":"Figure"}}]""");

        var envelope = Handler.Get(Ctx(file, new JsonObject { ["path"] = "/tableOfFigures[1]" }));
        var data = Data(envelope);

        Assert.Equal("/tableOfFigures[1]", data["path"]!.GetValue<string>());
        Assert.Equal("tableOfFigures", data["type"]!.GetValue<string>());
        Assert.Equal("Figure", data["properties"]!["label"]!.GetValue<string>());
        Assert.Equal(2, data["properties"]!["entryCount"]!.GetValue<int>());
    }

    [Fact]
    public void Structure_view_lists_the_table_of_figures()
    {
        var file = CaptionedDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"tableOfFigures","props":{"label":"Figure"}}]""");

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));
        var tofs = data["tablesOfFigures"]!.AsArray();

        Assert.Single(tofs);
        Assert.Equal("/tableOfFigures[1]", tofs[0]!["path"]!.GetValue<string>());
        Assert.Equal("Figure", tofs[0]!["properties"]!["label"]!.GetValue<string>());
    }

    [Fact]
    public void Rerunning_refreshes_the_table_in_place()
    {
        var file = CaptionedDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"tableOfFigures","props":{"label":"Figure"}}]""");

        // Add a third figure caption, then refresh.
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"[chart 3]"}},
              {"op":"add","path":"/body/p[8]","type":"caption","props":{"label":"Figure","text":"Margin","position":"after"}}
            ]
            """);
        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"tableOfFigures","props":{"label":"Figure"}}]""");

        var op = FirstOp(envelope);
        Assert.True(op["refreshed"]!.GetValue<bool>());
        Assert.Equal(3, op["entries"]!.GetValue<int>());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        // Only one TOF block survives the refresh.
        Assert.Single(doc.MainDocumentPart!.Document!.Body!.Elements<SdtBlock>(), s =>
            s.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value
                == "Table of Figures");
        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_drops_the_table_but_keeps_the_captions()
    {
        var file = CaptionedDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"tableOfFigures","props":{"label":"Figure"}}]""");

        Edit(file, """[{"op":"remove","path":"/tableOfFigures[1]"}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.DoesNotContain(body.Elements<SdtBlock>(), s =>
            s.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value
                == "Table of Figures");
        // Captions survive.
        Assert.Equal(3, body.Elements<Paragraph>()
            .Count(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Caption"));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Unknown_label_is_invalid_args()
    {
        var file = CaptionedDoc();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"tableOfFigures","props":{"label":"Sidebar"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }
}
