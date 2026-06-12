using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class TocTests : WordTestBase
{
    /// <summary>A doc whose body is: H1 "Intro", text, H2 "Details", H3 "Fine print".</summary>
    private string CreateOutlinedDoc()
    {
        var file = CreateDoc();
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"Intro","style":"Heading1"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Some body text."}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Details","style":"Heading2"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Fine print","style":"Heading3"}}
            ]
            """);
        return file;
    }

    private static JsonNode FirstOp(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["data"]!["ops"]!.AsArray()[0]!;

    private static List<string> WarningCodes(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"] is JsonArray warnings
            ? [.. warnings.Select(w => w!["code"]!.GetValue<string>())]
            : [];

    [Fact]
    public void Add_toc_emits_field_plus_hydrated_entries_hyperlinked_to_heading_bookmarks()
    {
        var file = CreateOutlinedDoc();

        var envelope = Edit(
            file,
            """[{"op":"add","path":"/body","type":"toc","props":{"levels":"1-3","title":"Contents","position":"before /body/p[1]"}}]""");

        var op = FirstOp(envelope);
        Assert.Equal("/toc[1]", op["path"]!.GetValue<string>());
        Assert.Equal(3, op["entries"]!.GetValue<int>());
        Assert.Contains("toc_pages_unknown", WarningCodes(envelope));

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sdt = Assert.Single(body.Elements<SdtBlock>());
            Assert.Equal(
                "Table of Contents",
                sdt.SdtProperties!.GetFirstChild<SdtContentDocPartObject>()!.GetFirstChild<DocPartGallery>()!.Val!.Value);

            // The TOC sits before the first heading.
            Assert.True(body.ChildElements.ToList().IndexOf(sdt) <
                body.ChildElements.ToList().IndexOf(body.Elements<Paragraph>().First(p => p.InnerText == "Intro")));

            // Complex field: begin + TOC instruction + separate ... end.
            var fieldTypes = sdt.Descendants<FieldChar>().Select(f => f.FieldCharType!.Value).ToList();
            Assert.Equal(new[] { FieldCharValues.Begin, FieldCharValues.Separate, FieldCharValues.End }, fieldTypes);
            var instruction = Assert.Single(sdt.Descendants<FieldCode>());
            Assert.Contains("TOC \\o \"1-3\" \\h", instruction.Text, StringComparison.Ordinal);

            // Hydrated entries: text per heading, hyperlinked to a real _Toc bookmark on that heading.
            var links = sdt.Descendants<Hyperlink>().ToList();
            Assert.Equal(new[] { "Intro", "Details", "Fine print" }, links.Select(l => l.InnerText));
            foreach (var link in links)
            {
                var anchor = link.Anchor!.Value!;
                Assert.StartsWith("_Toc", anchor, StringComparison.Ordinal);
                var start = Assert.Single(body.Descendants<BookmarkStart>(), b => b.Name?.Value == anchor);
                Assert.Single(body.Descendants<BookmarkEnd>(), e => e.Id?.Value == start.Id?.Value);
                Assert.Equal(link.InnerText, start.Ancestors<Paragraph>().First().InnerText);
            }

            // Entries carry the standard TOC styles and the title its TOCHeading style.
            var entryStyles = sdt.Descendants<Paragraph>()
                .Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value)
                .ToList();
            Assert.Equal(new string?[] { "TOCHeading", "TOC1", "TOC2", "TOC3" }, entryStyles);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_toc_reports_levels_entryCount_and_title()
    {
        var file = CreateOutlinedDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"toc","props":{"levels":"1-2","title":"目录"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/toc[1]" })));

        Assert.Equal("toc", got["type"]!.GetValue<string>());
        Assert.Equal("1-2", got["properties"]!["levels"]!.GetValue<string>());
        Assert.Equal(2, got["properties"]!["entryCount"]!.GetValue<int>());
        Assert.Equal("目录", got["properties"]!["title"]!.GetValue<string>());
    }

    [Fact]
    public void Levels_range_filters_headings()
    {
        var file = CreateOutlinedDoc();

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"toc","props":{"levels":"2-3"}}]""");

        Assert.Equal(2, FirstOp(envelope)["entries"]!.GetValue<int>());
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var sdt = doc.MainDocumentPart!.Document!.Body!.Elements<SdtBlock>().Single();
        // The highest included level (Heading2) maps to TOC1.
        Assert.Equal(
            new string?[] { "TOC1", "TOC2" },
            sdt.Descendants<Paragraph>().Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value));
    }

    [Fact]
    public void Re_adding_refreshes_entries_in_place_with_a_note()
    {
        var file = CreateOutlinedDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"toc","props":{"position":"before /body/p[1]"}}]""");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Appendix","style":"Heading1"}}]""");

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"toc"}]""");

        var op = FirstOp(envelope);
        Assert.Equal(4, op["entries"]!.GetValue<int>());
        Assert.True(op["refreshed"]!.GetValue<bool>());
        Assert.Contains("toc_refreshed", WarningCodes(envelope));

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sdt = Assert.Single(body.Elements<SdtBlock>()); // still exactly one TOC
            Assert.Equal(4, sdt.Descendants<Hyperlink>().Count());
            // The refreshed TOC kept the old spot: still ahead of the first heading.
            Assert.True(body.ChildElements.ToList().IndexOf(sdt) <
                body.ChildElements.ToList().IndexOf(body.Elements<Paragraph>().First(p => p.InnerText == "Intro")));
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_toc_drops_the_block()
    {
        var file = CreateOutlinedDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"toc"}]""");

        Edit(file, """[{"op":"remove","path":"/toc[1]"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Elements<SdtBlock>());
        }

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/toc[1]" })));
        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Structure_view_lists_the_toc()
    {
        var file = CreateOutlinedDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"toc","props":{"levels":"1-3"}}]""");

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));

        var toc = Assert.Single(data["tocs"]!.AsArray())!;
        Assert.Equal("/toc[1]", toc["path"]!.GetValue<string>());
        Assert.Equal("1-3", toc["properties"]!["levels"]!.GetValue<string>());
        Assert.Equal(3, toc["properties"]!["entryCount"]!.GetValue<int>());
    }

    [Fact]
    public void Toc_without_headings_keeps_the_field_with_zero_entries()
    {
        var file = CreateDoc();

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"toc"}]""");

        Assert.Equal(0, FirstOp(envelope)["entries"]!.GetValue<int>());
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sdt = doc.MainDocumentPart!.Document!.Body!.Elements<SdtBlock>().Single();
            Assert.Equal(
                new[] { FieldCharValues.Begin, FieldCharValues.Separate, FieldCharValues.End },
                sdt.Descendants<FieldChar>().Select(f => f.FieldCharType!.Value));
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Bad_levels_or_position_are_invalid_args()
    {
        var file = CreateOutlinedDoc();

        var levels = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"toc","props":{"levels":"3-1"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, levels.Code);

        var position = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"toc","props":{"position":"under /body/p[1]"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, position.Code);
        Assert.Contains("inside", position.Candidates!);
    }

    [Fact]
    public void Tracked_toc_add_is_unsupported()
    {
        var file = CreateOutlinedDoc();

        var ex = Assert.Throws<AiofficeException>(() => Edit(
            file,
            """[{"op":"add","path":"/body","type":"toc"}]""",
            new JsonObject { ["track"] = true }));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }
}
