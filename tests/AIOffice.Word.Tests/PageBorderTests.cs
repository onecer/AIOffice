using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class PageBorderTests : WordTestBase
{
    private JsonNode GetSection(string file, string path = "/section[1]") =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    [Fact]
    public void Set_page_border_writes_pgBorders_on_all_four_sides()
    {
        var file = CreateDoc(title: "Bordered");

        Edit(file, """
            [{"op":"set","path":"/section[1]","props":{"pageBorder":{"style":"single","color":"38BDF8","widthPt":1.5,"sides":"all"}}}]
            """);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
            var pgBorders = sectPr.GetFirstChild<PageBorders>()!;
            Assert.NotNull(pgBorders.TopBorder);
            Assert.NotNull(pgBorders.BottomBorder);
            Assert.NotNull(pgBorders.LeftBorder);
            Assert.NotNull(pgBorders.RightBorder);
            Assert.Equal(BorderValues.Single, pgBorders.TopBorder!.Val!.Value);
            Assert.Equal("38BDF8", pgBorders.TopBorder!.Color!.Value);
            Assert.Equal(12u, pgBorders.TopBorder!.Size!.Value); // 1.5pt = 12 eighths
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_section_reports_the_page_border()
    {
        var file = CreateDoc(title: "Bordered");
        Edit(file, """
            [{"op":"set","path":"/section[1]","props":{"pageBorder":{"style":"double","color":"FF0000","widthPt":2,"sides":"all"}}}]
            """);

        var border = GetSection(file)["properties"]!["pageBorder"]!;
        Assert.Equal("double", border["style"]!.GetValue<string>());
        Assert.Equal("FF0000", border["color"]!.GetValue<string>());
        Assert.Equal(2.0, border["widthPt"]!.GetValue<double>());
        Assert.Equal("all", border["sides"]!.GetValue<string>());
    }

    [Fact]
    public void Only_selected_side_is_emitted_and_reported()
    {
        var file = CreateDoc(title: "Top only");

        Edit(file, """
            [{"op":"set","path":"/section[1]","props":{"pageBorder":{"style":"thick","sides":"top"}}}]
            """);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var pgBorders = doc.MainDocumentPart!.Document!.Body!
                .Elements<SectionProperties>().Single()
                .GetFirstChild<PageBorders>()!;
            Assert.NotNull(pgBorders.TopBorder);
            Assert.Null(pgBorders.BottomBorder);
            Assert.Null(pgBorders.LeftBorder);
            Assert.Null(pgBorders.RightBorder);
        }

        Assert.Equal("top", GetSection(file)["properties"]!["pageBorder"]!["sides"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Clearing_with_none_removes_the_page_border()
    {
        var file = CreateDoc(title: "Cleared");
        Edit(file, """
            [{"op":"set","path":"/section[1]","props":{"pageBorder":{"style":"single","sides":"all"}}}]
            """);
        Assert.NotNull(GetSection(file)["properties"]!["pageBorder"]);

        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageBorder":"none"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
            Assert.Null(sectPr.GetFirstChild<PageBorders>());
        }

        Assert.True(GetSection(file)["properties"]!["pageBorder"] is null);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Unknown_style_is_invalid_args()
    {
        var file = CreateDoc(title: "Bad");

        var ex = Assert.Throws<AiofficeException>(() => Edit(file,
            """[{"op":"set","path":"/section[1]","props":{"pageBorder":{"style":"zigzag"}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Invalid_side_is_invalid_args()
    {
        var file = CreateDoc(title: "Bad");

        var ex = Assert.Throws<AiofficeException>(() => Edit(file,
            """[{"op":"set","path":"/section[1]","props":{"pageBorder":{"style":"single","sides":"diagonal"}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }
}
