using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// Multi-column section layout: equal-width and explicit-width <c>w:cols</c>, the
/// column gap, and hard column breaks (<c>w:br w:type="column"</c>). Every
/// mutation reopens to verify the column setup and stays validator-clean.
/// </summary>
public sealed class MultiColumnTests : WordTestBase
{
    private static Columns? Cols(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.GetFirstChild<SectionProperties>()?.GetFirstChild<Columns>();
    }

    [Fact]
    public void Two_equal_columns_set_column_count_and_equal_width()
    {
        var file = CreateDoc(title: "Cols");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"columns":2,"columnGap":"1cm"}}]""");

        var cols = Cols(file);
        Assert.NotNull(cols);
        Assert.Equal((short)2, cols!.ColumnCount?.Value);
        Assert.NotEqual(DocumentFormat.OpenXml.OnOffValue.FromBoolean(false), cols.EqualWidth);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_section_reports_column_count_and_gap()
    {
        var file = CreateDoc(title: "Cols");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"columns":3,"columnGap":"1cm"}}]""");

        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/section[1]" })))["properties"]!;
        Assert.Equal(3, properties["columns"]!.GetValue<int>());
        Assert.Equal(1.0, properties["columnGapCm"]!.GetValue<double>(), 1);
    }

    [Fact]
    public void Explicit_widths_create_one_column_each_with_widths()
    {
        var file = CreateDoc(title: "Cols");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"columns":[{"width":"6cm"},{"width":"4cm"}]}}]""");

        var cols = Cols(file);
        Assert.NotNull(cols);
        Assert.Equal(2, cols!.Elements<Column>().Count());
        Assert.Equal(DocumentFormat.OpenXml.OnOffValue.FromBoolean(false), cols.EqualWidth);

        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/section[1]" })))["properties"]!;
        var widths = properties["columnWidthsCm"]!.AsArray();
        Assert.Equal(6.0, widths[0]!.GetValue<double>(), 1);
        Assert.Equal(4.0, widths[1]!.GetValue<double>(), 1);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Single_column_is_the_default_when_no_cols_present()
    {
        var file = CreateDoc(title: "Cols");
        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/section[1]" })))["properties"]!;
        Assert.Equal(1, properties["columns"]!.GetValue<int>());
    }

    [Fact]
    public void Column_gap_alone_retunes_the_running_layout()
    {
        var file = CreateDoc(title: "Cols");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"columns":2}}]""");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"columnGap":"2cm"}}]""");

        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/section[1]" })))["properties"]!;
        Assert.Equal(2, properties["columns"]!.GetValue<int>()); // count preserved
        Assert.Equal(2.0, properties["columnGapCm"]!.GetValue<double>(), 1);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Column_break_inserts_a_column_type_break()
    {
        var file = CreateDoc(title: "Cols");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"first column"}}]""");
        var result = Data(Edit(file, """[{"op":"add","path":"/body/p[2]","type":"columnBreak"}]"""));

        Assert.Equal("columnBreak", result["ops"]![0]!["type"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var breaks = doc.MainDocumentPart!.Document!.Body!.Descendants<Break>()
                .Where(b => b.Type?.Value == BreakValues.Column)
                .ToList();
            Assert.Single(breaks);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Column_break_before_lands_at_the_paragraph_start()
    {
        var file = CreateDoc(title: "Cols");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"content"}}]""");
        Edit(file, """[{"op":"add","path":"/body/p[2]","type":"columnBreak","position":"before"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ElementAt(1);
            // The first non-pPr child is the break run.
            var firstContent = paragraph.ChildElements.First(c => c is not ParagraphProperties);
            Assert.Contains(firstContent.Descendants<Break>(), b => b.Type?.Value == BreakValues.Column);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Invalid_column_count_is_a_clear_error()
    {
        var file = CreateDoc(title: "Cols");
        var ex = Assert.Throws<AIOffice.Core.AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/section[1]","props":{"columns":0}}]"""));

        Assert.Equal(AIOffice.Core.ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotEmpty(ex.Suggestion);
    }

    [Fact]
    public void Columns_and_page_setup_coexist_in_one_set()
    {
        var file = CreateDoc(title: "Cols");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"A4","orientation":"landscape","columns":2,"columnGap":"1.5cm"}}]""");

        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/section[1]" })))["properties"]!;
        Assert.Equal("landscape", properties["orientation"]!.GetValue<string>());
        Assert.Equal(2, properties["columns"]!.GetValue<int>());
        AssertValidatesClean(file);
    }
}
