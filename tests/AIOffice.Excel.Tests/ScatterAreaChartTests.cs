using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M3 chart kinds: scatter (numbers against numbers, two value axes) and
/// area, on the same op surface as bar | line | pie. Raw assertions reopen the
/// file with the OpenXml SDK; every mutating test ends validator-clean.
/// </summary>
public sealed class ScatterAreaChartTests : ExcelTestBase
{
    /// <summary>Numeric X in column A, two numeric series in B/C, header row.</summary>
    private string CreateNumericWorkbook()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1:C5", ("values", new JsonArray(
                new JsonArray("X", "Squared", "Doubled"),
                new JsonArray(1, 1, 2),
                new JsonArray(2, 4, 4),
                new JsonArray(3, 9, 6),
                new JsonArray(4, 16, 8)))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    private static ChartPart SingleChartPart(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return Assert.Single(worksheetPart.DrawingsPart!.Parts.Select(p => p.OpenXmlPart).OfType<ChartPart>());
    }

    [Fact]
    public void Add_scatter_chart_writes_xVal_yVal_and_two_value_axes()
    {
        var file = CreateNumericWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "scatter"), ("dataRange", "A1:C5"), ("anchor", "E2"), ("title", "Growth")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("/Sheet1/chart[1]", Json(envelope)["data"]!["ops"]![0]!["path"]!.GetValue<string>());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var chartSpace = SingleChartPart(document).ChartSpace!;
        var scatter = Assert.Single(chartSpace.Descendants<C.ScatterChart>().ToList());
        var series = scatter.Elements<C.ScatterChartSeries>().ToList();
        Assert.Equal(2, series.Count);

        // Scatter series carry xVal/yVal (not cat/val) with cached numbers.
        var xValues = Assert.Single(series[0].Elements<C.XValues>().ToList());
        Assert.Equal("'Sheet1'!$A$2:$A$5", xValues.Descendants<C.Formula>().Single().Text);
        Assert.Contains("1", xValues.Descendants<C.NumericValue>().Select(v => v.Text));
        var yValues = Assert.Single(series[0].Elements<C.YValues>().ToList());
        Assert.Equal("'Sheet1'!$B$2:$B$5", yValues.Descendants<C.Formula>().Single().Text);
        Assert.Empty(series[0].Descendants<C.CategoryAxisData>());

        // Value-against-value: both plot-area axes are value axes.
        var plotArea = chartSpace.Descendants<C.PlotArea>().Single();
        Assert.Equal(2, plotArea.Elements<C.ValueAxis>().Count());
        Assert.Empty(plotArea.Elements<C.CategoryAxis>());
    }

    [Fact]
    public void Scatter_with_text_x_values_is_invalid_args()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B3", ("values", new JsonArray(
            new JsonArray("Jan", 1),
            new JsonArray("Feb", 2),
            new JsonArray("Mar", 3))))).IsOk);

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "scatter"), ("dataRange", "A1:B3"), ("anchor", "D2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("numeric X", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("line", envelope.Error.Suggestion, StringComparison.Ordinal); // names the workaround
    }

    [Fact]
    public void Add_area_chart_with_two_series_validator_clean()
    {
        var file = CreateNumericWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "area"), ("dataRange", "A1:C5"), ("anchor", "E2")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var chartSpace = SingleChartPart(document).ChartSpace!;
        var area = Assert.Single(chartSpace.Descendants<C.AreaChart>().ToList());
        Assert.Equal(2, area.Elements<C.AreaChartSeries>().Count());

        // Area keeps the category model: cat/val plus category + value axes.
        Assert.NotEmpty(area.Descendants<C.CategoryAxisData>());
        var plotArea = chartSpace.Descendants<C.PlotArea>().Single();
        Assert.Single(plotArea.Elements<C.CategoryAxis>());
        Assert.Single(plotArea.Elements<C.ValueAxis>());
    }

    [Fact]
    public void Scatter_and_area_read_back_through_get_and_structure()
    {
        var file = CreateNumericWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "chart", ("kind", "scatter"), ("dataRange", "A1:B5"), ("anchor", "E2")),
            AddOp("/Sheet1", "chart", ("kind", "area"), ("dataRange", "A1:C5"), ("anchor", "E20"))).IsOk);

        var scatter = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]"))));
        Assert.Equal("scatter", scatter["chartKind"]!.GetValue<string>());
        Assert.Equal("A1:B5", scatter["dataRange"]!.GetValue<string>());
        Assert.Equal(1, scatter["series"]!.GetValue<int>());

        var area = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[2]"))));
        Assert.Equal("area", area["chartKind"]!.GetValue<string>());
        Assert.Equal(2, area["series"]!.GetValue<int>());

        var structure = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        var kinds = structure["sheets"]![0]!["charts"]!.AsArray()
            .Select(c => c!["kind"]!.GetValue<string>())
            .ToList();
        Assert.Equal(["scatter", "area"], kinds);
    }

    [Fact]
    public void Scatter_and_area_survive_a_later_unrelated_edit()
    {
        var file = CreateNumericWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "scatter"), ("dataRange", "A1:B5"), ("anchor", "E2"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/A9", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]"))));
        Assert.Equal("scatter", data["chartKind"]!.GetValue<string>());
    }
}
