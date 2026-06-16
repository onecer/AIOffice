using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.8 <c>seriesColors</c> brand-palette chart prop (additive): an array of
/// 6-hex RGB strings, one per series in dataRange order (a short list cycles),
/// accepted BOTH when adding a chart (type:chart props) AND as a set on an existing
/// chart's path. Raw assertions reopen the file with the OpenXml SDK so the chart
/// writer cannot grade its own homework; every mutating test ends validator-clean.
/// </summary>
public sealed class SeriesColorTests : ExcelTestBase
{
    /// <summary>Header data Month/Sales/Cost in A1:C5 (categories + two numeric series).</summary>
    private string CreateCategoryWorkbook()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1:C5", ("values", new JsonArray(
                new JsonArray("Month", "Sales", "Cost"),
                new JsonArray("Jan", 10, 4),
                new JsonArray("Feb", 20, 8),
                new JsonArray("Mar", 15, 6),
                new JsonArray("Apr", 25, 9)))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    /// <summary>Region/Share in A1:B4 (one pie series over three slices).</summary>
    private string CreatePieWorkbook()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1:B4", ("values", new JsonArray(
                new JsonArray("Region", "Share"),
                new JsonArray("NA", 40),
                new JsonArray("EU", 35),
                new JsonArray("APAC", 25)))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    private static C.Chart ChartOf(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var chartPart = worksheetPart.DrawingsPart!.Parts.Select(p => p.OpenXmlPart).OfType<ChartPart>().Single();
        return chartPart.ChartSpace!.GetFirstChild<C.Chart>()!;
    }

    private static List<string> SeriesFillColors(C.Chart chart) =>
        chart.Descendants<C.BarChartSeries>()
            .Select(s => s.GetFirstChild<C.ChartShapeProperties>()
                ?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value)
            .Where(v => v is not null)
            .Select(v => v!)
            .ToList();

    // ----- create-time -------------------------------------------------------

    [Fact]
    public void Create_bar_with_seriesColors_fills_each_series_and_is_validator_clean()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("seriesColors", new JsonArray("2E5AAC", "#E8743B"))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var colors = SeriesFillColors(ChartOf(document));
        Assert.Equal(new[] { "2E5AAC", "E8743B" }, colors);
    }

    [Fact]
    public void Short_seriesColors_list_cycles_across_series()
    {
        var file = CreateCategoryWorkbook();

        // Two series, one color → the single color repeats.
        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("seriesColors", new JsonArray("2E5AAC"))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var colors = SeriesFillColors(ChartOf(document));
        Assert.Equal(new[] { "2E5AAC", "2E5AAC" }, colors);
    }

    [Fact]
    public void A_leading_hash_is_accepted_and_normalized()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("seriesColors", new JsonArray("#2e5aac", "#e8743b"))));

        Assert.True(envelope.IsOk, envelope.ToJson());

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var colors = SeriesFillColors(ChartOf(document));
        Assert.Equal(new[] { "2E5AAC", "E8743B" }, colors);
    }

    [Fact]
    public void Create_line_with_seriesColors_tints_the_series_line()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "line"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("seriesColors", new JsonArray("2E5AAC", "E8743B"))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var lineColors = ChartOf(document).Descendants<C.LineChartSeries>()
            .Select(s => s.GetFirstChild<C.ChartShapeProperties>()
                ?.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()
                ?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value)
            .ToList();
        Assert.Equal(new[] { "2E5AAC", "E8743B" }, lineColors);
    }

    [Fact]
    public void Create_pie_with_seriesColors_colors_each_slice_via_data_points()
    {
        var file = CreatePieWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "pie"), ("dataRange", "A1:B4"), ("anchor", "D2"),
            ("seriesColors", new JsonArray("2E5AAC", "E8743B", "6FB07F"))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var points = ChartOf(document).Descendants<C.PieChartSeries>().Single()
            .Elements<C.DataPoint>()
            .Select(p => p.GetFirstChild<C.ChartShapeProperties>()
                ?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value)
            .ToList();
        Assert.Equal(new[] { "2E5AAC", "E8743B", "6FB07F" }, points);
    }

    // ----- set-time + round-trip ---------------------------------------------

    [Fact]
    public void Set_seriesColors_on_an_existing_chart_is_applied_and_round_trips_via_get()
    {
        var file = CreateCategoryWorkbook();
        var added = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2")));
        Assert.True(added.IsOk, added.ToJson());

        var envelope = EditOps(file, SetOp("/Sheet1/chart[1]",
            ("seriesColors", new JsonArray("FF0000", "00FF00"))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            Assert.Equal(new[] { "FF0000", "00FF00" }, SeriesFillColors(ChartOf(document)));
        }

        // get reports the brand palette back (round-trip through DescribePolish).
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]"))));
        var reported = data["polish"]!["seriesColors"]!.AsArray().Select(n => n!.GetValue<string>()).ToArray();
        Assert.Equal(new[] { "FF0000", "00FF00" }, reported);
    }

    [Fact]
    public void A_later_set_overrides_an_earlier_seriesColors()
    {
        var file = CreateCategoryWorkbook();
        EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("seriesColors", new JsonArray("2E5AAC", "E8743B"))));

        var envelope = EditOps(file, SetOp("/Sheet1/chart[1]",
            ("seriesColors", new JsonArray("111111", "222222"))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        // No duplicate spPr — the override replaced rather than stacked.
        var firstSeries = ChartOf(document).Descendants<C.BarChartSeries>().First();
        Assert.Single(firstSeries.Elements<C.ChartShapeProperties>());
        Assert.Equal(new[] { "111111", "222222" }, SeriesFillColors(ChartOf(document)));
    }

    // ----- validation --------------------------------------------------------

    [Fact]
    public void A_non_hex_seriesColor_is_invalid_args()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("seriesColors", new JsonArray("notacolor"))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void An_empty_seriesColors_array_is_invalid_args()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("seriesColors", new JsonArray())));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, Json(envelope)["error"]!["code"]!.GetValue<string>());
    }

    [Fact]
    public void SeriesColors_combine_with_other_polish_props_in_one_op()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("legend", "right"),
            ("dataLabels", true),
            ("seriesColors", new JsonArray("2E5AAC", "E8743B"))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var chart = ChartOf(document);
        Assert.Equal(new[] { "2E5AAC", "E8743B" }, SeriesFillColors(chart));
        Assert.Equal(C.LegendPositionValues.Right, chart.GetFirstChild<C.Legend>()!.GetFirstChild<C.LegendPosition>()!.Val!.Value);
        Assert.Equal(2, chart.Descendants<C.DataLabels>().Count());
    }
}
