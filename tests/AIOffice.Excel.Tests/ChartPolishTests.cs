using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.3 chart-polish surface (additive): dataLabels, legend, axisTitles,
/// trendline, errorBars, gridlines and secondaryAxis, accepted BOTH when adding a
/// chart (type:chart props) AND as a set on an existing chart's path. Raw
/// assertions reopen the file with the OpenXml SDK so the chart writer cannot
/// grade its own homework; every mutating test ends OpenXmlValidator-clean.
/// </summary>
public sealed class ChartPolishTests : ExcelTestBase
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

    private static ChartPart SingleChartPart(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return Assert.Single(worksheetPart.DrawingsPart!.Parts.Select(p => p.OpenXmlPart).OfType<ChartPart>());
    }

    private static C.Chart ChartOf(SpreadsheetDocument document) =>
        SingleChartPart(document).ChartSpace!.GetFirstChild<C.Chart>()!;

    // ----- create-time polish -------------------------------------------------

    [Fact]
    public void Create_with_dataLabels_legend_axisTitles_gridlines_is_validator_clean_and_raw_verified()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"), ("title", "Sales"),
            ("dataLabels", new JsonObject { ["show"] = "value", ["position"] = "outEnd" }),
            ("legend", "right"),
            ("axisTitles", new JsonObject { ["category"] = "Month", ["value"] = "Amount" }),
            ("gridlines", new JsonObject { ["major"] = true, ["minor"] = false })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var chart = ChartOf(document);

        // Legend at right.
        Assert.Equal(C.LegendPositionValues.Right, chart.GetFirstChild<C.Legend>()!.GetFirstChild<C.LegendPosition>()!.Val!.Value);

        // Data labels on every series: show value, at outEnd.
        var dLbls = chart.Descendants<C.DataLabels>().ToList();
        Assert.Equal(2, dLbls.Count);
        Assert.All(dLbls, d => Assert.True(d.GetFirstChild<C.ShowValue>()!.Val!.Value));
        Assert.All(dLbls, d => Assert.Equal(C.DataLabelPositionValues.OutsideEnd, d.GetFirstChild<C.DataLabelPosition>()!.Val!.Value));

        // Axis titles on cat + val axes.
        var catAx = chart.Descendants<C.CategoryAxis>().First();
        var valAx = chart.Descendants<C.ValueAxis>().First();
        Assert.Equal("Month", string.Concat(catAx.GetFirstChild<C.Title>()!.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text)));
        Assert.Equal("Amount", string.Concat(valAx.GetFirstChild<C.Title>()!.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text)));

        // Major gridlines present on value axis; no minor.
        Assert.NotNull(valAx.GetFirstChild<C.MajorGridlines>());
        Assert.Null(valAx.GetFirstChild<C.MinorGridlines>());
    }

    [Fact]
    public void Create_with_dataLabels_true_shows_value_at_default_position()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "line"), ("dataRange", "A1:B5"), ("anchor", "E2"),
            ("dataLabels", true)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var dLbls = Assert.Single(ChartOf(document).Descendants<C.DataLabels>().ToList());
        Assert.True(dLbls.GetFirstChild<C.ShowValue>()!.Val!.Value);
        Assert.Null(dLbls.GetFirstChild<C.DataLabelPosition>()); // default position omitted
    }

    [Fact]
    public void Create_with_trendline_and_errorBars_writes_per_series_elements()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "line"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("trendline", "linear"), ("errorBars", "percent")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var chart = ChartOf(document);

        var trendlines = chart.Descendants<C.Trendline>().ToList();
        Assert.Equal(2, trendlines.Count);
        Assert.All(trendlines, t => Assert.Equal(C.TrendlineValues.Linear, t.GetFirstChild<C.TrendlineType>()!.Val!.Value));

        var errBars = chart.Descendants<C.ErrorBars>().ToList();
        Assert.Equal(2, errBars.Count);
        Assert.All(errBars, e => Assert.Equal(C.ErrorValues.Percentage, e.GetFirstChild<C.ErrorBarValueType>()!.Val!.Value));
        Assert.All(errBars, e => Assert.NotNull(e.GetFirstChild<C.ErrorBarValue>()));
    }

    [Fact]
    public void Create_with_movingAverage_trendline_sets_period_two()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "line"), ("dataRange", "A1:B5"), ("anchor", "E2"),
            ("trendline", "movingAverage")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var trendline = Assert.Single(ChartOf(document).Descendants<C.Trendline>().ToList());
        Assert.Equal(C.TrendlineValues.MovingAverage, trendline.GetFirstChild<C.TrendlineType>()!.Val!.Value);
        Assert.Equal(2u, trendline.GetFirstChild<C.Period>()!.Val!.Value);
    }

    [Fact]
    public void Create_with_secondaryAxis_moves_named_series_to_a_second_value_axis()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("secondaryAxis", new JsonArray("Cost"))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var chart = ChartOf(document);

        // Two value axes (primary + secondary) and two category axes.
        Assert.Equal(2, chart.Descendants<C.ValueAxis>().Count());
        Assert.Equal(2, chart.Descendants<C.CategoryAxis>().Count());

        // The secondary value axis crosses at the maximum (right side).
        var secondaryVal = chart.Descendants<C.ValueAxis>()
            .First(a => a.GetFirstChild<C.AxisPosition>()!.Val!.Value == C.AxisPositionValues.Right);
        Assert.Equal(C.CrossesValues.Maximum, secondaryVal.GetFirstChild<C.Crosses>()!.Val!.Value);

        // The 'Cost' series sits in a group that references the secondary axes.
        var groups = chart.GetFirstChild<C.PlotArea>()!.Elements<C.BarChart>().ToList();
        Assert.Equal(2, groups.Count);
        var secondaryGroup = groups.Single(g => g.Elements<C.AxisId>().Any(a => a.Val!.Value == 100003u));
        var movedSeries = Assert.Single(secondaryGroup.Elements<C.BarChartSeries>().ToList());
        Assert.Equal("Cost", movedSeries.GetFirstChild<C.SeriesText>()!.StringReference!.StringCache!
            .Descendants<C.StringPoint>().First().NumericValue!.Text);
    }

    [Fact]
    public void Create_secondaryAxis_with_unknown_series_is_invalid_args_with_candidates()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("secondaryAxis", new JsonArray("Profit"))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("Sales", envelope.Error.Candidates!);
        Assert.Contains("Cost", envelope.Error.Candidates!);
    }

    [Fact]
    public void Create_secondaryAxis_moving_every_series_is_refused()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("secondaryAxis", new JsonArray("Sales", "Cost"))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("primary", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unsupported_legend_position_is_unsupported_feature()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("legend", "diagonal")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("right", envelope.Error.Candidates!);
    }

    // ----- set-time polish on an existing chart -------------------------------

    private string CreateChartWorkbook(string kind = "bar", string dataRange = "A1:C5")
    {
        var file = CreateCategoryWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", kind), ("dataRange", dataRange), ("anchor", "E2"), ("title", "Sales")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    [Fact]
    public void Set_polish_on_existing_chart_path_applies_and_is_validator_clean()
    {
        var file = CreateChartWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/chart[1]",
            ("legend", "bottom"),
            ("dataLabels", new JsonObject { ["show"] = "value" }),
            ("gridlines", new JsonObject { ["major"] = true })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var op = OkData(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("/Sheet1/chart[1]", op["path"]!.GetValue<string>());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var chart = ChartOf(document);
        Assert.Equal(C.LegendPositionValues.Bottom, chart.GetFirstChild<C.Legend>()!.GetFirstChild<C.LegendPosition>()!.Val!.Value);
        Assert.Equal(2, chart.Descendants<C.DataLabels>().Count());
        Assert.NotNull(chart.Descendants<C.ValueAxis>().First().GetFirstChild<C.MajorGridlines>());
    }

    [Fact]
    public void Set_legend_none_removes_the_legend()
    {
        var file = CreateChartWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/chart[1]", ("legend", "right"))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/chart[1]", ("legend", "none")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Null(ChartOf(document).GetFirstChild<C.Legend>());
    }

    [Fact]
    public void Set_dataLabels_false_clears_existing_labels()
    {
        var file = CreateChartWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/chart[1]", ("dataLabels", true))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/chart[1]", ("dataLabels", false)));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Empty(ChartOf(document).Descendants<C.DataLabels>());
    }

    [Fact]
    public void Set_secondaryAxis_on_existing_chart_then_get_reports_it()
    {
        var file = CreateChartWorkbook();

        var setEnvelope = EditOps(file, SetOp("/Sheet1/chart[1]", ("secondaryAxis", new JsonArray("Cost"))));
        Assert.True(setEnvelope.IsOk, setEnvelope.ToJson());
        AssertValidatorClean(file);

        var getEnvelope = Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]")));
        var polish = OkData(getEnvelope)["polish"]!;
        var secondary = polish["secondaryAxis"]!.AsArray();
        Assert.Equal("Cost", secondary[0]!.GetValue<string>());
    }

    [Fact]
    public void Set_polish_mixed_with_non_polish_prop_on_chart_path_is_invalid_args()
    {
        var file = CreateChartWorkbook();

        // A polish prop routes the op to the chart-polish handler; mixing in a
        // non-polish prop (bold) is then invalid_args naming the polish props.
        var envelope = EditOps(file, SetOp("/Sheet1/chart[1]", ("legend", "right"), ("bold", true)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("legend", envelope.Error.Candidates!);
    }

    [Fact]
    public void Set_only_non_polish_prop_on_chart_path_stays_unsupported_feature()
    {
        var file = CreateChartWorkbook();

        // No polish prop at all → the pre-1.3 behavior: a chart is not a cell-set
        // target, so it is unsupported_feature.
        var envelope = EditOps(file, SetOp("/Sheet1/chart[1]", ("bold", true)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
    }

    [Fact]
    public void Set_polish_on_missing_chart_index_is_invalid_path()
    {
        var file = CreateChartWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/chart[5]", ("legend", "right")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1/chart[1]", envelope.Error.Candidates!);
    }

    // ----- get reports polish -------------------------------------------------

    [Fact]
    public void Get_reports_polish_settings()
    {
        var file = CreateCategoryWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"),
            ("legend", "top"),
            ("dataLabels", new JsonObject { ["show"] = "percent" }),
            ("trendline", "linear"),
            ("errorBars", "stdDev"),
            ("gridlines", new JsonObject { ["major"] = true }),
            ("axisTitles", new JsonObject { ["value"] = "Amount" }))).IsOk);

        var envelope = Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]")));
        var polish = OkData(envelope)["polish"]!;
        Assert.Equal("top", polish["legend"]!.GetValue<string>());
        Assert.Equal("percent", polish["dataLabels"]!["show"]!.GetValue<string>());
        Assert.Equal("linear", polish["trendline"]!.GetValue<string>());
        Assert.Equal("stdDev", polish["errorBars"]!.GetValue<string>());
        Assert.True(polish["gridlines"]!["major"]!.GetValue<bool>());
        Assert.Equal("Amount", polish["axisTitles"]!["value"]!.GetValue<string>());
    }

    // ----- polish on multiple chart kinds (validator coverage) ----------------

    [Theory]
    [InlineData("bar", "A1:C5")]
    [InlineData("line", "A1:C5")]
    [InlineData("area", "A1:C5")]
    [InlineData("stackedBar", "A1:C5")]
    public void DataLabels_and_legend_validator_clean_across_kinds(string kind, string dataRange)
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", kind), ("dataRange", dataRange), ("anchor", "E2"),
            ("dataLabels", new JsonObject { ["show"] = "value" }),
            ("legend", "right"),
            ("gridlines", new JsonObject { ["major"] = true })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.NotEmpty(ChartOf(document).Descendants<C.DataLabels>());
    }
}
