using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.1 expanded chart kinds (additive to bar | line | pie | scatter |
/// area): doughnut, radar, bubble, stackedBar, percentStackedBar, stackedArea
/// and combo. Same op surface (type:chart, kind:...). Raw assertions reopen the
/// file with the OpenXml SDK so the chart writer cannot grade its own homework;
/// every mutating test ends OpenXmlValidator-clean.
/// </summary>
public sealed class ExpandedChartKindsTests : ExcelTestBase
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

    [Fact]
    public void Doughnut_writes_doughnutChart_with_holeSize_and_no_axes()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "doughnut"), ("dataRange", "A1:C5"), ("anchor", "E2"), ("title", "Mix")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var space = SingleChartPart(document).ChartSpace!;
        var doughnut = Assert.Single(space.Descendants<C.DoughnutChart>().ToList());
        Assert.Equal(2, doughnut.Elements<C.PieChartSeries>().Count()); // doughnut allows multiple rings
        Assert.NotNull(doughnut.GetFirstChild<C.HoleSize>());
        Assert.Empty(space.Descendants<C.CategoryAxis>());
        Assert.Empty(space.Descendants<C.ValueAxis>());

        // Cached series essentials survive.
        var first = doughnut.Elements<C.PieChartSeries>().First();
        Assert.Equal("'Sheet1'!$B$2:$B$5", first.GetFirstChild<C.Values>()!.NumberReference!.Formula!.Text);
        Assert.Equal(
            ["Jan", "Feb", "Mar", "Apr"],
            first.GetFirstChild<C.CategoryAxisData>()!.StringReference!.StringCache!
                .Descendants<C.StringPoint>().Select(p => p.NumericValue!.Text));
    }

    [Fact]
    public void Radar_writes_radarChart_with_radarStyle_and_category_axes()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "radar"), ("dataRange", "A1:C5"), ("anchor", "E2")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var space = SingleChartPart(document).ChartSpace!;
        var radar = Assert.Single(space.Descendants<C.RadarChart>().ToList());
        Assert.NotNull(radar.GetFirstChild<C.RadarStyle>());
        Assert.Equal(2, radar.Elements<C.RadarChartSeries>().Count());

        var plotArea = space.Descendants<C.PlotArea>().Single();
        Assert.Single(plotArea.Elements<C.CategoryAxis>());
        Assert.Single(plotArea.Elements<C.ValueAxis>());
        Assert.NotEmpty(radar.Descendants<C.CategoryAxisData>());
    }

    [Fact]
    public void StackedBar_writes_bar_grouping_stacked_with_full_overlap()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "stackedBar"), ("dataRange", "A1:C5"), ("anchor", "E2")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var space = SingleChartPart(document).ChartSpace!;
        var bar = Assert.Single(space.Descendants<C.BarChart>().ToList());
        Assert.Equal(C.BarGroupingValues.Stacked, bar.GetFirstChild<C.BarGrouping>()!.Val!.Value);
        Assert.Equal(100, bar.GetFirstChild<C.Overlap>()!.Val!.Value);
        Assert.Equal(2, bar.Elements<C.BarChartSeries>().Count());
    }

    [Fact]
    public void PercentStackedBar_writes_bar_grouping_percentStacked()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "percentStackedBar"), ("dataRange", "A1:C5"), ("anchor", "E2")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var bar = Assert.Single(SingleChartPart(document).ChartSpace!.Descendants<C.BarChart>().ToList());
        Assert.Equal(C.BarGroupingValues.PercentStacked, bar.GetFirstChild<C.BarGrouping>()!.Val!.Value);
        Assert.Equal(100, bar.GetFirstChild<C.Overlap>()!.Val!.Value);
    }

    [Fact]
    public void StackedArea_writes_area_grouping_stacked()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "stackedArea"), ("dataRange", "A1:C5"), ("anchor", "E2")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var area = Assert.Single(SingleChartPart(document).ChartSpace!.Descendants<C.AreaChart>().ToList());
        Assert.Equal(C.GroupingValues.Stacked, area.GetFirstChild<C.Grouping>()!.Val!.Value);
        Assert.Equal(2, area.Elements<C.AreaChartSeries>().Count());
    }

    [Fact]
    public void Combo_writes_a_bar_group_and_a_line_group_sharing_axes()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "combo"), ("dataRange", "A1:C5"), ("anchor", "E2"), ("title", "Sales vs Cost")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var space = SingleChartPart(document).ChartSpace!;
        var bar = Assert.Single(space.Descendants<C.BarChart>().ToList());
        var line = Assert.Single(space.Descendants<C.LineChart>().ToList());

        // First series is the column; the rest go to the line.
        Assert.Single(bar.Elements<C.BarChartSeries>());
        Assert.Single(line.Elements<C.LineChartSeries>());
        Assert.Equal(
            "'Sheet1'!$B$2:$B$5",
            bar.Elements<C.BarChartSeries>().Single().GetFirstChild<C.Values>()!.NumberReference!.Formula!.Text);
        Assert.Equal(
            "'Sheet1'!$C$2:$C$5",
            line.Elements<C.LineChartSeries>().Single().GetFirstChild<C.Values>()!.NumberReference!.Formula!.Text);

        // Both groups reference exactly one shared category/value axis pair.
        var plotArea = space.Descendants<C.PlotArea>().Single();
        Assert.Single(plotArea.Elements<C.CategoryAxis>());
        Assert.Single(plotArea.Elements<C.ValueAxis>());
        var barAxes = bar.Elements<C.AxisId>().Select(a => a.Val!.Value).ToList();
        var lineAxes = line.Elements<C.AxisId>().Select(a => a.Val!.Value).ToList();
        Assert.Equal(barAxes, lineAxes);
    }

    [Fact]
    public void Combo_with_one_series_is_invalid_args()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "combo"), ("dataRange", "A1:B5"), ("anchor", "E2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("two series", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Bubble_writes_xVal_yVal_bubbleSize_and_two_value_axes()
    {
        var file = CreateWorkbook();
        // X, Y, size triple.
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:C5", ("values", new JsonArray(
            new JsonArray("X", "Y", "Size"),
            new JsonArray(1, 10, 4),
            new JsonArray(2, 20, 8),
            new JsonArray(3, 15, 6),
            new JsonArray(4, 25, 9))))).IsOk);

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bubble"), ("dataRange", "A1:C5"), ("anchor", "E2"), ("title", "Spread")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var space = SingleChartPart(document).ChartSpace!;
        var bubble = Assert.Single(space.Descendants<C.BubbleChart>().ToList());
        var series = Assert.Single(bubble.Elements<C.BubbleChartSeries>().ToList());

        var xValues = Assert.Single(series.Elements<C.XValues>().ToList());
        Assert.Equal("'Sheet1'!$A$2:$A$5", xValues.Descendants<C.Formula>().Single().Text);
        var yValues = Assert.Single(series.Elements<C.YValues>().ToList());
        Assert.Equal("'Sheet1'!$B$2:$B$5", yValues.Descendants<C.Formula>().Single().Text);
        var sizes = Assert.Single(series.Elements<C.BubbleSize>().ToList());
        Assert.Equal("'Sheet1'!$C$2:$C$5", sizes.Descendants<C.Formula>().Single().Text);
        Assert.Contains("4", sizes.Descendants<C.NumericValue>().Select(v => v.Text));

        // Value-against-value: both plot-area axes are value axes.
        var plotArea = space.Descendants<C.PlotArea>().Single();
        Assert.Equal(2, plotArea.Elements<C.ValueAxis>().Count());
        Assert.Empty(plotArea.Elements<C.CategoryAxis>());
    }

    [Fact]
    public void Bubble_with_two_yz_pairs_makes_two_series()
    {
        var file = CreateWorkbook();
        // X, Y1, size1, Y2, size2.
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:E4", ("values", new JsonArray(
            new JsonArray("X", "EU", "EUw", "US", "USw"),
            new JsonArray(1, 10, 4, 12, 5),
            new JsonArray(2, 20, 8, 18, 7),
            new JsonArray(3, 15, 6, 22, 9))))).IsOk);

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bubble"), ("dataRange", "A1:E4"), ("anchor", "G2")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var bubble = Assert.Single(SingleChartPart(document).ChartSpace!.Descendants<C.BubbleChart>().ToList());
        var series = bubble.Elements<C.BubbleChartSeries>().ToList();
        Assert.Equal(2, series.Count);
        Assert.Equal("'Sheet1'!$B$2:$B$4", series[0].Elements<C.YValues>().Single().Descendants<C.Formula>().Single().Text);
        Assert.Equal("'Sheet1'!$C$2:$C$4", series[0].Elements<C.BubbleSize>().Single().Descendants<C.Formula>().Single().Text);
        Assert.Equal("'Sheet1'!$D$2:$D$4", series[1].Elements<C.YValues>().Single().Descendants<C.Formula>().Single().Text);
        Assert.Equal("'Sheet1'!$E$2:$E$4", series[1].Elements<C.BubbleSize>().Single().Descendants<C.Formula>().Single().Text);
    }

    [Fact]
    public void Bubble_with_even_column_count_is_invalid_args()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B4", ("values", new JsonArray(
            new JsonArray("X", "Y"),
            new JsonArray(1, 10),
            new JsonArray(2, 20),
            new JsonArray(3, 15))))).IsOk);

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bubble"), ("dataRange", "A1:B4"), ("anchor", "E2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("y/size", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Bubble_with_text_x_is_invalid_args()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:C4", ("values", new JsonArray(
            new JsonArray("X", "Y", "Size"),
            new JsonArray("Jan", 10, 4),
            new JsonArray("Feb", 20, 8),
            new JsonArray("Mar", 15, 6))))).IsOk);

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bubble"), ("dataRange", "A1:C4"), ("anchor", "E2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("numeric X", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Get_and_structure_report_the_new_kinds()
    {
        var file = CreateCategoryWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "chart", ("kind", "doughnut"), ("dataRange", "A1:C5"), ("anchor", "E2")),
            AddOp("/Sheet1", "chart", ("kind", "stackedBar"), ("dataRange", "A1:C5"), ("anchor", "E20")),
            AddOp("/Sheet1", "chart", ("kind", "percentStackedBar"), ("dataRange", "A1:C5"), ("anchor", "E40")),
            AddOp("/Sheet1", "chart", ("kind", "stackedArea"), ("dataRange", "A1:C5"), ("anchor", "E60")),
            AddOp("/Sheet1", "chart", ("kind", "radar"), ("dataRange", "A1:C5"), ("anchor", "E80")),
            AddOp("/Sheet1", "chart", ("kind", "combo"), ("dataRange", "A1:C5"), ("anchor", "E100"))).IsOk);

        var doughnut = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]"))));
        Assert.Equal("doughnut", doughnut["chartKind"]!.GetValue<string>());

        var stackedBar = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[2]"))));
        Assert.Equal("stackedBar", stackedBar["chartKind"]!.GetValue<string>());

        var percent = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[3]"))));
        Assert.Equal("percentStackedBar", percent["chartKind"]!.GetValue<string>());

        var stackedArea = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[4]"))));
        Assert.Equal("stackedArea", stackedArea["chartKind"]!.GetValue<string>());

        var combo = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[6]"))));
        Assert.Equal("combo", combo["chartKind"]!.GetValue<string>());
        Assert.Equal(2, combo["series"]!.GetValue<int>()); // both groups counted
        Assert.Equal("A1:C5", combo["dataRange"]!.GetValue<string>()); // union across both groups

        var kinds = OkData(Handler.Read(Ctx(file, ("view", "structure"))))["sheets"]![0]!["charts"]!.AsArray()
            .Select(c => c!["kind"]!.GetValue<string>()).ToList();
        Assert.Equal(
            ["doughnut", "stackedBar", "percentStackedBar", "stackedArea", "radar", "combo"],
            kinds);
    }

    [Fact]
    public void New_kinds_survive_a_later_unrelated_edit()
    {
        var file = CreateCategoryWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "combo"), ("dataRange", "A1:C5"), ("anchor", "E2"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/A9", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        Assert.Equal("combo", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]"))))["chartKind"]!.GetValue<string>());
    }

    [Fact]
    public void Unsupported_kind_lists_the_expanded_supported_set()
    {
        var file = CreateCategoryWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "surface"), ("dataRange", "A1:C5"), ("anchor", "E2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Equal(
            ["bar", "line", "pie", "scatter", "area",
             "doughnut", "radar", "bubble",
             "stackedBar", "percentStackedBar", "stackedArea", "combo"],
            envelope.Error.Candidates!);
        foreach (var name in new[] { "doughnut", "radar", "bubble", "stackedBar", "combo" })
        {
            Assert.Contains(name, envelope.Error.Suggestion, StringComparison.Ordinal);
        }
    }
}
