using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// Chart-polish props (v1.3.0, additive): dataLabels, legend, axisTitles,
/// trendline, errorBars, gridlines, secondaryAxis — accepted at create and as
/// in-place set edits, reopen-asserted and validator-clean.
/// </summary>
public sealed class ChartPolishTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() => TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private ErrorBody EditFail(string code, params EditOp[] ops) =>
        TestEnv.AssertFail(_handler.Edit(_ws.Ctx("deck.pptx"), ops), code);

    private JsonObject Get(string path) => TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));

    private static JsonObject ChartProps(string kind, int seriesCount = 2, params (string Key, JsonNode? Value)[] extra)
    {
        var series = new JsonArray();
        for (var i = 0; i < seriesCount; i++)
        {
            series.Add(new JsonObject
            {
                ["name"] = $"Series {(char)('A' + i)}",
                ["values"] = new JsonArray(10 + (i * 5), 20 + (i * 5), 15 + (i * 5), 30 + (i * 5)),
            });
        }

        var props = TestEnv.Props(
            ("kind", kind),
            ("categories", new JsonArray("Q1", "Q2", "Q3", "Q4")),
            ("series", series),
            ("x", "2cm"), ("y", "3cm"), ("w", "20cm"), ("h", "12cm"));
        foreach (var (key, value) in extra)
        {
            props[key] = value;
        }

        return props;
    }

    private C.ChartSpace SingleChartSpace()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        return doc.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace!;
    }

    private void AddBar(params (string Key, JsonNode? Value)[] polish)
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar", extra: polish)));
    }

    [Fact]
    public void DataLabels_True_AtCreate_WritesDLblsAndValidates()
    {
        AddBar(("dataLabels", true));

        var dLbls = SingleChartSpace().Descendants<C.DataLabels>().Single();
        Assert.True(dLbls.GetFirstChild<C.ShowValue>()!.Val!.Value);
        Assert.False(dLbls.GetFirstChild<C.ShowPercent>()!.Val!.Value);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void DataLabels_ObjectWithShowPercentAndPosition_RoundTrips()
    {
        AddBar(("dataLabels", new JsonObject { ["show"] = "percent", ["position"] = "center" }));

        var dLbls = SingleChartSpace().Descendants<C.DataLabels>().Single();
        Assert.True(dLbls.GetFirstChild<C.ShowPercent>()!.Val!.Value);
        Assert.Equal(C.DataLabelPositionValues.Center, dLbls.GetFirstChild<C.DataLabelPosition>()!.Val!.Value);

        var polish = Get("/slide[1]/chart[1]")["polish"]!["dataLabels"]!.AsObject();
        Assert.Equal("percent", polish["show"]!.GetValue<string>());
        Assert.Equal("center", polish["position"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Legend_Bottom_WritesLegendInChartOrderAndValidates()
    {
        AddBar(("legend", "bottom"));

        var chart = SingleChartSpace().GetFirstChild<C.Chart>()!;
        var legend = chart.GetFirstChild<C.Legend>()!;
        Assert.Equal(C.LegendPositionValues.Bottom, legend.GetFirstChild<C.LegendPosition>()!.Val!.Value);

        // c:legend must precede c:plotVisOnly (schema order).
        var children = chart.ChildElements.Select(e => e.LocalName).ToList();
        Assert.True(children.IndexOf("legend") < children.IndexOf("plotVisOnly"));
        Assert.Equal("bottom", Get("/slide[1]/chart[1]")["polish"]!["legend"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Legend_None_RemovesAnyLegend()
    {
        AddBar(("legend", "right"));
        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(("legend", "none"))));

        Assert.Empty(SingleChartSpace().Descendants<C.Legend>());
        Assert.Null(Get("/slide[1]/chart[1]")["polish"]!["legend"]);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AxisTitles_OnCategoryAndValueAxes_RoundTrip()
    {
        AddBar(("axisTitles", new JsonObject { ["category"] = "Quarter", ["value"] = "Revenue" }));

        var space = SingleChartSpace();
        Assert.Contains("Quarter", space.Descendants<C.CategoryAxis>().Single().GetFirstChild<C.Title>()!.InnerText, StringComparison.Ordinal);
        Assert.Contains("Revenue", space.Descendants<C.ValueAxis>().Single().GetFirstChild<C.Title>()!.InnerText, StringComparison.Ordinal);

        var titles = Get("/slide[1]/chart[1]")["polish"]!["axisTitles"]!.AsObject();
        Assert.Equal("Quarter", titles["category"]!.GetValue<string>());
        Assert.Equal("Revenue", titles["value"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Trendline_Linear_AddsOnePerValueSeries()
    {
        AddBar(("trendline", "linear"));

        var trendlines = SingleChartSpace().Descendants<C.Trendline>().ToList();
        Assert.Equal(2, trendlines.Count); // one per series
        Assert.Equal(C.TrendlineValues.Linear, trendlines[0].GetFirstChild<C.TrendlineType>()!.Val!.Value);
        Assert.Equal("linear", Get("/slide[1]/chart[1]")["polish"]!["trendline"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Trendline_MovingAverage_CarriesPeriodTwo()
    {
        AddBar(("trendline", "movingAverage"));

        var trendline = SingleChartSpace().Descendants<C.Trendline>().First();
        Assert.Equal(C.TrendlineValues.MovingAverage, trendline.GetFirstChild<C.TrendlineType>()!.Val!.Value);
        Assert.Equal(2u, trendline.GetFirstChild<C.Period>()!.Val!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ErrorBars_StdErr_WritesErrBarsPerSeries()
    {
        AddBar(("errorBars", "stdErr"));

        var errBars = SingleChartSpace().Descendants<C.ErrorBars>().ToList();
        Assert.Equal(2, errBars.Count);
        Assert.Equal(C.ErrorValues.StandardError, errBars[0].GetFirstChild<C.ErrorBarValueType>()!.Val!.Value);
        Assert.Equal("stdErr", Get("/slide[1]/chart[1]")["polish"]!["errorBars"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Gridlines_MajorOnMinorOff_WritesMajorGridlinesOnly()
    {
        AddBar(("gridlines", new JsonObject { ["major"] = true, ["minor"] = false }));

        var valueAxis = SingleChartSpace().Descendants<C.ValueAxis>().Single();
        Assert.NotNull(valueAxis.GetFirstChild<C.MajorGridlines>());
        Assert.Null(valueAxis.GetFirstChild<C.MinorGridlines>());

        var grid = Get("/slide[1]/chart[1]")["polish"]!["gridlines"]!.AsObject();
        Assert.True(grid["major"]!.GetValue<bool>());
        Assert.False(grid["minor"]!.GetValue<bool>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SecondaryAxis_MovesNamedSeriesToASecondValueAxis()
    {
        AddBar(("secondaryAxis", new JsonArray("Series B")));

        var space = SingleChartSpace();
        // The chart now has two value axes (primary + secondary).
        Assert.Equal(2, space.Descendants<C.ValueAxis>().Count());
        // Series B rides on a line group (the secondary axis group).
        Assert.Contains(
            space.Descendants<C.LineChartSeries>(),
            s => s.Descendants<C.NumericValue>().Any(n => n.Text == "Series B"));

        var secondary = Get("/slide[1]/chart[1]")["polish"]!["secondaryAxis"]!.AsArray();
        Assert.Contains(secondary, n => n!.GetValue<string>() == "Series B");
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Polish_SetInPlace_OnExistingLineChart_Validates()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("line", seriesCount: 3)));
        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(
            ("dataLabels", true),
            ("legend", "top"),
            ("trendline", "exponential"),
            ("gridlines", new JsonObject { ["major"] = true }))));

        var space = SingleChartSpace();
        Assert.NotNull(space.Descendants<C.DataLabels>().FirstOrDefault());
        Assert.NotNull(space.Descendants<C.Legend>().FirstOrDefault());
        Assert.Equal(3, space.Descendants<C.Trendline>().Count());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Polish_AllSevenAtOnce_AtCreate_Validates()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar",
            extra:
            [
                ("dataLabels", new JsonObject { ["show"] = "value", ["position"] = "outEnd" }),
                ("legend", "right"),
                ("axisTitles", new JsonObject { ["category"] = "Q", ["value"] = "USD" }),
                ("trendline", "linear"),
                ("errorBars", "stdDev"),
                ("gridlines", new JsonObject { ["major"] = true, ["minor"] = true }),
                ("secondaryAxis", new JsonArray("Series B")),
            ])));

        TestEnv.AssertValid(_ws, "deck.pptx");
        var polish = Get("/slide[1]/chart[1]")["polish"]!.AsObject();
        Assert.NotNull(polish["dataLabels"]);
        Assert.Equal("right", polish["legend"]!.GetValue<string>());
        Assert.Equal("linear", polish["trendline"]!.GetValue<string>());
        Assert.Equal("stdDev", polish["errorBars"]!.GetValue<string>());
    }

    [Fact]
    public void BadLegend_IsTypedUnsupported()
    {
        AddBar(("legend", "right"));
        var error = EditFail(ErrorCodes.UnsupportedFeature,
            TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(("legend", "diagonal"))));
        Assert.Contains("right", error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void BadTrendline_IsTypedUnsupported()
    {
        AddBar();
        EditFail(ErrorCodes.UnsupportedFeature,
            TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(("trendline", "logarithmic"))));
    }

    [Fact]
    public void BadDataLabelPosition_IsTypedUnsupported()
    {
        AddBar();
        EditFail(ErrorCodes.UnsupportedFeature, TestEnv.Op("set", "/slide[1]/chart[1]",
            props: TestEnv.Props(("dataLabels", new JsonObject { ["position"] = "floating" }))));
    }

    [Fact]
    public void SvgRender_BarWithLegendAndLabels_ShowsLegendStripAndValues()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar",
            extra: [("legend", "bottom"), ("dataLabels", true)])));

        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data.ToJsonString();
        Assert.Contains("aio-chart-legend", svg, StringComparison.Ordinal);
        Assert.Contains("aio-chart-label", svg, StringComparison.Ordinal);
    }
}
