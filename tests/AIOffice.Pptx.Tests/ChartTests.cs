using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>Native charts: graphicFrame + ChartPart with reference caches and an embedded data workbook (M4).</summary>
public sealed class ChartTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
    }

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private static JsonObject ChartProps(string kind, string? title = null, int seriesCount = 2)
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
        if (title is not null)
        {
            props["title"] = title;
        }

        return props;
    }

    private ChartPart SingleChartPart()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        return doc.PresentationPart!.SlideParts.Single().ChartParts.Single();
    }

    [Fact]
    public void AddBarChart_WritesReferenceCachedChartMl()
    {
        Create();
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar", "Revenue")));
        Assert.StartsWith("/slide[1]/shape[@id=", data["results"]![0]!["target"]!.GetValue<string>(), StringComparison.Ordinal);

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var frame = slidePart.Slide!.Descendants<P.GraphicFrame>().Single();
            Assert.Equal(720_000L, frame.Transform!.Offset!.X!.Value); // 2cm

            var chartPart = slidePart.ChartParts.Single();
            var barChart = chartPart.ChartSpace!.Descendants<C.BarChart>().Single();
            var series = barChart.Elements<C.BarChartSeries>().ToList();
            Assert.Equal(2, series.Count);

            // Categories live in a string reference cache pointing at the embedded Sheet1.
            var categoryRef = series[0].Descendants<C.StringReference>()
                .Single(r => r.Parent is C.CategoryAxisData);
            Assert.Equal("Sheet1!$A$2:$A$5", categoryRef.GetFirstChild<C.Formula>()!.Text);
            var stringCache = categoryRef.GetFirstChild<C.StringCache>()!;
            Assert.Equal(4u, stringCache.GetFirstChild<C.PointCount>()!.Val!.Value);
            Assert.Equal("Q1", stringCache.Elements<C.StringPoint>().First().NumericValue!.Text);

            var valuesRef = series[0].Descendants<C.NumberReference>().Single();
            Assert.Equal("Sheet1!$B$2:$B$5", valuesRef.GetFirstChild<C.Formula>()!.Text);
            var numberCache = valuesRef.GetFirstChild<C.NumberingCache>()!;
            Assert.Equal(4u, numberCache.GetFirstChild<C.PointCount>()!.Val!.Value);
            Assert.Equal("10", numberCache.Elements<C.NumericPoint>().First().NumericValue!.Text);

            // The second series points at its own column.
            Assert.Equal(
                "Sheet1!$C$2:$C$5",
                series[1].Descendants<C.NumberReference>().Single().GetFirstChild<C.Formula>()!.Text);

            Assert.NotNull(chartPart.ChartSpace.Descendants<C.CategoryAxis>().SingleOrDefault());
            Assert.NotNull(chartPart.ChartSpace.Descendants<C.ValueAxis>().SingleOrDefault());
            Assert.Contains("Revenue", chartPart.ChartSpace.Descendants<C.Title>().Single().InnerText, StringComparison.Ordinal);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddLineChart_WritesLineGroupAndAxes()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("line", seriesCount: 3)));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var chartPart = doc.PresentationPart!.SlideParts.Single().ChartParts.Single();
            var lineChart = chartPart.ChartSpace!.Descendants<C.LineChart>().Single();
            Assert.Equal(3, lineChart.Elements<C.LineChartSeries>().Count());
            Assert.NotNull(chartPart.ChartSpace.Descendants<C.CategoryAxis>().SingleOrDefault());
            Assert.NotNull(chartPart.ChartSpace.Descendants<C.ValueAxis>().SingleOrDefault());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddPieChart_WritesSingleSeriesNoAxes()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("pie", seriesCount: 1)));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var chartPart = doc.PresentationPart!.SlideParts.Single().ChartParts.Single();
            var pieChart = chartPart.ChartSpace!.Descendants<C.PieChart>().Single();
            Assert.Single(pieChart.Elements<C.PieChartSeries>());
            Assert.Empty(chartPart.ChartSpace.Descendants<C.CategoryAxis>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddChart_AttachesNoWarning_DataIsEmbedded()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar"))]);

        TestEnv.AssertOk(envelope);
        Assert.Null(envelope.Meta.Warnings); // the M3 chart_data_not_editable warning is gone
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_ChartPath_ReportsDataAndEditable()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar", "Revenue")));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/chart[1]"))));
        Assert.Equal("/slide[1]/chart[1]", detail["path"]!.GetValue<string>());
        Assert.StartsWith("/slide[1]/shape[@id=", detail["shapePath"]!.GetValue<string>(), StringComparison.Ordinal);
        Assert.Equal("chart", detail["kind"]!.GetValue<string>());
        Assert.Equal("bar", detail["chartKind"]!.GetValue<string>());
        Assert.Equal("Revenue", detail["title"]!.GetValue<string>());
        Assert.True(detail["dataEditable"]!.GetValue<bool>());

        var categories = detail["categories"]!.AsArray();
        Assert.Equal(["Q1", "Q2", "Q3", "Q4"], categories.Select(c => c!.GetValue<string>()));

        var series = detail["series"]!.AsArray();
        Assert.Equal(2, series.Count);
        Assert.Equal("Series A", series[0]!["name"]!.GetValue<string>());
        Assert.Equal(10, series[0]!["values"]![0]!.GetValue<double>());
        Assert.Equal(2, detail["x"]!.GetValue<double>());
        Assert.Equal(12, detail["h"]!.GetValue<double>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_ShapeIdPath_IncludesChartSummary()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("pie", "Share", seriesCount: 1)));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", shapePath))));
        Assert.Equal("graphicFrame", detail["kind"]!.GetValue<string>());
        Assert.Equal("/slide[1]/chart[1]", detail["chartPath"]!.GetValue<string>());

        var chart = detail["chart"]!.AsObject();
        Assert.Equal("pie", chart["kind"]!.GetValue<string>());
        Assert.Equal("Share", chart["title"]!.GetValue<string>());
        Assert.True(chart["dataEditable"]!.GetValue<bool>());
        Assert.Equal("Series A", chart["seriesNames"]![0]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Remove_ChartPath_DeletesFrameAndPart()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar")));

        var data = Edit(TestEnv.Op("remove", "/slide[1]/chart[1]"));
        Assert.Equal("/slide[1]/chart[1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            Assert.Empty(slidePart.Slide!.Descendants<P.GraphicFrame>());
            Assert.Empty(slidePart.ChartParts);
        }

        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/chart[1]")));
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Remove_ByShapeIdPath_AlsoDeletesChartPart()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("line", seriesCount: 1)));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("remove", shapePath));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            Assert.Empty(slidePart.Slide!.Descendants<P.GraphicFrame>());
            Assert.Empty(slidePart.ChartParts);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Set_OnChartPath_IsTypedUnsupported()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar")));

        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(("title", "New")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("remove", error.Suggestion, StringComparison.OrdinalIgnoreCase);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Set_GeometryOnChartShapePath_MovesTheFrame()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar")));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("x", "5cm"), ("w", "15cm"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/chart[1]"))));
        Assert.Equal(5, detail["x"]!.GetValue<double>());
        Assert.Equal(15, detail["w"]!.GetValue<double>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void PieChart_WithTwoSeries_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("pie", seriesCount: 2))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void UnknownChartKind_IsTypedUnsupportedWithCandidates()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("waterfall"))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Equal(
            ["bar", "line", "pie", "doughnut", "radar", "bubble", "stackedBar", "percentStackedBar", "stackedArea", "combo"],
            error.Candidates!);
    }

    [Theory]
    [InlineData("stackedBar", "barChart")]
    [InlineData("percentStackedBar", "barChart")]
    [InlineData("stackedArea", "areaChart")]
    [InlineData("radar", "radarChart")]
    public void AddExpandedKind_ReopensValidWithGroup(string kind, string groupLocalName)
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps(kind, seriesCount: 3)));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var chartPart = doc.PresentationPart!.SlideParts.Single().ChartParts.Single();
            var plotArea = chartPart.ChartSpace!.Descendants<C.PlotArea>().Single();
            var group = plotArea.ChildElements.Single(e => e.LocalName == groupLocalName);
            Assert.Equal(3, group.ChildElements.Count(e => e.LocalName == "ser"));
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void StackedBar_UsesStackedGroupingAndOverlap()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("stackedBar")));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var barChart = doc.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace!.Descendants<C.BarChart>().Single();
        Assert.Equal(C.BarGroupingValues.Stacked, barChart.GetFirstChild<C.BarGrouping>()!.Val!.Value);
        Assert.Equal(100, barChart.GetFirstChild<C.Overlap>()!.Val!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void PercentStackedBar_UsesPercentGrouping()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("percentStackedBar")));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var barChart = doc.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace!.Descendants<C.BarChart>().Single();
        Assert.Equal(C.BarGroupingValues.PercentStacked, barChart.GetFirstChild<C.BarGrouping>()!.Val!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Doughnut_WritesDoughnutChartWithHole()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("doughnut", "Share", seriesCount: 1)));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var chartSpace = doc.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace!;
        var doughnut = chartSpace.Descendants<C.DoughnutChart>().Single();
        Assert.Single(doughnut.Elements<C.PieChartSeries>());
        Assert.Equal(50, doughnut.GetFirstChild<C.HoleSize>()!.Val!.Value);
        Assert.Empty(chartSpace.Descendants<C.CategoryAxis>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Doughnut_WithTwoSeries_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("doughnut", seriesCount: 2))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Combo_WritesColumnGroupPlusLineGroup()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("combo", seriesCount: 3)));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var chartSpace = doc.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace!;
        var bar = chartSpace.Descendants<C.BarChart>().Single();
        var line = chartSpace.Descendants<C.LineChart>().Single();
        Assert.Single(bar.Elements<C.BarChartSeries>()); // first series as columns
        Assert.Equal(2, line.Elements<C.LineChartSeries>().Count()); // the rest as a line
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Combo_WithOneSeries_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("combo", seriesCount: 1))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Bubble_WritesXYSizeTriplePerSeries()
    {
        Create();
        var props = TestEnv.Props(
            ("kind", "bubble"),
            ("categories", new JsonArray("A", "B", "C")),
            ("series", new JsonArray(new JsonObject
            {
                ["name"] = "Deals",
                ["values"] = new JsonArray(10, 20, 30),
                ["x"] = new JsonArray(1, 2, 3),
                ["size"] = new JsonArray(5, 8, 3),
            })));
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: props));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var chartSpace = doc.PresentationPart!.SlideParts.Single().ChartParts.Single().ChartSpace!;
            var bubble = chartSpace.Descendants<C.BubbleChart>().Single();
            var series = bubble.Elements<C.BubbleChartSeries>().Single();
            Assert.NotNull(series.GetFirstChild<C.XValues>());
            Assert.NotNull(series.GetFirstChild<C.YValues>());
            var size = series.GetFirstChild<C.BubbleSize>()!;
            // The size cache carries the true sizes (5, 8, 3).
            Assert.Equal("5", size.Descendants<C.NumericValue>().First().Text);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Bubble_DefaultsXAndSizeWhenOmitted()
    {
        Create();
        // Same {name, values} shape as the other kinds: x defaults to the index, size to 1.
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bubble", seriesCount: 1)));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Bubble_MismatchedSizeLength_IsInvalidArgs()
    {
        Create();
        var props = TestEnv.Props(
            ("kind", "bubble"),
            ("categories", new JsonArray("A", "B")),
            ("series", new JsonArray(new JsonObject
            {
                ["values"] = new JsonArray(10, 20),
                ["size"] = new JsonArray(5), // one short
            })));
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "chart", props: props)]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Theory]
    [InlineData("stackedBar")]
    [InlineData("percentStackedBar")]
    [InlineData("stackedArea")]
    [InlineData("doughnut")]
    [InlineData("combo")]
    public void Svg_ExpandedKind_DrawsSomething(string kind)
    {
        Create();
        var seriesCount = kind == "doughnut" ? 1 : 2;
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps(kind, seriesCount: seriesCount)));

        var svg = RenderSlide1Svg();
        // Each new kind draws a real mini-chart element (bar/area/slice/hole), not a bare placeholder.
        var marker = kind switch
        {
            "doughnut" => "aio-chart-hole",
            "stackedArea" => "aio-chart-area",
            _ => "aio-chart-bar",
        };
        Assert.Contains(marker, svg, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("bubble")]
    [InlineData("radar")]
    public void Svg_BubbleAndRadar_RenderHonestPlaceholder(string kind)
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps(kind, seriesCount: 1)));

        var svg = RenderSlide1Svg();
        Assert.Contains("[chart] " + kind, svg, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SeriesValuesLengthMismatch_IsInvalidArgs()
    {
        Create();
        var props = TestEnv.Props(
            ("kind", "bar"),
            ("categories", new JsonArray("Q1", "Q2")),
            ("series", new JsonArray(new JsonObject { ["values"] = new JsonArray(1, 2, 3) })));
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "chart", props: props)]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("categories", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Svg_BarChart_DrawsBarsAxesAndLabels()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar", "Revenue")));

        var svg = RenderSlide1Svg();
        Assert.Equal(8, svg.Split("class=\"aio-chart-bar\"").Length - 1); // 2 series x 4 categories
        Assert.Contains("Q1", svg, StringComparison.Ordinal);
        Assert.Contains("Q4", svg, StringComparison.Ordinal);
        Assert.Contains("Revenue", svg, StringComparison.Ordinal);
        Assert.Contains("<g data-aio-path=\"/slide[1]/shape[@id=", svg, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Svg_LineChart_DrawsOnePolylinePerSeries()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("line", seriesCount: 3)));

        var svg = RenderSlide1Svg();
        Assert.Equal(3, svg.Split("class=\"aio-chart-line\"").Length - 1);
        Assert.Contains("Q2", svg, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Svg_PieChart_DrawsOneWedgePerCategory()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("pie", seriesCount: 1)));

        var svg = RenderSlide1Svg();
        Assert.Equal(4, svg.Split("class=\"aio-chart-slice\"").Length - 1);
        Assert.Contains("Q3", svg, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SlideDetail_ListsChartPaths()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: ChartProps("bar")));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        Assert.Equal("/slide[1]/chart[1]", detail["charts"]![0]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    private string RenderSlide1Svg()
    {
        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        return data["slides"]![0]!["svg"]!.GetValue<string>();
    }
}
