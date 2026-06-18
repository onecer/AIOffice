using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.12 "completion depth": editing an existing chart's title/categories/series in
/// place (set /slide[i]/chart[k]). Both the chart-XML cached values and the embedded
/// "Edit Data" workbook are rewritten so get/render and PowerPoint reflect the edit.
/// </summary>
public sealed class ChartEditInPlaceTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private void CreateBarChart(string? title = "Revenue", int seriesCount = 2)
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
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
            ("kind", "bar"),
            ("categories", new JsonArray("Q1", "Q2", "Q3", "Q4")),
            ("series", series));
        if (title is not null)
        {
            props["title"] = title;
        }

        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: props));
    }

    private JsonObject GetChart() =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/chart[1]"))));

    private ChartPart SingleChartPart(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.Single().ChartParts.Single();

    // ----- title ---------------------------------------------------------------

    [Fact]
    public void SetTitle_ReplacesExistingTitle()
    {
        CreateBarChart(title: "Old");

        var result = Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(("title", "New Title"))));
        Assert.Equal("/slide[1]/chart[1]", result["results"]![0]!["target"]!.GetValue<string>());

        Assert.Equal("New Title", GetChart()["title"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var chart = SingleChartPart(doc).ChartSpace!.Descendants<C.Title>().Single();
            Assert.Contains("New Title", chart.InnerText, StringComparison.Ordinal);
            Assert.False(SingleChartPart(doc).ChartSpace!.Descendants<C.AutoTitleDeleted>().Single().Val!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetTitle_OnUntitledChart_AddsTitle()
    {
        CreateBarChart(title: null);
        Assert.Null(GetChart()["title"]);

        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(("title", "Fresh"))));

        Assert.Equal("Fresh", GetChart()["title"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetTitle_False_RemovesTitle()
    {
        CreateBarChart(title: "Doomed");

        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(("title", JsonValue.Create(false)))));

        Assert.Null(GetChart()["title"]);
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Empty(SingleChartPart(doc).ChartSpace!.Descendants<C.Title>());
            Assert.True(SingleChartPart(doc).ChartSpace!.Descendants<C.AutoTitleDeleted>().Single().Val!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetTitle_True_IsInvalidArgs()
    {
        CreateBarChart();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "set", "/slide[1]/chart[1]", props: TestEnv.Props(("title", JsonValue.Create(true))))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ----- categories ----------------------------------------------------------

    [Fact]
    public void SetCategories_RelabelsEverySeries()
    {
        CreateBarChart();

        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(
            ("categories", new JsonArray("Jan", "Feb", "Mar", "Apr")))));

        var categories = GetChart()["categories"]!.AsArray();
        Assert.Equal(["Jan", "Feb", "Mar", "Apr"], categories.Select(c => c!.GetValue<string>()));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            // Every series' c:cat cache carries the new labels.
            foreach (var ser in SingleChartPart(doc).ChartSpace!.Descendants<C.BarChartSeries>())
            {
                var cats = ser.Descendants<C.CategoryAxisData>().Single()
                    .Descendants<C.StringPoint>().Select(p => p.NumericValue!.Text).ToList();
                Assert.Equal(["Jan", "Feb", "Mar", "Apr"], cats);
            }

            // And the embedded "Edit Data" workbook column A is rewritten to match.
            using var stream = SingleChartPart(doc).EmbeddedPackagePart!.GetStream();
            using var workbook = SpreadsheetDocument.Open(stream, false);
            var wbPart = workbook.WorkbookPart!;
            var sheet = wbPart.Workbook!.Descendants<S.Sheet>().Single();
            var cells = ((WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!))
                .Worksheet!.Descendants<S.Cell>()
                .ToDictionary(c => c.CellReference!.Value!, c => c.InnerText);
            Assert.Equal("Jan", cells["A2"]);
            Assert.Equal("Apr", cells["A5"]);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetCategories_CountChangeWithoutSeries_IsInvalidArgs()
    {
        CreateBarChart();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "set", "/slide[1]/chart[1]", props: TestEnv.Props(
                ("categories", new JsonArray("H1", "H2"))))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("series", error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    // ----- series --------------------------------------------------------------

    [Fact]
    public void SetSeries_ReplacesValuesAndNames_AndSyncsWorkbook()
    {
        CreateBarChart();

        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(
            ("series", new JsonArray(
                new JsonObject { ["name"] = "Plan", ["values"] = new JsonArray(1, 2, 3, 4) },
                new JsonObject { ["name"] = "Actual", ["values"] = new JsonArray(5, 6, 7, 8) })))));

        var detail = GetChart();
        var series = detail["series"]!.AsArray();
        Assert.Equal("Plan", series[0]!["name"]!.GetValue<string>());
        Assert.Equal("Actual", series[1]!["name"]!.GetValue<string>());
        Assert.Equal(1, series[0]!["values"]![0]!.GetValue<double>());
        Assert.Equal(8, series[1]!["values"]![3]!.GetValue<double>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            // The embedded workbook holds the new names and values.
            using var stream = SingleChartPart(doc).EmbeddedPackagePart!.GetStream();
            using var workbook = SpreadsheetDocument.Open(stream, false);
            var wbPart = workbook.WorkbookPart!;
            var sheet = wbPart.Workbook!.Descendants<S.Sheet>().Single();
            var cells = ((WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!))
                .Worksheet!.Descendants<S.Cell>()
                .ToDictionary(c => c.CellReference!.Value!, c => c.InnerText);
            Assert.Equal("Plan", cells["B1"]);
            Assert.Equal("Actual", cells["C1"]);
            Assert.Equal("1", cells["B2"]);
            Assert.Equal("8", cells["C5"]);
        }

        Assert.True(detail["dataEditable"]!.GetValue<bool>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetSeries_KeepsNameWhenOmitted()
    {
        CreateBarChart();

        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(
            ("series", new JsonArray(
                new JsonObject { ["values"] = new JsonArray(9, 9, 9, 9) },
                new JsonObject { ["values"] = new JsonArray(8, 8, 8, 8) })))));

        var series = GetChart()["series"]!.AsArray();
        Assert.Equal("Series A", series[0]!["name"]!.GetValue<string>()); // unchanged
        Assert.Equal(9, series[0]!["values"]![0]!.GetValue<double>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetSeries_FewerThanExisting_LeavesTrailingUntouched()
    {
        CreateBarChart(seriesCount: 2);

        // Only re-supply the first series; the second keeps its original data.
        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(
            ("series", new JsonArray(
                new JsonObject { ["name"] = "Only", ["values"] = new JsonArray(100, 100, 100, 100) })))));

        var series = GetChart()["series"]!.AsArray();
        Assert.Equal(2, series.Count);
        Assert.Equal("Only", series[0]!["name"]!.GetValue<string>());
        Assert.Equal(100, series[0]!["values"]![0]!.GetValue<double>());
        Assert.Equal("Series B", series[1]!["name"]!.GetValue<string>()); // untouched
        Assert.Equal(15, series[1]!["values"]![0]!.GetValue<double>()); // 10 + (1*5)
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetSeries_ValueCountMismatch_IsInvalidArgs()
    {
        CreateBarChart();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "set", "/slide[1]/chart[1]", props: TestEnv.Props(
                ("series", new JsonArray(
                    new JsonObject { ["values"] = new JsonArray(1, 2) }))))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("categories", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ----- combined ------------------------------------------------------------

    [Fact]
    public void SetTitleCategoriesAndSeries_Together_ReshapeTheChart()
    {
        CreateBarChart(title: "Before");

        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(
            ("title", "After"),
            ("categories", new JsonArray("North", "South", "East")),
            ("series", new JsonArray(
                new JsonObject { ["name"] = "2025", ["values"] = new JsonArray(3, 6, 9) },
                new JsonObject { ["name"] = "2026", ["values"] = new JsonArray(4, 8, 12) })))));

        var detail = GetChart();
        Assert.Equal("After", detail["title"]!.GetValue<string>());
        Assert.Equal(["North", "South", "East"], detail["categories"]!.AsArray().Select(c => c!.GetValue<string>()));
        var series = detail["series"]!.AsArray();
        Assert.Equal("2025", series[0]!["name"]!.GetValue<string>());
        Assert.Equal(12, series[1]!["values"]![2]!.GetValue<double>());

        // Render still produces a real bar chart over the new data.
        var svg = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))))
            ["slides"]![0]!["svg"]!.GetValue<string>();
        Assert.Equal(6, svg.Split("class=\"aio-chart-bar\"").Length - 1); // 2 series x 3 categories
        Assert.Contains("After", svg, StringComparison.Ordinal);
        Assert.Contains("North", svg, StringComparison.Ordinal);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetData_AlongsidePolish_AppliesBoth()
    {
        CreateBarChart();

        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(
            ("title", "Polished"),
            ("legend", "bottom"))));

        var detail = GetChart();
        Assert.Equal("Polished", detail["title"]!.GetValue<string>());
        Assert.Equal("bottom", detail["polish"]!["legend"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ----- foreign cached-only chart -------------------------------------------

    [Fact]
    public void SetSeries_OnCachedOnlyChart_UpdatesCachesWithoutWorkbook()
    {
        CreateBarChart();
        StripEmbeddedWorkbook(); // simulate a foreign cached-only chart
        Assert.False(GetChart()["dataEditable"]!.GetValue<bool>());

        Edit(TestEnv.Op("set", "/slide[1]/chart[1]", props: TestEnv.Props(
            ("series", new JsonArray(
                new JsonObject { ["name"] = "X", ["values"] = new JsonArray(7, 7, 7, 7) },
                new JsonObject { ["name"] = "Y", ["values"] = new JsonArray(2, 2, 2, 2) })))));

        var series = GetChart()["series"]!.AsArray();
        Assert.Equal(7, series[0]!["values"]![0]!.GetValue<double>());
        Assert.False(GetChart()["dataEditable"]!.GetValue<bool>()); // still no workbook
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    private void StripEmbeddedWorkbook()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true);
        var chartPart = SingleChartPart(doc);
        foreach (var external in chartPart.ChartSpace!.Elements<C.ExternalData>().ToList())
        {
            external.Remove();
        }

        if (chartPart.EmbeddedPackagePart is { } embedded)
        {
            chartPart.DeletePart(embedded);
        }
    }
}
