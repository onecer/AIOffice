using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Pptx.Tests;

/// <summary>M4 flagship: every chart embeds a minimal real workbook so PowerPoint's "Edit Data" works.</summary>
public sealed class EmbeddedChartDataTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void CreateWithChart(string kind = "bar", int seriesCount = 2)
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var series = new JsonArray();
        for (var i = 0; i < seriesCount; i++)
        {
            series.Add(new JsonObject
            {
                ["name"] = $"Series {(char)('A' + i)}",
                ["values"] = new JsonArray(10 + (i * 5), 20 + (i * 5), 15 + (i * 5)),
            });
        }

        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "add", "/slide[1]", type: "chart", props: TestEnv.Props(
                ("kind", kind),
                ("categories", new JsonArray("Jan", "Feb", "Mar")),
                ("series", series)))]));
    }

    private ChartPart SingleChartPart(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.Single().ChartParts.Single();

    [Fact]
    public void AddChart_EmbedsWorkbookPart_WiredViaExternalData()
    {
        CreateWithChart();

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var chartPart = SingleChartPart(doc);
            var embedded = chartPart.EmbeddedPackagePart;
            Assert.NotNull(embedded);
            Assert.Equal(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                embedded!.ContentType);

            var external = chartPart.ChartSpace!.Elements<C.ExternalData>().Single();
            Assert.False(external.GetFirstChild<C.AutoUpdate>()!.Val!.Value);
            Assert.Same(embedded, chartPart.GetPartById(external.Id!.Value!));
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void EmbeddedWorkbook_HoldsTheData_AtTheCachedRanges()
    {
        CreateWithChart();

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var chartPart = SingleChartPart(doc);

        // The caches reference Sheet1 ranges...
        var series = chartPart.ChartSpace!.Descendants<C.BarChartSeries>().ToList();
        Assert.Equal(
            "Sheet1!$B$1",
            series[0].Descendants<C.StringReference>().Single(r => r.Parent is C.SeriesText)
                .GetFirstChild<C.Formula>()!.Text);
        Assert.Equal(
            "Sheet1!$A$2:$A$4",
            series[0].Descendants<C.StringReference>().Single(r => r.Parent is C.CategoryAxisData)
                .GetFirstChild<C.Formula>()!.Text);
        Assert.Equal(
            "Sheet1!$C$2:$C$4",
            series[1].Descendants<C.NumberReference>().Single().GetFirstChild<C.Formula>()!.Text);

        // ...and the embedded workbook holds exactly those cells on Sheet1.
        using var workbookStream = chartPart.EmbeddedPackagePart!.GetStream();
        using var workbook = SpreadsheetDocument.Open(workbookStream, false);
        var workbookPart = workbook.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().Single();
        Assert.Equal("Sheet1", sheet.Name!.Value);

        var cells = ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!))
            .Worksheet!.Descendants<S.Cell>()
            .ToDictionary(c => c.CellReference!.Value!, c => c.InnerText);
        Assert.Equal("Series A", cells["B1"]);
        Assert.Equal("Series B", cells["C1"]);
        Assert.Equal("Jan", cells["A2"]);
        Assert.Equal("Mar", cells["A4"]);
        Assert.Equal("10", cells["B2"]);
        Assert.Equal("15", cells["C2"]);
        Assert.Equal("20", cells["C4"]);
    }

    [Fact]
    public void SetEmbedData_RetrofitsACachedOnlyChart()
    {
        CreateWithChart();
        StripEmbeddedWorkbook(); // simulate an M3-era / foreign cached-only chart

        var before = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/chart[1]"))));
        Assert.False(before["dataEditable"]!.GetValue<bool>());

        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "set", "/slide[1]/chart[1]", props: TestEnv.Props(("embedData", JsonValue.Create(true))))]));
        Assert.Equal("/slide[1]/chart[1]", data["results"]![0]!["target"]!.GetValue<string>());

        var after = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/chart[1]"))));
        Assert.True(after["dataEditable"]!.GetValue<bool>());
        Assert.Equal("Jan", after["categories"]![0]!.GetValue<string>());
        Assert.Equal(10, after["series"]![0]!["values"]![0]!.GetValue<double>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var chartPart = SingleChartPart(doc);
            Assert.NotNull(chartPart.EmbeddedPackagePart);
            Assert.Equal(
                "Sheet1!$B$2:$B$4",
                chartPart.ChartSpace!.Descendants<C.BarChartSeries>().First()
                    .Descendants<C.NumberReference>().Single().GetFirstChild<C.Formula>()!.Text);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetEmbedData_IsIdempotent_OnAlreadyEmbeddedCharts()
    {
        CreateWithChart();
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "set", "/slide[1]/chart[1]", props: TestEnv.Props(("embedData", JsonValue.Create(true))))]));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var chartPart = SingleChartPart(doc);
            Assert.Single(chartPart.ChartSpace!.Elements<C.ExternalData>());
            Assert.NotNull(chartPart.EmbeddedPackagePart);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetEmbedData_False_IsInvalidArgs()
    {
        CreateWithChart();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "set", "/slide[1]/chart[1]", props: TestEnv.Props(("embedData", JsonValue.Create(false))))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void PieChart_AlsoEmbedsItsWorkbook()
    {
        CreateWithChart(kind: "pie", seriesCount: 1);

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var chartPart = SingleChartPart(doc);
            Assert.NotNull(chartPart.EmbeddedPackagePart);
            Assert.Single(chartPart.ChartSpace!.Elements<C.ExternalData>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    /// <summary>Reverts a chart to M3 form: literal caches only, no externalData, no embedded part.</summary>
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
