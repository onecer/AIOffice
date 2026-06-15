using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M1 chart slice: bar | line | pie charts authored on raw OpenXml in a
/// post-save pass over the package ClosedXML writes. Raw assertions reopen the
/// file with the OpenXml SDK so the chart writer cannot grade its own homework.
/// </summary>
public sealed class ChartTests : ExcelTestBase
{
    /// <summary>Header data Month/Sales/Cost in A1:C5.</summary>
    private string CreateDataWorkbook()
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

    /// <summary>Chart parts of the first sheet in drawing (anchor) order, plus the worksheet part.</summary>
    private static (WorksheetPart Sheet, List<ChartPart> Charts) RawCharts(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var charts = new List<ChartPart>();
        var drawings = worksheetPart.DrawingsPart;
        if (drawings?.WorksheetDrawing is { } root)
        {
            foreach (var anchor in root.Elements<Xdr.TwoCellAnchor>())
            {
                if (anchor.Descendants<C.ChartReference>().FirstOrDefault()?.Id?.Value is { } rid)
                {
                    charts.Add((ChartPart)drawings.GetPartById(rid));
                }
            }
        }

        return (worksheetPart, charts);
    }

    [Fact]
    public void Add_bar_line_pie_charts_validator_clean_with_raw_xml_essentials()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(
            file,
            AddOp("/Sheet1", "chart",
                ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"), ("title", "Sales by month")),
            AddOp("/Sheet1", "chart",
                ("kind", "line"), ("dataRange", "A1:B5"), ("anchor", "E20")),
            AddOp("/Sheet1", "chart",
                ("kind", "pie"), ("dataRange", "A1:B5"), ("anchor", "E40"),
                ("widthCells", 6), ("heightCells", 10)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var details = Json(envelope)["data"]!["ops"]!.AsArray();
        Assert.Equal("/Sheet1/chart[1]", details[0]!["path"]!.GetValue<string>());
        Assert.Equal("/Sheet1/chart[3]", details[2]!["path"]!.GetValue<string>());
        Assert.Equal(2, details[0]!["series"]!.GetValue<int>());

        AssertValidatorClean(file);

        // Raw oracle: reopen with the OpenXml SDK and pin the chart XML essentials.
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (_, charts) = RawCharts(document);
        Assert.Equal(3, charts.Count);

        // chart[1]: bar, two series, cached refs and values.
        var bar = charts[0].ChartSpace!;
        Assert.NotNull(bar.Descendants<C.BarChart>().SingleOrDefault());
        var barSeries = bar.Descendants<C.BarChartSeries>().ToList();
        Assert.Equal(2, barSeries.Count);

        var sales = barSeries[0];
        Assert.Equal(
            "'Sheet1'!$B$1",
            sales.GetFirstChild<C.SeriesText>()!.StringReference!.Formula!.Text);
        Assert.Equal(
            "'Sheet1'!$A$2:$A$5",
            sales.GetFirstChild<C.CategoryAxisData>()!.StringReference!.Formula!.Text);
        var salesValues = sales.GetFirstChild<C.Values>()!.NumberReference!;
        Assert.Equal("'Sheet1'!$B$2:$B$5", salesValues.Formula!.Text);
        Assert.Equal(
            ["10", "20", "15", "25"],
            salesValues.NumberingCache!.Descendants<C.NumericPoint>().Select(p => p.NumericValue!.Text));
        Assert.Equal(
            ["Jan", "Feb", "Mar", "Apr"],
            sales.GetFirstChild<C.CategoryAxisData>()!.StringReference!.StringCache!
                .Descendants<C.StringPoint>().Select(p => p.NumericValue!.Text));
        Assert.Equal(
            "'Sheet1'!$C$2:$C$5",
            barSeries[1].GetFirstChild<C.Values>()!.NumberReference!.Formula!.Text);
        Assert.Equal(
            "Sales by month",
            string.Concat(bar.Descendants<C.Title>().Single().Descendants<A.Text>().Select(t => t.Text)));

        // chart[2]: line, one series, no title (autoTitleDeleted).
        var line = charts[1].ChartSpace!;
        Assert.NotNull(line.Descendants<C.LineChart>().SingleOrDefault());
        Assert.Single(line.Descendants<C.LineChartSeries>());
        Assert.Empty(line.Descendants<C.Title>());
        Assert.True(line.Descendants<C.AutoTitleDeleted>().Single().Val!.Value);

        // chart[3]: pie, one series, no axes.
        var pie = charts[2].ChartSpace!;
        Assert.NotNull(pie.Descendants<C.PieChart>().SingleOrDefault());
        Assert.Single(pie.Descendants<C.PieChartSeries>());
        Assert.Empty(pie.Descendants<C.CategoryAxis>());
    }

    [Fact]
    public void Get_chart_returns_kind_title_dataRange_anchor_series()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"), ("title", "Sales by month"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]"))));

        Assert.Equal("/Sheet1/chart[1]", data["path"]!.GetValue<string>());
        Assert.Equal("chart", data["kind"]!.GetValue<string>());
        Assert.Equal("bar", data["chartKind"]!.GetValue<string>());
        Assert.Equal("Sales by month", data["title"]!.GetValue<string>());
        Assert.Equal("A1:C5", data["dataRange"]!.GetValue<string>()); // union of name/cat/val refs
        Assert.Equal("E2", data["anchor"]!.GetValue<string>());
        Assert.Equal(2, data["series"]!.GetValue<int>());
    }

    [Fact]
    public void Get_missing_chart_is_invalid_path_with_candidates()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "pie"), ("dataRange", "A1:B5"), ("anchor", "E2"))).IsOk);

        var envelope = Handler.Get(Ctx(file, ("path", "/Sheet1/chart[2]")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1/chart[1]", envelope.Error.Candidates!);
    }

    [Fact]
    public void Read_structure_lists_charts()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "chart",
                ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"), ("title", "Sales by month")),
            AddOp("/Sheet1", "chart",
                ("kind", "pie"), ("dataRange", "A1:B5"), ("anchor", "E20"))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "structure"))));

        var charts = data["sheets"]![0]!["charts"]!.AsArray();
        Assert.Equal(2, charts.Count);
        Assert.Equal("/Sheet1/chart[1]", charts[0]!["path"]!.GetValue<string>());
        Assert.Equal("bar", charts[0]!["kind"]!.GetValue<string>());
        Assert.Equal("Sales by month", charts[0]!["title"]!.GetValue<string>());
        Assert.Equal("pie", charts[1]!["kind"]!.GetValue<string>());
        Assert.Equal("E20", charts[1]!["anchor"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_chart_shifts_indices_and_cleans_up_empty_drawings()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "chart", ("kind", "bar"), ("dataRange", "A1:B5"), ("anchor", "E2")),
            AddOp("/Sheet1", "chart", ("kind", "line"), ("dataRange", "A1:B5"), ("anchor", "E20"))).IsOk);

        var remove = EditOps(file, RemoveOp("/Sheet1/chart[1]"));
        Assert.True(remove.IsOk, remove.ToJson());
        AssertValidatorClean(file);

        // The line chart shifted into chart[1].
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]"))));
        Assert.Equal("line", data["chartKind"]!.GetValue<string>());

        // Removing the last chart deletes the now-empty drawings part AND the
        // worksheet's <drawing> element (validator-clean emptiness).
        Assert.True(EditOps(file, RemoveOp("/Sheet1/chart[1]")).IsOk);
        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (worksheetPart, charts) = RawCharts(document);
        Assert.Empty(charts);
        Assert.Null(worksheetPart.DrawingsPart);
        Assert.Empty(worksheetPart.Worksheet!.Elements<S.Drawing>());
    }

    [Fact]
    public void Remove_missing_chart_is_invalid_path_with_candidates()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:B5"), ("anchor", "E2"))).IsOk);

        var envelope = EditOps(file, RemoveOp("/Sheet1/chart[3]"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1/chart[1]", envelope.Error.Candidates!);
    }

    [Fact]
    public void Other_chart_kinds_are_unsupported_naming_the_supported_set()
    {
        var file = CreateDataWorkbook();

        // bubble joined the supported set in v1.1; a genuinely unsupported kind
        // still names the (now expanded) supported set.
        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "surface"), ("dataRange", "A1:B5"), ("anchor", "E2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("bar", envelope.Error.Suggestion, StringComparison.Ordinal);
        Assert.Contains("line", envelope.Error.Suggestion, StringComparison.Ordinal);
        Assert.Contains("pie", envelope.Error.Suggestion, StringComparison.Ordinal);
        Assert.Contains("scatter", envelope.Error.Suggestion, StringComparison.Ordinal);
        Assert.Contains("area", envelope.Error.Suggestion, StringComparison.Ordinal);
        Assert.Equal(
            ["bar", "line", "pie", "scatter", "area",
             "doughnut", "radar", "bubble",
             "stackedBar", "percentStackedBar", "stackedArea", "combo"],
            envelope.Error.Candidates!);
    }

    [Fact]
    public void Chart_survives_a_later_unrelated_edit()
    {
        // Pins the cooperation contract with ClosedXML's save pipeline: chart
        // parts written by the post-save pass must survive later edits that go
        // through ClosedXML.
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:C5"), ("anchor", "E2"), ("title", "Sales by month"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/A9", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]"))));
        Assert.Equal("bar", data["chartKind"]!.GetValue<string>());
        Assert.Equal("Sales by month", data["title"]!.GetValue<string>());
    }

    [Fact]
    public void Headerless_dataRange_gets_default_series_names()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B3", ("values", new JsonArray(
            new JsonArray("Jan", 10),
            new JsonArray("Feb", 20),
            new JsonArray("Mar", 15))))).IsOk);

        Assert.True(EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:B3"), ("anchor", "D1"))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (_, charts) = RawCharts(document);
        var series = Assert.Single(charts.Single().ChartSpace!.Descendants<C.BarChartSeries>());

        // No header row: the series name is a literal <c:v>, not a cell ref,
        // and the categories include row 1.
        var seriesText = series.GetFirstChild<C.SeriesText>()!;
        Assert.Null(seriesText.StringReference);
        Assert.Equal("Series 1", seriesText.GetFirstChild<C.NumericValue>()!.Text);
        Assert.Equal(
            "'Sheet1'!$A$1:$A$3",
            series.GetFirstChild<C.CategoryAxisData>()!.StringReference!.Formula!.Text);
    }

    [Fact]
    public void Pie_with_multiple_series_is_refused()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "pie"), ("dataRange", "A1:C5"), ("anchor", "E2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("one series", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Text_in_series_data_is_refused_naming_the_cell()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/B3", ("value", "oops"), ("valueType", "text"))).IsOk);

        var envelope = EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:B5"), ("anchor", "E2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("/Sheet1/B3", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Missing_required_chart_prop_fails_actionably()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "chart", ("kind", "bar"), ("anchor", "E2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("dataRange", envelope.Error.Message, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion));
    }

    [Fact]
    public void Dry_run_chart_add_writes_nothing()
    {
        var file = CreateDataWorkbook();
        var revBefore = Rev.OfFile(file);
        var snapshotsBefore = Snapshots.List(file).Count;

        var envelope = Handler.Edit(
            Ctx(file, ("dryRun", true)),
            [AddOp("/Sheet1", "chart", ("kind", "bar"), ("dataRange", "A1:B5"), ("anchor", "E2"))]);

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.True(Json(envelope)["data"]!["dryRun"]!.GetValue<bool>());
        Assert.Equal(revBefore, Rev.OfFile(file));
        Assert.Equal(snapshotsBefore, Snapshots.List(file).Count);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (_, charts) = RawCharts(document);
        Assert.Empty(charts);
    }

    [Fact]
    public void Failing_batch_with_chart_op_leaves_no_trace()
    {
        var file = CreateDataWorkbook();
        var bytesBefore = File.ReadAllBytes(file);

        var envelope = EditOps(
            file,
            AddOp("/Sheet1", "chart", ("kind", "bar"), ("dataRange", "A1:B5"), ("anchor", "E2")),
            SetOp("/Nope/A1", ("value", 1))); // unknown sheet aborts the batch

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Equal(bytesBefore, File.ReadAllBytes(file)); // atomic: nothing written
    }

    [Fact]
    public void Set_on_a_chart_is_a_typed_unsupported_feature()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "chart",
            ("kind", "bar"), ("dataRange", "A1:B5"), ("anchor", "E2"))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/chart[1]", ("value", 1)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("add", envelope.Error.Suggestion, StringComparison.Ordinal); // workaround named
    }

    [Fact]
    public void Chart_on_quoted_sheet_name_quotes_its_formula_refs()
    {
        var file = NewFile();
        Assert.True(Handler.Create(Ctx(file, ("title", "Q3 Data"))).IsOk);
        Assert.True(EditOps(file, SetOp("/'Q3 Data'/A1:B3", ("values", new JsonArray(
            new JsonArray("Item", "Qty"),
            new JsonArray("ant", 5),
            new JsonArray("bee", 7))))).IsOk);

        Assert.True(EditOps(file, AddOp("/'Q3 Data'", "chart",
            ("kind", "bar"), ("dataRange", "A1:B3"), ("anchor", "D1"))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (_, charts) = RawCharts(document);
        var series = Assert.Single(charts.Single().ChartSpace!.Descendants<C.BarChartSeries>());
        Assert.Equal(
            "'Q3 Data'!$B$2:$B$3",
            series.GetFirstChild<C.Values>()!.NumberReference!.Formula!.Text);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/'Q3 Data'/chart[1]"))));
        Assert.Equal("A1:B3", data["dataRange"]!.GetValue<string>());
    }
}
