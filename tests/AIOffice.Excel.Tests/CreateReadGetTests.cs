using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

public sealed class CreateReadGetTests : ExcelTestBase
{
    [Fact]
    public void Create_produces_validator_clean_workbook_with_titled_sheet()
    {
        var file = NewFile("report.xlsx");

        var envelope = Handler.Create(Ctx(file, ("title", "Q3 Data")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.True(File.Exists(file));
        Assert.NotNull(envelope.Meta.Rev);
        var data = Json(envelope)["data"]!;
        Assert.Equal("Q3 Data", data["sheets"]![0]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Create_refuses_to_overwrite_an_existing_file()
    {
        var file = CreateWorkbook();

        var envelope = Handler.Create(Ctx(file));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    [Fact]
    public void Read_stats_counts_cells_and_formulas()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1)),
            SetOp("/Sheet1/A2", ("value", 2)),
            SetOp("/Sheet1/A3", ("value", "=SUM(A1:A2)"))).IsOk);

        var data = OkData(Handler.Read(Ctx(file))); // stats is the default view

        var sheet = data["sheets"]![0]!;
        Assert.Equal("Sheet1", sheet["name"]!.GetValue<string>());
        Assert.Equal("A1:A3", sheet["usedRange"]!.GetValue<string>());
        Assert.Equal(3, sheet["cellCount"]!.GetValue<int>());
        Assert.Equal(1, sheet["formulaCount"]!.GetValue<int>());
    }

    [Fact]
    public void Read_text_returns_a_csv_window_for_a_range()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:B2", ("values", new JsonArray(
                new JsonArray("a,b", 1),
                new JsonArray("c", 2))))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "text"), ("range", "A1:B2"))));

        var content = data["content"]!.GetValue<string>();
        Assert.Contains("# Sheet1!A1:B2", content, StringComparison.Ordinal);
        Assert.Contains("\"a,b\",1", content, StringComparison.Ordinal); // RFC 4180 escaping
        Assert.False(data["truncated"]!.GetValue<bool>());
    }

    [Fact]
    public void Read_structure_lists_tables_and_merges()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:B2", ("values", new JsonArray(
                new JsonArray("Name", "Qty"),
                new JsonArray("ant", 2)))),
            AddOp("/Sheet1/A1:B2", "table", ("name", "Stock")),
            SetOp("/Sheet1/D1:E1", ("merge", true))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "structure"))));

        var sheet = data["sheets"]![0]!;
        var table = sheet["tables"]![0]!;
        Assert.Equal("Stock", table["name"]!.GetValue<string>());
        Assert.Equal("A1:B2", table["range"]!.GetValue<string>());
        Assert.Equal("Name", table["columns"]![0]!.GetValue<string>());
        Assert.Equal("D1:E1", sheet["mergedRanges"]![0]!.GetValue<string>());
    }

    [Fact]
    public void Read_unknown_view_lists_the_valid_views()
    {
        var file = CreateWorkbook();

        var envelope = Handler.Read(Ctx(file, ("view", "tree")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("outline", envelope.Error.Candidates!);
    }

    [Fact]
    public void Get_formula_cell_reports_value_formula_and_cached_value()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1)),
            SetOp("/Sheet1/A2", ("value", 2)),
            SetOp("/Sheet1/A3", ("value", "=SUM(A1:A2)"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A3"))));

        Assert.Equal(3.0, data["value"]!.GetValue<double>());
        Assert.Equal("=SUM(A1:A2)", data["formula"]!.GetValue<string>());
        Assert.Equal(3.0, data["cachedValue"]!.GetValue<double>());
        Assert.Equal("number", data["type"]!.GetValue<string>());
    }

    [Fact]
    public void Get_range_returns_compact_2d_values()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:B2", ("values", new JsonArray(
                new JsonArray(1, 2),
                new JsonArray(3, 4))))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1:B2"))));

        Assert.Equal(2, data["rows"]!.GetValue<int>());
        Assert.Equal(2, data["columns"]!.GetValue<int>());
        Assert.Equal(4.0, data["values"]![1]![1]!.GetValue<double>());
        Assert.False(data["truncated"]!.GetValue<bool>());
    }

    [Fact]
    public void Get_large_range_truncates_with_a_warning()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", 1))).IsOk);

        var envelope = Handler.Get(Ctx(file, ("path", "/Sheet1/A1:B10"), ("maxCells", 4)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var data = Json(envelope)["data"]!;
        Assert.True(data["truncated"]!.GetValue<bool>());
        Assert.Equal(2, data["values"]!.AsArray().Count); // 4 cells / 2 columns = 2 rows
        var warning = Assert.Single(envelope.Meta.Warnings!);
        Assert.Equal("result_truncated", warning.Code);
    }

    [Fact]
    public void Get_unknown_sheet_attaches_nearest_candidates()
    {
        var file = CreateWorkbook();

        var envelope = Handler.Get(Ctx(file, ("path", "/Shee1/A1")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1", envelope.Error.Candidates!);
    }

    [Fact]
    public void Quoted_sheet_names_resolve()
    {
        var file = CreateWorkbook(title: "Q3 Data");
        Assert.True(EditOps(file, SetOp("/'Q3 Data'/A1", ("value", 7))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/'Q3 Data'/A1"))));

        Assert.Equal(7.0, data["value"]!.GetValue<double>());
        Assert.Equal("/'Q3 Data'/A1", data["path"]!.GetValue<string>());
    }
}
