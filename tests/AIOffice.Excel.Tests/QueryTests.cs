using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

public sealed class QueryTests : ExcelTestBase
{
    [Fact]
    public void Numeric_filter_returns_canonical_paths()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 50)),
            SetOp("/Sheet1/A2", ("value", 150)),
            SetOp("/Sheet1/A3", ("value", 200))).IsOk);

        var data = OkData(Handler.Query(Ctx(file, ("selector", "cell[value>100]"))));

        Assert.Equal(2, data["total"]!.GetValue<int>());
        var paths = data["matches"]!.AsArray().Select(m => m!["path"]!.GetValue<string>()).ToHashSet();
        Assert.Contains("/Sheet1/A2", paths);
        Assert.Contains("/Sheet1/A3", paths);
    }

    [Fact]
    public void Bare_formula_attribute_finds_formula_cells()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1)),
            SetOp("/Sheet1/A2", ("value", 2)),
            SetOp("/Sheet1/A3", ("value", "=SUM(A1:A2)"))).IsOk);

        var data = OkData(Handler.Query(Ctx(file, ("selector", "cell[formula]"))));

        var match = Assert.Single(data["matches"]!.AsArray());
        Assert.Equal("/Sheet1/A3", match!["path"]!.GetValue<string>());
        Assert.Equal("=SUM(A1:A2)", match["formula"]!.GetValue<string>());
        Assert.Equal(3.0, match["value"]!.GetValue<double>());
    }

    [Fact]
    public void Table_rows_filter_by_column_value()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:B3", ("values", new JsonArray(
                new JsonArray("Name", "Qty"),
                new JsonArray("ant", 2),
                new JsonArray("bee", 7)))),
            AddOp("/Sheet1/A1:B3", "table", ("name", "Stock"))).IsOk);

        var data = OkData(Handler.Query(Ctx(file, ("selector", "row[Qty>5]"))));

        var match = Assert.Single(data["matches"]!.AsArray());
        Assert.Equal("/Sheet1/A3:B3", match!["path"]!.GetValue<string>());
        Assert.Equal("Stock", match["table"]!.GetValue<string>());
        Assert.Equal("bee", match["values"]!["Name"]!.GetValue<string>());
        Assert.Equal(7.0, match["values"]!["Qty"]!.GetValue<double>());
    }

    [Fact]
    public void Unknown_table_column_lists_known_columns_as_candidates()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:B2", ("values", new JsonArray(
                new JsonArray("Name", "Qty"),
                new JsonArray("ant", 2)))),
            AddOp("/Sheet1/A1:B2", "table")).IsOk);

        var envelope = Handler.Query(Ctx(file, ("selector", "row[Price>5]")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("Qty", envelope.Error.Candidates!);
        Assert.Contains("Name", envelope.Error.Candidates!);
    }

    [Fact]
    public void Contains_pseudo_class_matches_display_text()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "Q3 report")),
            SetOp("/Sheet1/A2", ("value", "Q4 report"))).IsOk);

        var data = OkData(Handler.Query(Ctx(file, ("selector", "cell:contains('q3')"))));

        var match = Assert.Single(data["matches"]!.AsArray());
        Assert.Equal("/Sheet1/A1", match!["path"]!.GetValue<string>());
    }

    [Fact]
    public void Unknown_cell_attribute_is_invalid_args_with_candidates()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", 1))).IsOk);

        var envelope = Handler.Query(Ctx(file, ("selector", "cell[font=Arial]")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("value", envelope.Error.Candidates!);
    }
}
