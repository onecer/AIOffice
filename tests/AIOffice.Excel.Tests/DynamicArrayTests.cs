using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.4) Golden-value tests for write-time dynamic-array evaluation + spill:
/// FILTER / UNIQUE / SORT / SORTBY / SEQUENCE / TRANSPOSE compute their result
/// array, spill it into the rectangle anchored at the formula cell with cached
/// values, and clear the formula_not_evaluated warning. RANDARRAY is exercised
/// only for shape (its values are non-deterministic vs. Excel, deterministic for
/// us). spill_blocked guards an occupied target.
/// </summary>
public sealed class DynamicArrayTests : ExcelTestBase
{
    private string SeedColumn(params object[] values)
    {
        var file = CreateWorkbook();
        var rows = new JsonArray();
        foreach (var v in values)
        {
            rows.Add(new JsonArray(JsonValue.Create(v)));
        }

        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("values", rows))).IsOk);
        return file;
    }

    [Fact]
    public void Unique_spills_the_distinct_set_in_first_seen_order()
    {
        var file = SeedColumn("apple", "banana", "apple", "cherry", "banana");

        var envelope = EditOps(file, SetOp("/Sheet1/C1", ("value", "=UNIQUE(A1:A5)")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Null(Json(envelope)["meta"]!["warnings"]); // evaluated, no formula_not_evaluated

        var range = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1:C3"))));
        Assert.Equal("apple", range["values"]![0]![0]!.GetValue<string>());
        Assert.Equal("banana", range["values"]![1]![0]!.GetValue<string>());
        Assert.Equal("cherry", range["values"]![2]![0]!.GetValue<string>());

        // The anchor reports the array formula AND the spill rectangle it fills.
        var anchor = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal("=UNIQUE(A1:A5)", anchor["formula"]!.GetValue<string>());
        Assert.Equal("C1:C3", anchor["spillRange"]!.GetValue<string>());

        var (formula, cached, _) = RawCell(file, "Sheet1", "C1");
        Assert.Contains("UNIQUE", formula!, StringComparison.Ordinal);
        Assert.NotNull(cached); // cached value present
        var (spillFormula, spillCached, _) = RawCell(file, "Sheet1", "C2");
        Assert.Null(spillFormula);
        Assert.NotNull(spillCached);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Sort_spills_sorted_ascending_then_descending()
    {
        var file = SeedColumn(3, 1, 2);

        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=SORT(A1:A3)"))).IsOk);
        var asc = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1:C3"))));
        Assert.Equal(1.0, asc["values"]![0]![0]!.GetValue<double>());
        Assert.Equal(2.0, asc["values"]![1]![0]!.GetValue<double>());
        Assert.Equal(3.0, asc["values"]![2]![0]!.GetValue<double>());

        Assert.True(EditOps(file, SetOp("/Sheet1/E1", ("value", "=SORT(A1:A3,1,-1)"))).IsOk);
        var desc = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/E1:E3"))));
        Assert.Equal(3.0, desc["values"]![0]![0]!.GetValue<double>());
        Assert.Equal(1.0, desc["values"]![2]![0]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Filter_spills_the_rows_whose_condition_is_truthy()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray("North", 100),
                new JsonArray("South", 50),
                new JsonArray("North", 80)))),
            // Flag column: keep rows 1 and 3.
            SetOp("/Sheet1/D1", ("values", new JsonArray(
                new JsonArray(true), new JsonArray(false), new JsonArray(true))))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/F1", ("value", "=FILTER(A1:B3,D1:D3)"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/F1:G2"))));
        Assert.Equal("North", data["values"]![0]![0]!.GetValue<string>());
        Assert.Equal(100.0, data["values"]![0]![1]!.GetValue<double>());
        Assert.Equal("North", data["values"]![1]![0]!.GetValue<string>());
        Assert.Equal(80.0, data["values"]![1]![1]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Sequence_spills_a_deterministic_counter_grid()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("value", "=SEQUENCE(3,2)")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Null(Json(envelope)["meta"]!["warnings"]);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1:B3"))));
        // 1 2 / 3 4 / 5 6
        Assert.Equal(1.0, data["values"]![0]![0]!.GetValue<double>());
        Assert.Equal(2.0, data["values"]![0]![1]!.GetValue<double>());
        Assert.Equal(5.0, data["values"]![2]![0]!.GetValue<double>());
        Assert.Equal(6.0, data["values"]![2]![1]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Transpose_swaps_rows_and_columns()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("values", new JsonArray(
            new JsonArray(1, 2, 3))))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/A3", ("value", "=TRANSPOSE(A1:C1)"))).IsOk);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A3:A5"))));
        Assert.Equal(1.0, data["values"]![0]![0]!.GetValue<double>());
        Assert.Equal(2.0, data["values"]![1]![0]!.GetValue<double>());
        Assert.Equal(3.0, data["values"]![2]![0]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Textsplit_spills_a_delimited_string_across_columns()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("value", "=TEXTSPLIT(\"a,b,c\",\",\")")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Null(Json(envelope)["meta"]!["warnings"]);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1:C1"))));
        Assert.Equal("a", data["values"]![0]![0]!.GetValue<string>());
        Assert.Equal("b", data["values"]![0]![1]!.GetValue<string>());
        Assert.Equal("c", data["values"]![0]![2]!.GetValue<string>());

        // The anchor reports the array formula and the 1x3 spill rectangle.
        var anchor = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal("=TEXTSPLIT(\"a,b,c\",\",\")", anchor["formula"]!.GetValue<string>());
        Assert.Equal("A1:C1", anchor["spillRange"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Textsplit_with_a_row_delimiter_spills_a_grid()
    {
        var file = CreateWorkbook();
        // Column delimiter ",", row delimiter ";" -> a 2x2 grid.
        Assert.True(EditOps(
            file, SetOp("/Sheet1/A1", ("value", "=TEXTSPLIT(\"a,b;c,d\",\",\",\";\")"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1:B2"))));
        Assert.Equal("a", data["values"]![0]![0]!.GetValue<string>());
        Assert.Equal("b", data["values"]![0]![1]!.GetValue<string>());
        Assert.Equal("c", data["values"]![1]![0]!.GetValue<string>());
        Assert.Equal("d", data["values"]![1]![1]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Textsplit_reads_the_text_from_a_cell_reference()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "x|y|z"))).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=TEXTSPLIT(A1,\"|\")"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1:E1"))));
        Assert.Equal("x", data["values"]![0]![0]!.GetValue<string>());
        Assert.Equal("y", data["values"]![0]![1]!.GetValue<string>());
        Assert.Equal("z", data["values"]![0]![2]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Spill_blocked_when_the_target_range_is_occupied()
    {
        var file = SeedColumn("x", "y", "z");
        // C2 is in the way of a 3-row spill anchored at C1.
        Assert.True(EditOps(file, SetOp("/Sheet1/C2", ("value", "BLOCK"))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/C1", ("value", "=UNIQUE(A1:A3)")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.SpillBlocked, envelope.Error!.Code);
        Assert.Contains("C2", envelope.Error.Message, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion));

        // The edit was atomic: C2 still holds its original content (nothing spilled).
        var c2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C2"))));
        Assert.Equal("BLOCK", c2["value"]!.GetValue<string>());
    }

    [Fact]
    public void Spilled_cached_values_persist_across_reopen()
    {
        var file = SeedColumn(30, 10, 20, 10);
        Assert.True(EditOps(file, SetOp("/Sheet1/E1", ("value", "=SORT(A1:A4)"))).IsOk);

        // Reopen via a fresh get (the handler reopens the file from disk each call),
        // so this reads the cached values straight off disk, not from memory.
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/E1:E4"))));
        Assert.Equal(10.0, data["values"]![0]![0]!.GetValue<double>());
        Assert.Equal(10.0, data["values"]![1]![0]!.GetValue<double>());
        Assert.Equal(20.0, data["values"]![2]![0]!.GetValue<double>());
        Assert.Equal(30.0, data["values"]![3]![0]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Spill_survives_a_later_unrelated_edit_round_trip()
    {
        var file = SeedColumn("a", "b", "a", "c");
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=UNIQUE(A1:A4)"))).IsOk);

        // A second edit reopens the file through ClosedXML, mutates an unrelated
        // cell, and saves again — the spill (formula + cached values) must survive.
        Assert.True(EditOps(file, SetOp("/Sheet1/Z1", ("value", "untouched marker"))).IsOk);

        var range = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1:C3"))));
        Assert.Equal("a", range["values"]![0]![0]!.GetValue<string>());
        Assert.Equal("b", range["values"]![1]![0]!.GetValue<string>());
        Assert.Equal("c", range["values"]![2]![0]!.GetValue<string>());

        var anchor = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal("C1:C3", anchor["spillRange"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Randarray_spills_the_requested_shape_within_bounds()
    {
        var file = CreateWorkbook();
        // Deterministic seed (our implementation) keeps this test platform-stable;
        // we assert shape + bounds, never exact values.
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "=RANDARRAY(2,2,1,10,TRUE)"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1:B2"))));
        Assert.Equal(2, data["rows"]!.GetValue<int>());
        Assert.Equal(2, data["columns"]!.GetValue<int>());
        foreach (var row in data["values"]!.AsArray())
        {
            foreach (var cell in row!.AsArray())
            {
                var v = cell!.GetValue<double>();
                Assert.InRange(v, 1.0, 10.0);
                Assert.Equal(Math.Floor(v), v); // integer:TRUE
            }
        }

        AssertValidatorClean(file);
    }
}
