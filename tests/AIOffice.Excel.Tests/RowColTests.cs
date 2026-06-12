using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M4 row/column structure ops: insert (before/after) and delete with formula
/// references shifting, plus sizing (height/width) and hidden — all verified
/// by reopening the saved file.
/// </summary>
public sealed class RowColTests : ExcelTestBase
{
    private static EditOp InsertOp(string path, string type, string? position = null) =>
        new() { Op = "add", Path = path, Type = type, Position = position };

    private string CreateSummedColumn()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1)),
            SetOp("/Sheet1/A2", ("value", 2)),
            SetOp("/Sheet1/A3", ("value", 3)),
            SetOp("/Sheet1/A4", ("value", "=SUM(A1:A3)"))).IsOk);
        Assert.Equal("6", RawCell(file, "Sheet1", "A4").CachedValue);
        return file;
    }

    [Fact]
    public void Insert_row_before_shifts_formula_references()
    {
        var file = CreateSummedColumn();

        var envelope = EditOps(file, InsertOp("/Sheet1/row[2]", "row"));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("/Sheet1/row[2]", Json(envelope)["data"]!["ops"]![0]!["path"]!.GetValue<string>());

        // The =SUM moved down a row and widened: ClosedXML rewrote the references.
        var raw = RawCell(file, "Sheet1", "A5");
        Assert.Equal("SUM(A1:A4)", raw.Formula);
        Assert.Equal("6", raw.CachedValue);
        Assert.Equal("blank", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A2"))))["type"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Insert_row_after_lands_below_the_anchor()
    {
        var file = CreateSummedColumn();

        var envelope = EditOps(file, new EditOp
        {
            Op = "add",
            Path = "/Sheet1/row[1]",
            Type = "row",
            Position = "after",
            Props = new JsonObject { ["values"] = new JsonArray(99) },
        });

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("/Sheet1/row[2]", Json(envelope)["data"]!["ops"]![0]!["path"]!.GetValue<string>());
        Assert.Equal(99.0, OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A2"))))["value"]!.GetValue<double>());
        Assert.Equal(2.0, OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A3"))))["value"]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Delete_row_shrinks_formula_references()
    {
        var file = CreateSummedColumn();

        var envelope = EditOps(file, RemoveOp("/Sheet1/row[2]"));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var raw = RawCell(file, "Sheet1", "A3");
        Assert.Equal("SUM(A1:A2)", raw.Formula);
        Assert.Equal("4", raw.CachedValue); // 1 + 3 after the 2 went away
        AssertValidatorClean(file);
    }

    [Fact]
    public void Insert_and_delete_column_shift_formula_references()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1)),
            SetOp("/Sheet1/B1", ("value", 2)),
            SetOp("/Sheet1/C1", ("value", "=SUM(A1:B1)"))).IsOk);

        var inserted = EditOps(file, InsertOp("/Sheet1/col[B]", "col"));
        Assert.True(inserted.IsOk, inserted.ToJson());
        Assert.Equal("/Sheet1/col[B]", Json(inserted)["data"]!["ops"]![0]!["path"]!.GetValue<string>());

        var widened = RawCell(file, "Sheet1", "D1");
        Assert.Equal("SUM(A1:C1)", widened.Formula);
        Assert.Equal("3", widened.CachedValue);

        var removed = EditOps(file, RemoveOp("/Sheet1/col[B]"));
        Assert.True(removed.IsOk, removed.ToJson());
        var narrowed = RawCell(file, "Sheet1", "C1");
        Assert.Equal("SUM(A1:B1)", narrowed.Formula);
        Assert.Equal("3", narrowed.CachedValue);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Insert_column_after_via_letter_path()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "a")),
            SetOp("/Sheet1/B1", ("value", "b"))).IsOk);

        var envelope = EditOps(file, InsertOp("/Sheet1/col[A]", "col", "after"));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("/Sheet1/col[B]", Json(envelope)["data"]!["ops"]![0]!["path"]!.GetValue<string>());
        Assert.Equal("a", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<string>());
        Assert.Equal("blank", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))))["type"]!.GetValue<string>());
        Assert.Equal("b", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))))["value"]!.GetValue<string>());
    }

    [Fact]
    public void Row_height_and_hidden_survive_a_reopen()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A2", ("value", "x"))).IsOk);

        var envelope = EditOps(
            file,
            SetOp("/Sheet1/row[2]", ("height", 30)),
            SetOp("/Sheet1/row[3]", ("hidden", true)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        using (var workbook = new XLWorkbook(file))
        {
            var sheet = workbook.Worksheet("Sheet1");
            Assert.Equal(30, sheet.Row(2).Height);
            Assert.True(sheet.Row(3).IsHidden);
            Assert.False(sheet.Row(2).IsHidden);
        }

        var row = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/row[2]"))));
        Assert.Equal(30.0, row["height"]!.GetValue<double>());
        Assert.Null(row["hidden"]);
        Assert.True(OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/row[3]"))))["hidden"]!.GetValue<bool>());

        // Unhide round-trips too.
        Assert.True(EditOps(file, SetOp("/Sheet1/row[3]", ("hidden", false))).IsOk);
        using (var workbook = new XLWorkbook(file))
        {
            Assert.False(workbook.Worksheet("Sheet1").Row(3).IsHidden);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Column_width_and_hidden_survive_a_reopen()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "x"))).IsOk);

        var envelope = EditOps(
            file,
            SetOp("/Sheet1/col[C]", ("width", 18)),
            SetOp("/Sheet1/col[D]", ("hidden", true)));
        Assert.True(envelope.IsOk, envelope.ToJson());

        using (var workbook = new XLWorkbook(file))
        {
            var sheet = workbook.Worksheet("Sheet1");
            Assert.Equal(18, sheet.Column(3).Width, 0.5); // ClosedXML may round-trip via Excel's unit math
            Assert.True(sheet.Column(4).IsHidden);
        }

        var column = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/col[C]"))));
        Assert.Equal("/Sheet1/col[C]", column["path"]!.GetValue<string>());
        Assert.Equal("C", column["column"]!.GetValue<string>());
        Assert.Equal(18.0, column["width"]!.GetValue<double>(), 0.5);
        Assert.True(OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/col[D]"))))["hidden"]!.GetValue<bool>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Numeric_column_index_is_accepted_but_letter_form_is_canonical()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/col[2]", ("width", 12))).IsOk);

        var column = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/col[B]"))));
        Assert.Equal("/Sheet1/col[B]", column["path"]!.GetValue<string>());
        Assert.Equal(12.0, column["width"]!.GetValue<double>(), 0.5);
    }

    [Fact]
    public void Sizing_props_on_the_wrong_axis_are_rejected()
    {
        var file = CreateWorkbook();

        var heightOnCol = EditOps(file, SetOp("/Sheet1/col[C]", ("height", 20)));
        Assert.False(heightOnCol.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, heightOnCol.Error!.Code);
        Assert.Contains("width", heightOnCol.Error.Suggestion, StringComparison.Ordinal);

        var widthOnRow = EditOps(file, SetOp("/Sheet1/row[2]", ("width", 20)));
        Assert.False(widthOnRow.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, widthOnRow.Error!.Code);

        var hiddenOnCell = EditOps(file, SetOp("/Sheet1/A1", ("hidden", true)));
        Assert.False(hiddenOnCell.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, hiddenOnCell.Error!.Code);
    }

    [Fact]
    public void Out_of_range_sizes_are_rejected()
    {
        var file = CreateWorkbook();

        var tooTall = EditOps(file, SetOp("/Sheet1/row[1]", ("height", 500)));
        Assert.False(tooTall.IsOk);
        Assert.Contains("409", tooTall.Error!.Message, StringComparison.Ordinal);

        var tooWide = EditOps(file, SetOp("/Sheet1/col[A]", ("width", 300)));
        Assert.False(tooWide.IsOk);
        Assert.Contains("255", tooWide.Error!.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_insert_position_is_rejected_with_candidates()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, InsertOp("/Sheet1/col[A]", "col", "sideways"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("before", envelope.Error.Candidates!);
        Assert.Contains("after", envelope.Error.Candidates!);
    }

    [Fact]
    public void Add_col_on_a_non_column_path_is_rejected()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, InsertOp("/Sheet1", "col"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("col[C]", envelope.Error.Message, StringComparison.Ordinal);
    }
}
