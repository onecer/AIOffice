using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.4) What-if data tables: {op:add, type:dataTable, path:/Sheet1/A1:C10,
/// props:{rowInput, colInput?}} recomputes the corner formula across the row/
/// column input axes and fills the body with cached results, carrying the Excel
/// {=TABLE(...)} construct (validator-clean). get/remove address /Sheet/dataTable[i].
/// </summary>
public sealed class DataTableTests : ExcelTestBase
{
    [Fact]
    public void Two_variable_table_fills_the_body_with_computed_values()
    {
        var file = CreateWorkbook();
        // Corner A1 = B1 * B2 (B1 row input, B2 col input).
        // Row axis across row 1: B1=2, C1=3. Column axis down column A: A2=10, A3=20.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B1", ("value", 2)),
            SetOp("/Sheet1/B2", ("value", 10)),
            SetOp("/Sheet1/A1", ("value", "=B1*B2")),
            SetOp("/Sheet1/C1", ("value", 3)),
            SetOp("/Sheet1/A2", ("value", 10)),
            SetOp("/Sheet1/A3", ("value", 20))).IsOk);

        var envelope = EditOps(
            file,
            AddOp("/Sheet1/A1:C3", "dataTable", ("rowInput", "B1"), ("colInput", "B2")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("dataTable", OkData(envelope)["ops"]![0]!["type"]!.GetValue<string>());

        // Body B2:C3 = rowAxis * colAxis: [2*10, 3*10] / [2*20, 3*20] = [20,30]/[40,60].
        var body = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2:C3"))));
        Assert.Equal(20.0, body["values"]![0]![0]!.GetValue<double>());
        Assert.Equal(30.0, body["values"]![0]![1]!.GetValue<double>());
        Assert.Equal(40.0, body["values"]![1]![0]!.GetValue<double>());
        Assert.Equal(60.0, body["values"]![1]![1]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void One_variable_column_table_recomputes_down_the_axis()
    {
        var file = CreateWorkbook();
        // Corner A1 = B1 * 100; colInput = B1; column A holds the substitutions.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B1", ("value", 0)),
            SetOp("/Sheet1/A1", ("value", "=B1*100")),
            SetOp("/Sheet1/A2", ("value", 1)),
            SetOp("/Sheet1/A3", ("value", 2)),
            SetOp("/Sheet1/A4", ("value", 5))).IsOk);

        Assert.True(EditOps(file, AddOp("/Sheet1/A1:B4", "dataTable", ("colInput", "B1"))).IsOk);

        // Body B2:B4 = colAxis * 100 = 100, 200, 500.
        var body = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2:B4"))));
        Assert.Equal(100.0, body["values"]![0]![0]!.GetValue<double>());
        Assert.Equal(200.0, body["values"]![1]![0]!.GetValue<double>());
        Assert.Equal(500.0, body["values"]![2]![0]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Loan_payment_sensitivity_table_computes_real_pmt_values()
    {
        var file = CreateWorkbook();
        // The canonical what-if: a loan payment table whose INPUT cells (E1=rate,
        // E2=term) live OUTSIDE the table range, as real data tables require.
        // Corner A1 = PMT(E1/12, E2, -100000).
        // Row axis (terms across row 1): B1=360, C1=180.
        // Column axis (rates down column A): A2=0.05, A3=0.06.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/E1", ("value", 0.05)),
            SetOp("/Sheet1/E2", ("value", 360)),
            SetOp("/Sheet1/A1", ("value", "=PMT(E1/12,E2,-100000)")),
            SetOp("/Sheet1/B1", ("value", 360)),
            SetOp("/Sheet1/C1", ("value", 180)),
            SetOp("/Sheet1/A2", ("value", 0.05)),
            SetOp("/Sheet1/A3", ("value", 0.06))).IsOk);

        Assert.True(EditOps(
            file,
            AddOp("/Sheet1/A1:C3", "dataTable", ("rowInput", "E2"), ("colInput", "E1"))).IsOk);

        var body = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2:C3"))));
        // [rate 0.05]: 360mo -> 536.82, 180mo -> 790.79
        // [rate 0.06]: 360mo -> 599.55, 180mo -> 843.86
        Assert.Equal(536.82, body["values"]![0]![0]!.GetValue<double>(), 1);
        Assert.Equal(790.79, body["values"]![0]![1]!.GetValue<double>(), 1);
        Assert.Equal(599.55, body["values"]![1]![0]!.GetValue<double>(), 1);
        Assert.Equal(843.86, body["values"]![1]![1]!.GetValue<double>(), 1);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Get_reports_the_data_table_metadata()
    {
        var file = SeedSimpleTable();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "dataTable", ("rowInput", "B1"), ("colInput", "B2"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/dataTable[1]"))));
        Assert.Equal("dataTable", data["kind"]!.GetValue<string>());
        Assert.Equal("B1", data["rowInput"]!.GetValue<string>());
        Assert.Equal("B2", data["colInput"]!.GetValue<string>());
        Assert.True(data["twoDimensional"]!.GetValue<bool>());
        Assert.Equal("B2:C3", data["body"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_clears_the_table_body_and_construct()
    {
        var file = SeedSimpleTable();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:C3", "dataTable", ("rowInput", "B1"), ("colInput", "B2"))).IsOk);

        var removeEnv = EditOps(file, RemoveOp("/Sheet1/dataTable[1]"));
        Assert.True(removeEnv.IsOk, removeEnv.ToJson());

        // The table is gone: get on dataTable[1] now fails.
        var envelope = Handler.Get(Ctx(file, ("path", "/Sheet1/dataTable[1]")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Add_data_table_on_too_small_a_range_is_rejected()
    {
        var file = SeedSimpleTable();
        var envelope = EditOps(file, AddOp("/Sheet1/A1", "dataTable", ("rowInput", "B1")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion));
    }

    [Fact]
    public void Add_data_table_without_an_input_cell_is_rejected()
    {
        var file = SeedSimpleTable();
        var envelope = EditOps(file, AddOp("/Sheet1/A1:C3", "dataTable"));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    /// <summary>
    /// A two-variable table: corner A1 = B1 * B2; row axis (B1,C1) = (2,3);
    /// column axis (A2,A3) = (10,20). Body should compute to a 2x2 product grid.
    /// </summary>
    private string SeedSimpleTable()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B1", ("value", 2)),
            SetOp("/Sheet1/B2", ("value", 10)),
            SetOp("/Sheet1/A1", ("value", "=B1*B2")),
            SetOp("/Sheet1/C1", ("value", 3)),
            SetOp("/Sheet1/A2", ("value", 10)),
            SetOp("/Sheet1/A3", ("value", 20))).IsOk);
        return file;
    }
}
