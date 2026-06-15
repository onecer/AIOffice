using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.7) Workbook calculation settings on the root set ({op:set, path:"/"}): the
/// calcPr mode/iteration/precision attributes, reopen-verified raw and reflected by
/// get /. A root set may carry both calc settings and structure protection.
/// </summary>
public sealed class CalculationModeTests : ExcelTestBase
{
    private static S.CalculationProperties? RawCalcPr(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        return document.WorkbookPart!.Workbook!.CalculationProperties;
    }

    [Fact]
    public void Calc_mode_and_iteration_settings_write_calcPr_and_get_reflects()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/",
            ("calculationMode", "manual"),
            ("iterativeCalc", true),
            ("maxIterations", 50),
            ("maxChange", 0.01),
            ("fullPrecision", false)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        var calcPr = RawCalcPr(file)!;
        Assert.Equal(S.CalculateModeValues.Manual, calcPr.CalculationMode!.Value);
        Assert.True(calcPr.Iterate!.Value);
        Assert.Equal(50u, calcPr.IterateCount!.Value);
        Assert.Equal(0.01, calcPr.IterateDelta!.Value);
        Assert.False(calcPr.FullPrecision!.Value);

        var calc = OkData(Handler.Get(Ctx(file, ("path", "/"))))["calculation"]!;
        Assert.Equal("manual", calc["calculationMode"]!.GetValue<string>());
        Assert.True(calc["iterativeCalc"]!.GetValue<bool>());
        Assert.Equal(50, calc["maxIterations"]!.GetValue<int>());
        Assert.Equal(0.01, calc["maxChange"]!.GetValue<double>());
        Assert.False(calc["fullPrecision"]!.GetValue<bool>());
    }

    [Theory]
    [InlineData("auto")]
    [InlineData("manual")]
    [InlineData("autoExceptTables")]
    public void Each_mode_maps_to_the_right_calcMode(string wire)
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/", ("calculationMode", wire))).IsOk);
        AssertValidatorClean(file);

        var calcPr = RawCalcPr(file)!;
        // "auto" is the OOXML default; ClosedXML may omit the attribute — get still reports it.
        var expected = wire switch
        {
            "manual" => S.CalculateModeValues.Manual,
            "autoExceptTables" => S.CalculateModeValues.AutoNoTable,
            _ => S.CalculateModeValues.Auto,
        };
        if (wire != "auto")
        {
            Assert.Equal(expected, calcPr.CalculationMode!.Value);
        }

        Assert.Equal(wire, OkData(Handler.Get(Ctx(file, ("path", "/"))))["calculation"]!["calculationMode"]!.GetValue<string>());
    }

    [Fact]
    public void Calc_settings_survive_an_unrelated_edit_and_reopen()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/", ("calculationMode", "manual"), ("iterativeCalc", true))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", 42))).IsOk);

        AssertValidatorClean(file);
        var calc = OkData(Handler.Get(Ctx(file, ("path", "/"))))["calculation"]!;
        Assert.Equal("manual", calc["calculationMode"]!.GetValue<string>());
        Assert.True(calc["iterativeCalc"]!.GetValue<bool>());
    }

    [Fact]
    public void Root_set_can_carry_both_calc_settings_and_structure_protection()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetOp("/",
            ("calculationMode", "manual"),
            ("protectStructure", true)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/"))));
        Assert.Equal("manual", data["calculation"]!["calculationMode"]!.GetValue<string>());
        Assert.NotNull(data["workbookProtection"]);
    }

    [Fact]
    public void Unknown_calc_mode_is_invalid_args_with_candidates()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(file, SetOp("/", ("calculationMode", "sometimes")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("auto", envelope.Error.Candidates!);
        Assert.Contains("manual", envelope.Error.Candidates!);
    }

    [Fact]
    public void Max_iterations_out_of_range_is_invalid_args()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(file, SetOp("/", ("maxIterations", 0)));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    [Fact]
    public void Negative_max_change_is_invalid_args()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(file, SetOp("/", ("maxChange", -1.0)));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    [Fact]
    public void Unknown_root_prop_is_invalid_args_with_candidates()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(file, SetOp("/", ("calcStyle", "fast")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("calculationMode", envelope.Error.Candidates!);
        Assert.Contains("protectStructure", envelope.Error.Candidates!);
    }

    [Fact]
    public void Get_root_on_a_fresh_workbook_reports_default_auto_or_null()
    {
        var file = CreateWorkbook();
        var data = OkData(Handler.Get(Ctx(file, ("path", "/"))));
        // A fresh workbook has no explicit calc settings (calculation may be null,
        // or auto when a calcPr exists); either way the call succeeds (it no longer
        // returns unsupported_feature).
        Assert.True(data["calculation"] is null || data["calculation"]!["calculationMode"]!.GetValue<string>() == "auto");
    }
}
