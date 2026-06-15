using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.3 advanced conditional-format kinds (additive, completing the rule
/// family): <c>formula</c> (an expression rule), <c>topBottom</c> (top/bottom N
/// or N%) and <c>aboveBelowAverage</c> (above/below the range average, optionally
/// ±N standard deviations). Raw assertions reopen the file with the OpenXml SDK
/// so the rule writer cannot grade its own homework; every mutating test ends
/// OpenXmlValidator-clean. The existing kinds are unchanged.
/// </summary>
public sealed class AdvancedConditionalFormatTests : ExcelTestBase
{
    /// <summary>Numbers 10,20,30,40,50 in A1:A5 plus a comparison column in B1:B5.</summary>
    private string CreateDataWorkbook()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1:B5", ("values", new JsonArray(
                new JsonArray(10, 15),
                new JsonArray(20, 5),
                new JsonArray(30, 35),
                new JsonArray(40, 10),
                new JsonArray(50, 60)))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    private static S.Worksheet RawSheet(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;
    }

    private static S.ConditionalFormattingRule SingleRule(SpreadsheetDocument document) =>
        Assert.Single(RawSheet(document).Descendants<S.ConditionalFormattingRule>().ToList());

    // ----- formula (expression) ----------------------------------------------

    [Fact]
    public void Formula_kind_writes_expression_cfRule_validator_clean()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "formula"), ("formula", "=$A1>$B1"), ("fill", "FFC7CE")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = SingleRule(document);
        Assert.Equal(S.ConditionalFormatValues.Expression, rule.Type!.Value);
        Assert.Equal("$A1>$B1", rule.GetFirstChild<S.Formula>()!.Text);
    }

    [Fact]
    public void Formula_kind_accepts_expression_without_leading_equals()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "formula"), ("formula", "MOD($A1,2)=0"), ("color", "9C0006"), ("bold", true)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Equal("MOD($A1,2)=0", SingleRule(document).GetFirstChild<S.Formula>()!.Text);
    }

    [Fact]
    public void Formula_kind_without_formula_is_invalid_args()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "formula"), ("fill", "FFC7CE")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("formula", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Formula_kind_without_style_is_invalid_args()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "formula"), ("formula", "=$A1>0")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("fill", envelope.Error.Message, StringComparison.Ordinal);
    }

    // ----- topBottom (top10) --------------------------------------------------

    [Fact]
    public void TopBottom_top_count_writes_top10_rule()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "topBottom"), ("mode", "top"), ("rank", 2), ("fill", "C6EFCE")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = SingleRule(document);
        Assert.Equal(S.ConditionalFormatValues.Top10, rule.Type!.Value);
        Assert.Equal(2u, rule.Rank!.Value);
        Assert.False(rule.Bottom?.Value ?? false);
        Assert.False(rule.Percent?.Value ?? false);
    }

    [Fact]
    public void TopBottom_bottom_percent_writes_percent_and_bottom_flags()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "topBottom"), ("mode", "bottom"), ("rank", 20), ("percent", true), ("fill", "FFC7CE")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = SingleRule(document);
        Assert.Equal(S.ConditionalFormatValues.Top10, rule.Type!.Value);
        Assert.Equal(20u, rule.Rank!.Value);
        Assert.True(rule.Bottom!.Value);
        Assert.True(rule.Percent!.Value);
    }

    [Fact]
    public void TopBottom_unknown_mode_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "topBottom"), ("mode", "middle"), ("rank", 3), ("fill", "FFC7CE")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("top", envelope.Error.Candidates!);
        Assert.Contains("bottom", envelope.Error.Candidates!);
    }

    [Fact]
    public void TopBottom_without_rank_is_invalid_args()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "topBottom"), ("mode", "top"), ("fill", "FFC7CE")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("rank", envelope.Error.Message, StringComparison.Ordinal);
    }

    // ----- aboveBelowAverage (raw-authored) -----------------------------------

    [Fact]
    public void AboveAverage_writes_aboveAverage_rule_validator_clean()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "aboveBelowAverage"), ("mode", "above"), ("fill", "C6EFCE")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = SingleRule(document);
        Assert.Equal(S.ConditionalFormatValues.AboveAverage, rule.Type!.Value);
        Assert.True(rule.AboveAverage!.Value);
        Assert.Null(rule.EqualAverage);
        Assert.Null(rule.StdDev);

        // A differential format (dxf) backs the rule's fill.
        var dxfs = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.GetFirstChild<S.DifferentialFormats>();
        Assert.NotNull(dxfs);
        Assert.True(dxfs!.Elements<S.DifferentialFormat>().Any());
        Assert.Equal(rule.FormatId!.Value, dxfs.Count!.Value - 1u);
    }

    [Fact]
    public void BelowOrEqualAverage_sets_aboveAverage_false_and_equalAverage_true()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "aboveBelowAverage"), ("mode", "belowOrEqual"), ("fill", "FFC7CE")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var rule = SingleRule(document);
        Assert.False(rule.AboveAverage!.Value);
        Assert.True(rule.EqualAverage!.Value);
    }

    [Fact]
    public void AboveAverage_with_stdDev_writes_the_stdDev_attribute()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "aboveBelowAverage"), ("mode", "above"), ("stdDev", 1), ("fill", "C6EFCE")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Equal(1, SingleRule(document).StdDev!.Value);
    }

    [Fact]
    public void AboveAverage_reopens_in_closedXml_and_get_reports_mode_and_stdDev()
    {
        var file = CreateDataWorkbook();

        Assert.True(EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "aboveBelowAverage"), ("mode", "above"), ("stdDev", 2), ("fill", "C6EFCE"))).IsOk);

        // A second edit op proves the ClosedXML DOM path still opens the file
        // (the raw-authored aboveAverage rule does not break the round trip).
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", 99))).IsOk);
        AssertValidatorClean(file);

        var envelope = Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]")));
        var data = OkData(envelope);
        Assert.Equal("aboveBelowAverage", data["cfKind"]!.GetValue<string>());
        Assert.Equal("above", data["mode"]!.GetValue<string>());
        Assert.Equal(2, data["stdDev"]!.GetValue<int>());
    }

    [Fact]
    public void AboveAverage_unknown_mode_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "aboveBelowAverage"), ("mode", "sideways"), ("fill", "FFC7CE")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("aboveOrEqual", envelope.Error.Candidates!);
    }

    [Fact]
    public void AboveAverage_negative_stdDev_is_invalid_args()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A5", "conditionalFormat",
            ("kind", "aboveBelowAverage"), ("mode", "above"), ("stdDev", -1), ("fill", "FFC7CE")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("stdDev", envelope.Error.Message, StringComparison.Ordinal);
    }

    // ----- get / describe for the native new kinds ----------------------------

    [Fact]
    public void Get_reports_formula_and_topBottom_details()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file,
            AddOp("/Sheet1/A1:A5", "conditionalFormat", ("kind", "formula"), ("formula", "=$A1>25"), ("fill", "FFC7CE")),
            AddOp("/Sheet1/A1:A5", "conditionalFormat", ("kind", "topBottom"), ("mode", "top"), ("rank", 3), ("fill", "C6EFCE"))).IsOk);

        var formula = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("formula", formula["cfKind"]!.GetValue<string>());
        Assert.Equal("$A1>25", formula["formula"]!.GetValue<string>());

        var topBottom = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[2]"))));
        Assert.Equal("topBottom", topBottom["cfKind"]!.GetValue<string>());
        Assert.Equal("top", topBottom["mode"]!.GetValue<string>());
        Assert.Equal(3, topBottom["rank"]!.GetValue<int>());
    }

    [Fact]
    public void Average_rule_index_follows_the_native_rules_in_a_mixed_batch()
    {
        var file = CreateDataWorkbook();

        // A native rule then an average rule: the average rule is authored last,
        // so it is conditionalFormat[2].
        var envelope = EditOps(file,
            AddOp("/Sheet1/A1:A5", "conditionalFormat", ("kind", "cellIs"), ("operator", ">"), ("value", 25), ("fill", "FFC7CE")),
            AddOp("/Sheet1/A1:A5", "conditionalFormat", ("kind", "aboveBelowAverage"), ("mode", "above"), ("fill", "C6EFCE")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var ops = OkData(envelope)["ops"]!.AsArray();
        Assert.Equal("/Sheet1/conditionalFormat[1]", ops[0]!["path"]!.GetValue<string>());
        Assert.Equal("/Sheet1/conditionalFormat[2]", ops[1]!["path"]!.GetValue<string>());

        AssertValidatorClean(file);

        var cellIs = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[1]"))));
        Assert.Equal("cellIs", cellIs["cfKind"]!.GetValue<string>());
        var average = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/conditionalFormat[2]"))));
        Assert.Equal("aboveBelowAverage", average["cfKind"]!.GetValue<string>());
    }
}
