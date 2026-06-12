using System.IO.Compression;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M5 data validations — every kind is added through the op surface, then
/// verified by REOPENING the file (ClosedXML and raw XML, never our own get
/// alone) plus the schema validator, so an Excel open is known-good.
/// </summary>
public sealed class DataValidationTests : ExcelTestBase
{
    private static EditOp DvOp(string path, params (string Key, JsonNode? Value)[] props) =>
        AddOp(path, "dataValidation", props);

    private static string RawSheetXml(string file)
    {
        using var zip = ZipFile.OpenRead(file);
        using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    [Fact]
    public void List_rule_with_literal_values_reopens_as_a_dropdown()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(file, DvOp(
            "/Sheet1/B2:B50",
            ("kind", "list"),
            ("values", new JsonArray("待办", "进行中", "完成")),
            ("allowBlank", true),
            ("inputTitle", "Status"),
            ("inputMessage", "Pick one"),
            ("errorTitle", "Invalid"),
            ("errorMessage", "Choose from the list"),
            ("errorStyle", "warning")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal("/Sheet1/dataValidation[1]", detail["path"]!.GetValue<string>());
        Assert.Equal("list", detail["dvKind"]!.GetValue<string>());

        using (var workbook = new XLWorkbook(file))
        {
            var dv = Assert.Single(workbook.Worksheet("Sheet1").DataValidations);
            Assert.Equal(XLAllowedValues.List, dv.AllowedValues);
            Assert.Equal("\"待办,进行中,完成\"", dv.Value);
            Assert.True(dv.InCellDropdown);
            Assert.True(dv.IgnoreBlanks);
            Assert.Equal("Status", dv.InputTitle);
            Assert.Equal("Choose from the list", dv.ErrorMessage);
            Assert.Equal(XLErrorStyle.Warning, dv.ErrorStyle);
            Assert.Equal("B2:B50", dv.Ranges.Single().RangeAddress.ToString());
        }

        Assert.Contains("dataValidation type=\"list\"", RawSheetXml(file), StringComparison.Ordinal);
        AssertValidatorClean(file);

        var got = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/dataValidation[1]"))));
        Assert.Equal("list", got["dvKind"]!.GetValue<string>());
        Assert.Equal(["待办", "进行中", "完成"], got["values"]!.AsArray().Select(v => v!.GetValue<string>()));
        Assert.Equal("warning", got["errorStyle"]!.GetValue<string>());
    }

    [Fact]
    public void List_rule_with_a_source_range_references_the_cells()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(
            file,
            AddOp("/Lists", "sheet"),
            SetOp("/Lists/A1:A3", ("values", new JsonArray(
                new JsonArray("red"), new JsonArray("green"), new JsonArray("blue")))),
            DvOp("/Sheet1/C2:C20", ("kind", "list"), ("sourceRange", "/Lists/A1:A3")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        using (var workbook = new XLWorkbook(file))
        {
            var dv = Assert.Single(workbook.Worksheet("Sheet1").DataValidations);
            Assert.Equal(XLAllowedValues.List, dv.AllowedValues);
            Assert.Contains("Lists", dv.Value, StringComparison.Ordinal);
            Assert.Contains("A1:A3", dv.Value.Replace("$", string.Empty), StringComparison.Ordinal);
        }

        AssertValidatorClean(file);
        var got = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/dataValidation[1]"))));
        Assert.NotNull(got["sourceRange"]);
        Assert.Null(got["values"]);
    }

    [Fact]
    public void WholeNumber_between_reopens_with_both_bounds()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, DvOp(
            "/Sheet1/A1:A10",
            ("kind", "wholeNumber"), ("operator", "between"), ("value", 1), ("value2", 10))).IsOk);

        using (var workbook = new XLWorkbook(file))
        {
            var dv = Assert.Single(workbook.Worksheet("Sheet1").DataValidations);
            Assert.Equal(XLAllowedValues.WholeNumber, dv.AllowedValues);
            Assert.Equal(XLOperator.Between, dv.Operator);
            Assert.Equal("1", dv.MinValue);
            Assert.Equal("10", dv.MaxValue);
        }

        AssertValidatorClean(file);
        var got = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/dataValidation[1]"))));
        Assert.Equal("wholeNumber", got["dvKind"]!.GetValue<string>());
        Assert.Equal("between", got["operator"]!.GetValue<string>());
        Assert.Equal("1", got["value"]!.GetValue<string>());
        Assert.Equal("10", got["value2"]!.GetValue<string>());
    }

    [Fact]
    public void Decimal_with_a_single_operand_operator_reopens()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, DvOp(
            "/Sheet1/B1:B5", ("kind", "decimal"), ("operator", ">="), ("value", 2.5))).IsOk);

        using (var workbook = new XLWorkbook(file))
        {
            var dv = Assert.Single(workbook.Worksheet("Sheet1").DataValidations);
            Assert.Equal(XLAllowedValues.Decimal, dv.AllowedValues);
            Assert.Equal(XLOperator.EqualOrGreaterThan, dv.Operator);
            Assert.Equal("2.5", dv.MinValue);
        }

        AssertValidatorClean(file);
        var got = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/dataValidation[1]"))));
        Assert.Equal(">=", got["operator"]!.GetValue<string>());
        Assert.Null(got["value2"]);
    }

    [Fact]
    public void Date_rule_stores_serials_but_reports_iso_dates()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, DvOp(
            "/Sheet1/D1:D31",
            ("kind", "date"), ("operator", "between"), ("value", "2026-01-01"), ("value2", "2026-12-31"))).IsOk);

        using (var workbook = new XLWorkbook(file))
        {
            var dv = Assert.Single(workbook.Worksheet("Sheet1").DataValidations);
            Assert.Equal(XLAllowedValues.Date, dv.AllowedValues);
            Assert.Equal(
                new DateTime(2026, 1, 1).ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture),
                dv.MinValue);
        }

        AssertValidatorClean(file);
        var got = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/dataValidation[1]"))));
        Assert.Equal("date", got["dvKind"]!.GetValue<string>());
        Assert.Equal("2026-01-01", got["value"]!.GetValue<string>());
        Assert.Equal("2026-12-31", got["value2"]!.GetValue<string>());
    }

    [Fact]
    public void TextLength_rule_reopens()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, DvOp(
            "/Sheet1/E1:E9", ("kind", "textLength"), ("operator", "<="), ("value", 10))).IsOk);

        using (var workbook = new XLWorkbook(file))
        {
            var dv = Assert.Single(workbook.Worksheet("Sheet1").DataValidations);
            Assert.Equal(XLAllowedValues.TextLength, dv.AllowedValues);
            Assert.Equal(XLOperator.EqualOrLessThan, dv.Operator);
            Assert.Equal("10", dv.MinValue);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Structure_lists_rules_and_remove_deletes_them()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            DvOp("/Sheet1/A1:A5", ("kind", "wholeNumber"), ("operator", ">"), ("value", 0)),
            DvOp("/Sheet1/B1:B5", ("kind", "list"), ("values", new JsonArray("x", "y")))).IsOk);

        var structure = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        var rules = structure["sheets"]![0]!["dataValidations"]!.AsArray();
        Assert.Equal(2, rules.Count);
        Assert.Equal("/Sheet1/dataValidation[1]", rules[0]!["path"]!.GetValue<string>());
        Assert.Equal("wholeNumber", rules[0]!["dvKind"]!.GetValue<string>());
        Assert.Equal("list", rules[1]!["dvKind"]!.GetValue<string>());

        var removed = EditOps(file, RemoveOp("/Sheet1/dataValidation[1]"));
        Assert.True(removed.IsOk, removed.ToJson());
        Assert.Equal("dataValidation", Json(removed)["data"]!["ops"]![0]!["removed"]!.GetValue<string>());

        using (var workbook = new XLWorkbook(file))
        {
            var dv = Assert.Single(workbook.Worksheet("Sheet1").DataValidations);
            Assert.Equal(XLAllowedValues.List, dv.AllowedValues); // the wholeNumber rule is gone
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Bad_ops_fail_with_typed_errors_and_candidates()
    {
        var file = CreateWorkbook();

        var unknownKind = EditOps(file, DvOp("/Sheet1/A1:A5", ("kind", "iconSet")));
        Assert.False(unknownKind.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, unknownKind.Error!.Code);
        Assert.Contains("list", unknownKind.Error.Candidates!);

        var bothSources = EditOps(file, DvOp(
            "/Sheet1/A1:A5", ("kind", "list"), ("values", new JsonArray("a")), ("sourceRange", "B1:B2")));
        Assert.False(bothSources.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, bothSources.Error!.Code);

        var commaValue = EditOps(file, DvOp(
            "/Sheet1/A1:A5", ("kind", "list"), ("values", new JsonArray("a,b"))));
        Assert.False(commaValue.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, commaValue.Error!.Code);
        Assert.Contains("sourceRange", commaValue.Error.Suggestion, StringComparison.Ordinal);

        var missingValue2 = EditOps(file, DvOp(
            "/Sheet1/A1:A5", ("kind", "wholeNumber"), ("operator", "between"), ("value", 1)));
        Assert.False(missingValue2.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, missingValue2.Error!.Code);

        var badStyle = EditOps(file, DvOp(
            "/Sheet1/A1:A5", ("kind", "wholeNumber"), ("operator", ">"), ("value", 1), ("errorStyle", "loud")));
        Assert.False(badStyle.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, badStyle.Error!.Code);
        Assert.Contains("stop", badStyle.Error.Candidates!);

        var unknownProp = EditOps(file, DvOp(
            "/Sheet1/A1:A5", ("kind", "wholeNumber"), ("operator", ">"), ("value", 1), ("colour", "red")));
        Assert.False(unknownProp.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, unknownProp.Error!.Code);

        var sheetTarget = EditOps(file, DvOp("/Sheet1", ("kind", "list"), ("values", new JsonArray("a"))));
        Assert.False(sheetTarget.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, sheetTarget.Error!.Code);

        var missingIndex = Handler.Get(Ctx(file, ("path", "/Sheet1/dataValidation[3]")));
        Assert.False(missingIndex.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, missingIndex.Error!.Code);

        var setOnRule = EditOps(file, SetOp("/Sheet1/dataValidation[1]", ("value", 1)));
        Assert.False(setOnRule.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, setOnRule.Error!.Code);
    }
}
