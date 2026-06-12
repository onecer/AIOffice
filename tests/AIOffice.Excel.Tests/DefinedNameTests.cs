using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M3 defined-names slice: add (workbook or sheet scope), get/remove via
/// the stable <c>/name[@name=X]</c> form, structure listing, and the payoff —
/// formulas referencing the name evaluate through the ClosedXML engine, so the
/// saved file carries real cached values.
/// </summary>
public sealed class DefinedNameTests : ExcelTestBase
{
    private string CreateDataWorkbook()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:B2", ("values", new JsonArray(
                new JsonArray(1, 2),
                new JsonArray(3, 4))))).IsOk);
        return file;
    }

    [Fact]
    public void Add_workbook_scoped_name_then_get_returns_the_range()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:B2", "name", ("name", "SalesData")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal("/name[@name=SalesData]", detail["path"]!.GetValue<string>());
        Assert.Equal("workbook", detail["scope"]!.GetValue<string>());
        AssertValidatorClean(file);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/name[@name=SalesData]"))));
        Assert.Equal("name", data["kind"]!.GetValue<string>());
        Assert.Equal("SalesData", data["name"]!.GetValue<string>());
        Assert.Equal("workbook", data["scope"]!.GetValue<string>());
        Assert.Equal("/Sheet1/A1:B2", data["ranges"]![0]!.GetValue<string>());
        Assert.Contains("A$1", data["refersTo"]!.GetValue<string>(), StringComparison.Ordinal);

        // Raw oracle: the definedName element exists in workbook.xml.
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var definedName = document.WorkbookPart!.Workbook!
            .Descendants<S.DefinedName>().Single(n => n.Name?.Value == "SalesData");
        Assert.Null(definedName.LocalSheetId); // workbook scope
    }

    [Fact]
    public void Add_sheet_scoped_name_is_addressable_with_and_without_the_sheet_prefix()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(
            file, AddOp("/Sheet1/A1:A2", "name", ("name", "LocalData"), ("scope", "sheet"))).IsOk);
        AssertValidatorClean(file);

        var qualified = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/name[@name=LocalData]"))));
        Assert.Equal("sheet", qualified["scope"]!.GetValue<string>());
        Assert.Equal("Sheet1", qualified["sheet"]!.GetValue<string>());
        Assert.Equal("/Sheet1/name[@name=LocalData]", qualified["path"]!.GetValue<string>());

        // The bare form searches workbook scope first, then every sheet.
        var bare = OkData(Handler.Get(Ctx(file, ("path", "/name[@name=LocalData]"))));
        Assert.Equal("sheet", bare["scope"]!.GetValue<string>());

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var definedName = document.WorkbookPart!.Workbook!
            .Descendants<S.DefinedName>().Single(n => n.Name?.Value == "LocalData");
        Assert.NotNull(definedName.LocalSheetId); // sheet scope
    }

    [Fact]
    public void Formulas_can_reference_the_name_and_save_cached_values()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:B2", "name", ("name", "SalesData"))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/D1", ("value", "=SUM(SalesData)")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Null(Json(envelope)["meta"]!["warnings"]); // evaluated, not formula_not_evaluated
        AssertValidatorClean(file);

        // Raw oracle: formula text AND its cached value are on disk.
        var (formula, cached, _) = RawCell(file, "Sheet1", "D1");
        Assert.Equal("SUM(SalesData)", formula);
        Assert.Equal("10", cached);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D1"))));
        Assert.Equal(10.0, data["value"]!.GetValue<double>());
    }

    [Fact]
    public void Structure_view_lists_names_with_scope()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1/A1:B2", "name", ("name", "Global")),
            AddOp("/Sheet1/A1:A2", "name", ("name", "Local"), ("scope", "sheet"))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "structure"))));

        var names = data["definedNames"]!.AsArray();
        var global = names.Single(n => n!["name"]!.GetValue<string>() == "Global")!;
        Assert.Equal("workbook", global["scope"]!.GetValue<string>());
        Assert.Equal("/name[@name=Global]", global["path"]!.GetValue<string>());
        var local = names.Single(n => n!["name"]!.GetValue<string>() == "Local")!;
        Assert.Equal("sheet", local["scope"]!.GetValue<string>());
        Assert.Equal("Sheet1", local["sheet"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_deletes_the_name_and_leaves_the_file_clean()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:B2", "name", ("name", "Doomed"))).IsOk);

        var envelope = EditOps(file, RemoveOp("/name[@name=Doomed]"));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("name", Json(envelope)["data"]!["ops"]![0]!["removed"]!.GetValue<string>());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.DoesNotContain(
            document.WorkbookPart!.Workbook!.Descendants<S.DefinedName>(),
            n => n.Name?.Value == "Doomed");
    }

    [Fact]
    public void Get_missing_name_is_invalid_path_with_candidates()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:B2", "name", ("name", "SalesData"))).IsOk);

        var envelope = Handler.Get(Ctx(file, ("path", "/name[@name=SalesDta]")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/name[@name=SalesData]", envelope.Error.Candidates!);
    }

    [Fact]
    public void Duplicate_name_in_the_same_scope_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:B2", "name", ("name", "SalesData"))).IsOk);

        var envelope = EditOps(file, AddOp("/Sheet1/A1:A2", "name", ("name", "SalesData")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("already exists", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("1Bad")]
    [InlineData("has space")]
    [InlineData("A1")] // looks like a cell reference
    public void Invalid_names_are_rejected_with_the_naming_rules(string name)
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1/A1:B2", "name", ("name", name)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("cell reference", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Add_name_on_a_sheet_path_is_invalid_args_with_the_recipe()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "name", ("name", "SalesData")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("/Sheet1/A1:B5", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Set_on_a_name_is_a_typed_unsupported_feature()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:B2", "name", ("name", "SalesData"))).IsOk);

        var envelope = EditOps(file, SetOp("/name[@name=SalesData]", ("value", 1)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("remove", envelope.Error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }
}
