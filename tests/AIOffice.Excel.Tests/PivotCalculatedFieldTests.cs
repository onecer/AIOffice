using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.3 pivot calculated-fields surface (additive): the pivot op accepts
/// <c>calculatedFields:[{name, formula}]</c> and authors a calculated cacheField
/// (formula + databaseField=0), a matching pivotField/dataField and per-record
/// placeholder values so ClosedXML can still reopen the file. Raw assertions
/// reopen the file with the OpenXml SDK so the writer cannot grade its own
/// homework; every mutating test ends OpenXmlValidator-clean.
/// </summary>
public sealed class PivotCalculatedFieldTests : ExcelTestBase
{
    /// <summary>Region/Revenue/Cost data in A1:C4 on Sheet1.</summary>
    private string CreateSourceWorkbook()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1:C4", ("values", new JsonArray(
                new JsonArray("Region", "Revenue", "Cost"),
                new JsonArray("North", 100, 60),
                new JsonArray("South", 200, 120),
                new JsonArray("East", 150, 70)))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    private EditOp PivotOp(JsonArray calculatedFields) => new()
    {
        Op = "add",
        Path = "/Sheet1",
        Type = "pivot",
        Props = new JsonObject
        {
            ["name"] = "P1",
            ["sourceRange"] = "A1:C4",
            ["targetSheet"] = "Pivot",
            ["rows"] = new JsonArray("Region"),
            ["values"] = new JsonArray(new JsonObject { ["field"] = "Revenue", ["agg"] = "sum" }),
            ["calculatedFields"] = calculatedFields,
        },
    };

    private static S.PivotCacheDefinition CacheDef(SpreadsheetDocument document) =>
        document.WorkbookPart!.PivotTableCacheDefinitionParts.First().PivotCacheDefinition!;

    private static S.PivotTableDefinition TableDef(SpreadsheetDocument document) =>
        document.WorkbookPart!.WorksheetParts
            .SelectMany(p => p.PivotTableParts)
            .First().PivotTableDefinition!;

    [Fact]
    public void Calculated_field_writes_calc_cacheField_dataField_and_record_placeholders()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(new JsonArray(
            new JsonObject { ["name"] = "Margin", ["formula"] = "=Revenue-Cost" })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var details = OkData(envelope)["ops"]!.AsArray()[0]!;
        var calc = details["calculatedFields"]!.AsArray();
        Assert.Equal("Margin", calc[0]!["name"]!.GetValue<string>());
        Assert.Equal("Revenue-Cost", calc[0]!["formula"]!.GetValue<string>());

        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);

        // 1) A calculated cacheField with databaseField=0 + the formula.
        var calcField = Assert.Single(CacheDef(document).CacheFields!.Elements<S.CacheField>()
            .Where(f => f.DatabaseField?.Value == false).ToList());
        Assert.Equal("Margin", calcField.Name!.Value);
        Assert.Equal("Revenue-Cost", calcField.Formula!.Value);

        // 2) The dataField exposes it as a value column.
        var dataField = Assert.Single(TableDef(document).DataFields!.Elements<S.DataField>()
            .Where(d => d.Name?.Value == "Margin").ToList());
        Assert.Equal(3u, dataField.Field!.Value); // 0-based: Region,Revenue,Cost,Margin

        // 3) Field counts stay consistent and each record gained a placeholder.
        Assert.Equal(4u, CacheDef(document).CacheFields!.Count!.Value);
        Assert.Equal(4u, TableDef(document).PivotFields!.Count!.Value);
        var records = document.WorkbookPart!.PivotTableCacheDefinitionParts.First()
            .PivotTableCacheRecordsPart!.PivotCacheRecords!;
        Assert.All(records.Elements<S.PivotCacheRecord>(), r => Assert.Equal(4, r.ChildElements.Count));
    }

    [Fact]
    public void File_with_calculated_field_reopens_through_the_closedXml_edit_path()
    {
        var file = CreateSourceWorkbook();
        Assert.True(EditOps(file, PivotOp(new JsonArray(
            new JsonObject { ["name"] = "Margin", ["formula"] = "=Revenue-Cost" }))).IsOk);

        // A second edit op proves the file still opens through the ClosedXML DOM
        // path (the placeholder record values keep it loadable).
        var second = EditOps(file, SetOp("/Sheet1/E1", ("value", 1)));
        Assert.True(second.IsOk, second.ToJson());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Get_pivot_reports_calculated_fields_and_excludes_them_from_values()
    {
        var file = CreateSourceWorkbook();
        Assert.True(EditOps(file, PivotOp(new JsonArray(
            new JsonObject { ["name"] = "Margin", ["formula"] = "=Revenue-Cost" }))).IsOk);

        var envelope = Handler.Get(Ctx(file, ("path", "/Pivot/pivot[@name=P1]")));
        var data = OkData(envelope);

        var calc = data["calculatedFields"]!.AsArray();
        Assert.Single(calc);
        Assert.Equal("Margin", calc[0]!["name"]!.GetValue<string>());
        Assert.Equal("Revenue-Cost", calc[0]!["formula"]!.GetValue<string>());

        // The calculated field is NOT double-listed in the regular value fields.
        var values = data["values"]!.AsArray();
        Assert.Single(values);
        Assert.Equal("Revenue", values[0]!["field"]!.GetValue<string>());
    }

    [Fact]
    public void Multiple_calculated_fields_are_all_authored()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(new JsonArray(
            new JsonObject { ["name"] = "Margin", ["formula"] = "=Revenue-Cost" },
            new JsonObject { ["name"] = "Markup", ["formula"] = "=Revenue/Cost" })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var calcFields = CacheDef(document).CacheFields!.Elements<S.CacheField>()
            .Where(f => f.DatabaseField?.Value == false)
            .Select(f => f.Name!.Value)
            .ToList();
        Assert.Equal(["Margin", "Markup"], calcFields);
        Assert.Equal(5u, CacheDef(document).CacheFields!.Count!.Value); // 3 source + 2 calc
    }

    [Fact]
    public void Calculated_field_referencing_unknown_field_is_invalid_args_with_candidates()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(new JsonArray(
            new JsonObject { ["name"] = "Bad", ["formula"] = "=Revenue-Expenses" })));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("Expenses", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("Revenue", envelope.Error.Candidates!);
        Assert.Contains("Cost", envelope.Error.Candidates!);
        Assert.Contains("Region", envelope.Error.Candidates!);
    }

    [Fact]
    public void Calculated_field_colliding_with_source_name_is_invalid_args()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(new JsonArray(
            new JsonObject { ["name"] = "Revenue", ["formula"] = "=Revenue*2" })));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("collides", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Calculated_field_without_formula_is_invalid_args()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(new JsonArray(
            new JsonObject { ["name"] = "Margin" })));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("formula", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Calculated_field_with_function_call_is_accepted()
    {
        var file = CreateSourceWorkbook();

        // A function call (ROUND) is not a field reference; only the bare field
        // names inside it are validated.
        var envelope = EditOps(file, PivotOp(new JsonArray(
            new JsonObject { ["name"] = "RoundedMargin", ["formula"] = "=ROUND(Revenue-Cost,2)" })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var calcField = Assert.Single(CacheDef(document).CacheFields!.Elements<S.CacheField>()
            .Where(f => f.DatabaseField?.Value == false).ToList());
        Assert.Equal("ROUND(Revenue-Cost,2)", calcField.Formula!.Value);
    }

    [Fact]
    public void Pivot_without_calculated_fields_still_works_and_reports_empty_list()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, new EditOp
        {
            Op = "add",
            Path = "/Sheet1",
            Type = "pivot",
            Props = new JsonObject
            {
                ["name"] = "Plain",
                ["sourceRange"] = "A1:C4",
                ["targetSheet"] = "Pivot",
                ["rows"] = new JsonArray("Region"),
                ["values"] = new JsonArray(new JsonObject { ["field"] = "Revenue", ["agg"] = "sum" }),
            },
        });

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Pivot/pivot[@name=Plain]"))));
        Assert.Empty(data["calculatedFields"]!.AsArray());
    }
}
