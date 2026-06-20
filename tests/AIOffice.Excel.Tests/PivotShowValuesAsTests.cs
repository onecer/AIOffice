using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The v1.13 pivot show-values-as surface. A value field accepts
/// <c>showAs</c> (+ <c>baseField</c>/<c>baseItem</c> where the mode needs it) and it
/// is now APPLIED — the latent bug (a <c>showAs</c> silently accepted and never
/// written, so the pivot still showed raw sums) is fixed: the dataField carries the
/// <c>showDataAs</c> setting so Excel computes it on open, an unknown showAs is
/// <c>invalid_args</c>, and a mode OOXML cannot carry is <c>unsupported_feature</c>.
/// Raw assertions reopen the file with the OpenXml SDK so the writer cannot grade
/// its own homework; every mutating test ends OpenXmlValidator-clean.
/// </summary>
public sealed class PivotShowValuesAsTests : ExcelTestBase
{
    /// <summary>Region/Month/Sales data in A1:C7 on Sheet1 — two regions × repeating months.</summary>
    private string CreateSourceWorkbook()
    {
        var file = CreateWorkbook();
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1:C7", ("values", new JsonArray(
                new JsonArray("Region", "Month", "Sales"),
                new JsonArray("North", "Jan", 100),
                new JsonArray("North", "Feb", 150),
                new JsonArray("North", "Mar", 200),
                new JsonArray("South", "Jan", 80),
                new JsonArray("South", "Feb", 120),
                new JsonArray("South", "Mar", 160)))));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    /// <summary>An add-pivot op whose single value field carries the given showAs props.</summary>
    private EditOp PivotOp(JsonObject valueEntry, string name = "P1") => new()
    {
        Op = "add",
        Path = "/Sheet1",
        Type = "pivot",
        Props = new JsonObject
        {
            ["name"] = name,
            ["sourceRange"] = "A1:C7",
            ["targetSheet"] = "Pivot",
            ["rows"] = new JsonArray("Region"),
            ["columns"] = new JsonArray("Month"),
            ["values"] = new JsonArray(valueEntry),
        },
    };

    private static JsonObject Value(string field, string agg, params (string Key, JsonNode? Value)[] extra)
    {
        var obj = new JsonObject { ["field"] = field, ["agg"] = agg };
        foreach (var (key, value) in extra)
        {
            obj[key] = value;
        }

        return obj;
    }

    private static S.DataField FirstDataField(SpreadsheetDocument document) =>
        document.WorkbookPart!.WorksheetParts
            .SelectMany(p => p.PivotTableParts)
            .First().PivotTableDefinition!.DataFields!.Elements<S.DataField>().First();

    // ----- the modes carry showDataAs onto the bytes --------------------------

    [Theory]
    [InlineData("percentOfTotal", "percentOfTotal")]
    [InlineData("percentOfColumn", "percentOfCol")]
    [InlineData("percentOfRow", "percentOfRow")]
    [InlineData("index", "index")]
    public void Flat_mode_writes_showDataAs_onto_the_dataField(string showAs, string ooxml)
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(Value("Sales", "sum", ("showAs", showAs))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = OkData(envelope)["ops"]!.AsArray()[0]!["values"]!.AsArray()[0]!;
        Assert.Equal(showAs, detail["showAs"]!.GetValue<string>());

        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var dataField = FirstDataField(document);
        Assert.NotNull(dataField.ShowDataAs);
        Assert.Equal(ooxml, dataField.ShowDataAs!.InnerText);
        // A flat mode carries no base.
        Assert.Null(dataField.BaseField);
    }

    [Fact]
    public void RunningTotal_writes_runTotal_and_resolves_baseField_to_its_cacheField_index()
    {
        var file = CreateSourceWorkbook();

        // Run the total down the Month axis (cacheField index 1: Region,Month,Sales).
        var envelope = EditOps(file, PivotOp(
            Value("Sales", "sum", ("showAs", "runningTotal"), ("baseField", "Month"))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var dataField = FirstDataField(document);
        Assert.Equal("runTotal", dataField.ShowDataAs!.InnerText);
        Assert.Equal(1, dataField.BaseField!.Value); // Month is the 2nd cacheField (0-based 1)
    }

    [Fact]
    public void DifferenceFrom_writes_difference_with_baseField_index_and_baseItem_index()
    {
        var file = CreateSourceWorkbook();

        // Difference vs the "Jan" item of the Month field: baseField=1, baseItem=index of Jan.
        var envelope = EditOps(file, PivotOp(
            Value("Sales", "sum", ("showAs", "differenceFrom"), ("baseField", "Month"), ("baseItem", "Jan"))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var dataField = FirstDataField(document);
        Assert.Equal("difference", dataField.ShowDataAs!.InnerText);
        Assert.Equal(1, dataField.BaseField!.Value);
        // Jan is the first Month item encountered in the source rows → sharedItems index 0.
        Assert.Equal(0u, dataField.BaseItem!.Value);
    }

    [Fact]
    public void PercentDifferenceFrom_previous_uses_the_previous_baseItem_sentinel()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(
            Value("Sales", "sum", ("showAs", "percentDifferenceFrom"), ("baseField", "Month"), ("baseItem", "(previous)"))));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var dataField = FirstDataField(document);
        Assert.Equal("percentDiff", dataField.ShowDataAs!.InnerText);
        Assert.Equal(1, dataField.BaseField!.Value);
        Assert.Equal(1048828u, dataField.BaseItem!.Value); // ECMA-376 "previous" sentinel
    }

    // ----- get reports the setting back ---------------------------------------

    [Fact]
    public void Get_pivot_reports_showAs_and_its_base()
    {
        var file = CreateSourceWorkbook();
        Assert.True(EditOps(file, PivotOp(
            Value("Sales", "sum", ("showAs", "differenceFrom"), ("baseField", "Month"), ("baseItem", "Jan")))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Pivot/pivot[@name=P1]"))));
        var value = data["values"]!.AsArray()[0]!;
        Assert.Equal("Sales", value["field"]!.GetValue<string>());
        Assert.Equal("differenceFrom", value["showAs"]!.GetValue<string>());
        Assert.Equal("Month", value["baseField"]!.GetValue<string>());
        Assert.Equal("Jan", value["baseItem"]!.GetValue<string>());
    }

    [Fact]
    public void Get_pivot_reports_percentOfTotal_with_no_base()
    {
        var file = CreateSourceWorkbook();
        Assert.True(EditOps(file, PivotOp(Value("Sales", "sum", ("showAs", "percentOfTotal")))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Pivot/pivot[@name=P1]"))));
        var value = data["values"]!.AsArray()[0]!;
        Assert.Equal("percentOfTotal", value["showAs"]!.GetValue<string>());
        Assert.Null(value["baseField"]);
    }

    [Fact]
    public void A_default_value_field_reports_no_showAs_and_writes_no_showDataAs()
    {
        var file = CreateSourceWorkbook();
        var envelope = EditOps(file, PivotOp(Value("Sales", "sum")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        // The details entry stays the lean {field, agg} shape (additive: no showAs key).
        var detail = OkData(envelope)["ops"]!.AsArray()[0]!["values"]!.AsArray()[0]!;
        Assert.Null(detail["showAs"]);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Null(FirstDataField(document).ShowDataAs);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Pivot/pivot[@name=P1]"))));
        Assert.Null(data["values"]!.AsArray()[0]!["showAs"]);
    }

    // ----- the bug is gone: unknown / unsupported are rejected -----------------

    [Fact]
    public void Unknown_showAs_is_invalid_args_with_candidates_not_silently_ignored()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(Value("Sales", "sum", ("showAs", "percentOfNonsense"))));

        Assert.False(envelope.IsOk, "an unknown showAs must be rejected, not silently ignored");
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("percentOfNonsense", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("percentOfTotal", envelope.Error.Candidates!);
        Assert.Contains("runningTotal", envelope.Error.Candidates!);
    }

    [Theory]
    [InlineData("percentOfParentTotal")]
    [InlineData("rankAscending")]
    [InlineData("rankDescending")]
    public void Unsupported_mode_is_unsupported_feature_not_silently_ignored(string showAs)
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(Value("Sales", "sum", ("showAs", showAs))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains(showAs, envelope.Error.Message, StringComparison.Ordinal);
        // The error names the supported alternatives.
        Assert.Contains("percentOfTotal", envelope.Error.Candidates!);
        Assert.DoesNotContain(showAs, envelope.Error.Candidates!);
    }

    // ----- base field / item validation ---------------------------------------

    [Fact]
    public void RunningTotal_without_baseField_is_invalid_args()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(Value("Sales", "sum", ("showAs", "runningTotal"))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("baseField", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DifferenceFrom_without_baseItem_is_invalid_args()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(
            Value("Sales", "sum", ("showAs", "differenceFrom"), ("baseField", "Month"))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("baseItem", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_baseField_is_invalid_args_with_the_real_headers()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(
            Value("Sales", "sum", ("showAs", "runningTotal"), ("baseField", "Quarter"))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("Quarter", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("Month", envelope.Error.Candidates!);
        Assert.Contains("Region", envelope.Error.Candidates!);
    }

    [Fact]
    public void PercentOfTotal_with_a_baseField_is_invalid_args()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(
            Value("Sales", "sum", ("showAs", "percentOfTotal"), ("baseField", "Month"))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("baseField", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_values_prop_is_invalid_args_with_candidates()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, PivotOp(Value("Sales", "sum", ("base", "Month"))));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("base", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("baseField", envelope.Error.Candidates!);
    }

    // ----- round-trip safety ---------------------------------------------------

    [Fact]
    public void File_with_showAs_reopens_through_the_closedXml_edit_path()
    {
        var file = CreateSourceWorkbook();
        Assert.True(EditOps(file, PivotOp(Value("Sales", "sum", ("showAs", "percentOfTotal")))).IsOk);

        // A second edit op proves the file still opens through the ClosedXML DOM path
        // (ClosedXML reads showDataAs back, so the round-trip is clean).
        var second = EditOps(file, SetOp("/Sheet1/E1", ("value", 1)));
        Assert.True(second.IsOk, second.ToJson());
        AssertValidatorClean(file);

        // And the showAs survived that second save.
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Equal("percentOfTotal", FirstDataField(document).ShowDataAs!.InnerText);
    }

    // ----- grand-total visibility (v1.14) -------------------------------------

    /// <summary>
    /// An add-pivot op over the show-values-as fixture carrying a <c>grandTotals</c>
    /// prop (string or object) plus the given value field. Reuses the same
    /// Region/Month/Sales source as the showAs tests.
    /// </summary>
    private EditOp GrandTotalsPivot(JsonNode? grandTotals, JsonObject valueEntry, string name = "GT") => new()
    {
        Op = "add",
        Path = "/Sheet1",
        Type = "pivot",
        Props = new JsonObject
        {
            ["name"] = name,
            ["sourceRange"] = "A1:C7",
            ["targetSheet"] = "Pivot",
            ["rows"] = new JsonArray("Region"),
            ["columns"] = new JsonArray("Month"),
            ["values"] = new JsonArray(valueEntry),
            ["grandTotals"] = grandTotals,
        },
    };

    /// <summary>Reopens the file through ClosedXML and reads the first pivot's grand-total flags.</summary>
    private static (bool Rows, bool Columns) ReopenGrandTotals(string file)
    {
        using var workbook = new XLWorkbook(file);
        var pivot = workbook.Worksheets
            .SelectMany(s => s.PivotTables)
            .First();
        return (pivot.ShowGrandTotalsRows, pivot.ShowGrandTotalsColumns);
    }

    [Theory]
    [InlineData("none", false, false)]
    [InlineData("rows", true, false)]
    [InlineData("columns", false, true)]
    [InlineData("both", true, true)]
    public void GrandTotals_string_survives_save_and_reopen(string mode, bool rows, bool columns)
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, GrandTotalsPivot(mode, Value("Sales", "sum")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        // Reported back in the add details, additively under grandTotals.
        var detail = OkData(envelope)["ops"]!.AsArray()[0]!["grandTotals"]!;
        Assert.Equal(rows, detail["rows"]!.GetValue<bool>());
        Assert.Equal(columns, detail["columns"]!.GetValue<bool>());

        // Authoritative oracle: reopen and read the persisted flags.
        var (reRows, reColumns) = ReopenGrandTotals(file);
        Assert.Equal(rows, reRows);
        Assert.Equal(columns, reColumns);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void GrandTotals_object_survives_save_and_reopen(bool rows, bool columns)
    {
        var file = CreateSourceWorkbook();
        var spec = new JsonObject { ["rows"] = rows, ["columns"] = columns };

        Assert.True(EditOps(file, GrandTotalsPivot(spec, Value("Sales", "sum"))).IsOk);
        AssertValidatorClean(file);

        var (reRows, reColumns) = ReopenGrandTotals(file);
        Assert.Equal(rows, reRows);
        Assert.Equal(columns, reColumns);
    }

    [Fact]
    public void GrandTotals_object_omits_a_flag_keeps_that_axis_shown()
    {
        var file = CreateSourceWorkbook();

        // Only columns is named false; rows is omitted and so stays shown (today's default).
        var spec = new JsonObject { ["columns"] = false };
        Assert.True(EditOps(file, GrandTotalsPivot(spec, Value("Sales", "sum"))).IsOk);

        var (rows, columns) = ReopenGrandTotals(file);
        Assert.True(rows);
        Assert.False(columns);
    }

    [Fact]
    public void Omitting_grandTotals_keeps_both_shown_and_writes_no_grand_total_attributes()
    {
        var file = CreateSourceWorkbook();

        // No grandTotals prop at all (the existing showAs PivotOp builder).
        Assert.True(EditOps(file, PivotOp(Value("Sales", "sum"), "Plain")).IsOk);
        AssertValidatorClean(file);

        var (rows, columns) = ReopenGrandTotals(file);
        Assert.True(rows);
        Assert.True(columns);

        // Byte-stability proxy: with the always-both default ClosedXML does not emit
        // the rowGrandTotals/colGrandTotals overrides, so the bytes match a pre-feature pivot.
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var definition = document.WorkbookPart!.WorksheetParts
            .SelectMany(p => p.PivotTableParts)
            .First().PivotTableDefinition!;
        Assert.Null(definition.RowGrandTotals);
        Assert.Null(definition.ColumnGrandTotals);
    }

    [Fact]
    public void GrandTotals_coexists_with_a_showAs_value_field_after_the_raw_pass()
    {
        var file = CreateSourceWorkbook();

        // grandTotals:none AND a showAs value field: proves the post-save raw show-values-as
        // pass does not clobber the grand-total flags (and vice-versa).
        var value = Value("Sales", "sum", ("showAs", "percentOfTotal"));
        var envelope = EditOps(file, GrandTotalsPivot("none", value));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        // The showAs landed on the dataField...
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            Assert.Equal("percentOfTotal", FirstDataField(document).ShowDataAs!.InnerText);
        }

        // ...and the grand totals survived the same save+raw pass.
        var (rows, columns) = ReopenGrandTotals(file);
        Assert.False(rows);
        Assert.False(columns);
    }

    [Fact]
    public void Get_pivot_reports_grandTotals_alongside_the_other_facets()
    {
        var file = CreateSourceWorkbook();
        var value = Value("Sales", "sum", ("showAs", "percentOfTotal"));
        Assert.True(EditOps(file, GrandTotalsPivot(
            new JsonObject { ["rows"] = true, ["columns"] = false }, value)).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Pivot/pivot[@name=GT]"))));

        // grandTotals is reported...
        var gt = data["grandTotals"]!;
        Assert.True(gt["rows"]!.GetValue<bool>());
        Assert.False(gt["columns"]!.GetValue<bool>());

        // ...and it coexists with the other facets get already surfaces.
        Assert.Equal("Region", data["rows"]!.AsArray()[0]!.GetValue<string>());
        Assert.Equal("Month", data["columns"]!.AsArray()[0]!.GetValue<string>());
        Assert.Equal("percentOfTotal", data["values"]!.AsArray()[0]!["showAs"]!.GetValue<string>());
        Assert.NotNull(data["filters"]);
        Assert.NotNull(data["calculatedFields"]);
    }

    [Fact]
    public void GrandTotals_unknown_string_is_invalid_args_with_candidates()
    {
        var file = CreateSourceWorkbook();

        var envelope = EditOps(file, GrandTotalsPivot("foo", Value("Sales", "sum")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("foo", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Equal(["both", "rows", "columns", "none"], envelope.Error.Candidates!);
    }

    [Fact]
    public void GrandTotals_non_boolean_flag_is_invalid_args()
    {
        var file = CreateSourceWorkbook();

        var spec = new JsonObject { ["rows"] = "yes" }; // string, not a boolean
        var envelope = EditOps(file, GrandTotalsPivot(spec, Value("Sales", "sum")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("rows", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void GrandTotals_unknown_sub_key_is_invalid_args_with_candidates()
    {
        var file = CreateSourceWorkbook();

        var spec = new JsonObject { ["row"] = false }; // typo for "rows"
        var envelope = EditOps(file, GrandTotalsPivot(spec, Value("Sales", "sum")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("row", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("rows", envelope.Error.Candidates!);
    }
}
