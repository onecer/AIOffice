using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M3 sheet-prop slice: freeze panes, autofilter (one per sheet) and page
/// setup (orientation | paperSize | fitToWidth | printArea). Raw assertions
/// reopen the file with the OpenXml SDK; get must reflect everything back.
/// </summary>
public sealed class SheetPropsTests : ExcelTestBase
{
    private string CreateDataWorkbook()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:D3", ("values", new JsonArray(
                new JsonArray("Name", "Qty", "Price", "Note"),
                new JsonArray("ant", 5, 1.5, "small"),
                new JsonArray("bee", 7, 2.5, "buzzy"))))).IsOk);
        return file;
    }

    private static S.Worksheet RawSheet(SpreadsheetDocument document, string name = "Sheet1")
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
        return ((WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!)).Worksheet!;
    }

    // ----- freeze panes -------------------------------------------------------

    [Fact]
    public void Freeze_rows_and_cols_writes_a_frozen_pane_and_get_reflects()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("freezeRows", 1), ("freezeCols", 2)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var pane = RawSheet(document).Descendants<S.Pane>().Single();
            Assert.Equal(1d, pane.VerticalSplit?.Value);
            Assert.Equal(2d, pane.HorizontalSplit?.Value);
            Assert.Equal(S.PaneStateValues.FrozenSplit, pane.State?.Value); // what ClosedXML emits for frozen panes
        }

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Equal(1, data["freezeRows"]!.GetValue<int>());
        Assert.Equal(2, data["freezeCols"]!.GetValue<int>());
    }

    [Fact]
    public void Freeze_zero_clears_and_get_omits_the_props()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("freezeRows", 3), ("freezeCols", 1))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1", ("freezeRows", 0), ("freezeCols", 0))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Null(data["freezeRows"]);
        Assert.Null(data["freezeCols"]);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Empty(RawSheet(document).Descendants<S.Pane>());
    }

    [Fact]
    public void Freeze_on_a_cell_path_is_invalid_args()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1", ("freezeRows", 1)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("/Sheet1", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    // ----- autofilter ----------------------------------------------------------

    [Fact]
    public void AutoFilter_set_writes_the_ref_and_shows_in_structure_and_get()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1:D3", ("autoFilter", true)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var filter = RawSheet(document).Elements<S.AutoFilter>().Single();
            Assert.Equal("A1:D3", filter.Reference?.Value);
        }

        var sheetInfo = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Equal("A1:D3", sheetInfo["autoFilter"]!.GetValue<string>());

        var structure = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        Assert.Equal("A1:D3", structure["sheets"]![0]!["autoFilter"]!.GetValue<string>());
    }

    [Fact]
    public void Second_autoFilter_on_another_range_is_invalid_args_with_the_clear_recipe()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:D3", ("autoFilter", true))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/A1:B3", ("autoFilter", true)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("one per sheet", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("autoFilter:false", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void AutoFilter_false_clears_and_views_stop_listing_it()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:D3", ("autoFilter", true))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/A1:D3", ("autoFilter", false)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);
        var sheetInfo = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Null(sheetInfo["autoFilter"]);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Empty(RawSheet(document).Elements<S.AutoFilter>());
    }

    [Fact]
    public void AutoFilter_on_a_sheet_path_is_invalid_args()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("autoFilter", true)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("range", envelope.Error.Message, StringComparison.Ordinal);
    }

    // ----- autofilter criteria (1.12) ------------------------------------------

    /// <summary>A Region/Amount table: header + 4 rows over A1:B5.</summary>
    private string CreateFilterTable()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:B5", ("values", new JsonArray(
                new JsonArray("Region", "Amount"),
                new JsonArray("East", 10),
                new JsonArray("West", 200),
                new JsonArray("North", 30),
                new JsonArray("South", 400))))).IsOk);
        return file;
    }

    private static bool RowHidden(SpreadsheetDocument document, int rowNumber, string sheet = "Sheet1")
    {
        var row = RawSheet(document, sheet).Descendants<S.Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowNumber);
        return row?.Hidden?.Value == true;
    }

    [Fact]
    public void AutoFilter_values_filter_hides_nonmatching_rows_and_round_trips()
    {
        var file = CreateFilterTable();

        var values = new JsonObject { ["column"] = "Region", ["values"] = new JsonArray("East", "West") };
        var envelope = EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", values)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            // East (row 2) and West (row 3) stay; North (4) and South (5) hide.
            Assert.False(RowHidden(document, 2));
            Assert.False(RowHidden(document, 3));
            Assert.True(RowHidden(document, 4));
            Assert.True(RowHidden(document, 5));

            var filter = RawSheet(document).Elements<S.AutoFilter>().Single();
            var column = filter.Elements<S.FilterColumn>().Single();
            Assert.Equal(0u, column.ColumnId?.Value); // first column of the range
            var vals = column.Elements<S.Filters>().Single().Elements<S.Filter>()
                .Select(f => f.Val?.Value).ToList();
            Assert.Equal(["East", "West"], vals);
        }

        var sheetInfo = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Equal("A1:B5", sheetInfo["autoFilter"]!.GetValue<string>());
        var columns = sheetInfo["autoFilterColumns"]!.AsArray();
        Assert.Single(columns);
        Assert.Equal("A", columns[0]!["column"]!.GetValue<string>());
        Assert.Equal("values", columns[0]!["kind"]!.GetValue<string>());
        Assert.Equal(
            ["East", "West"],
            columns[0]!["values"]!.AsArray().Select(v => v!.GetValue<string>()).ToList());
    }

    [Fact]
    public void AutoFilter_comparison_filter_hides_nonmatching_rows_and_round_trips()
    {
        var file = CreateFilterTable();

        // Amount > 100 keeps West (200) and South (400); hides East (10), North (30).
        var criteria = new JsonObject { ["column"] = "Amount", ["criteria"] = ">100" };
        var envelope = EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", criteria)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            Assert.True(RowHidden(document, 2)); // East 10
            Assert.False(RowHidden(document, 3)); // West 200
            Assert.True(RowHidden(document, 4)); // North 30
            Assert.False(RowHidden(document, 5)); // South 400

            var column = RawSheet(document).Elements<S.AutoFilter>().Single()
                .Elements<S.FilterColumn>().Single();
            Assert.Equal(1u, column.ColumnId?.Value); // second column of the range
            var custom = column.Elements<S.CustomFilters>().Single().Elements<S.CustomFilter>().Single();
            Assert.Equal(S.FilterOperatorValues.GreaterThan, custom.Operator?.Value);
            Assert.Equal("100", custom.Val?.Value);
        }

        var columns = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["autoFilterColumns"]!.AsArray();
        Assert.Single(columns);
        Assert.Equal("B", columns[0]!["column"]!.GetValue<string>());
        Assert.Equal("custom", columns[0]!["kind"]!.GetValue<string>());
        Assert.Equal("greaterThan", columns[0]!["criteria"]![0]!["operator"]!.GetValue<string>());
        Assert.Equal("100", columns[0]!["criteria"]![0]!["value"]!.GetValue<string>());
    }

    [Fact]
    public void AutoFilter_wildcard_text_criteria_uses_a_contains_filter()
    {
        var file = CreateFilterTable();

        // "*est*" is a contains filter on the inner text "est": only West contains it
        // (East = E-a-s-t has no "est"); North and South hide too.
        var criteria = new JsonObject { ["column"] = "A", ["criteria"] = "*est*" };
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", criteria))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.True(RowHidden(document, 2)); // East
        Assert.False(RowHidden(document, 3)); // West
        Assert.True(RowHidden(document, 4)); // North
        Assert.True(RowHidden(document, 5)); // South

        var custom = RawSheet(document).Elements<S.AutoFilter>().Single()
            .Elements<S.FilterColumn>().Single()
            .Elements<S.CustomFilters>().Single().Elements<S.CustomFilter>().Single();
        Assert.Equal("*est*", custom.Val?.Value); // a contains filter is stored as *inner*
    }

    [Fact]
    public void AutoFilter_multiple_columns_apply_an_and_filter()
    {
        var file = CreateFilterTable();

        var filters = new JsonArray(
            new JsonObject { ["column"] = "Region", ["values"] = new JsonArray("East", "West", "South") },
            new JsonObject { ["column"] = "Amount", ["criteria"] = ">=200" });
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", filters))).IsOk);

        AssertValidatorClean(file);
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            // Region in {East,West,South} AND Amount>=200 → West(200), South(400).
            Assert.True(RowHidden(document, 2)); // East 10 (fails amount)
            Assert.False(RowHidden(document, 3)); // West 200
            Assert.True(RowHidden(document, 4)); // North (fails region)
            Assert.False(RowHidden(document, 5)); // South 400
        }

        var columns = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["autoFilterColumns"]!.AsArray();
        Assert.Equal(2, columns.Count);
    }

    [Fact]
    public void AutoFilter_criteria_by_one_based_index_resolves_within_the_range()
    {
        var file = CreateFilterTable();

        // index 2 within A1:B5 is the Amount column.
        var criteria = new JsonObject { ["column"] = 2, ["criteria"] = "<100" };
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", criteria))).IsOk);

        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.False(RowHidden(document, 2)); // East 10
        Assert.True(RowHidden(document, 3)); // West 200
        Assert.False(RowHidden(document, 4)); // North 30
        Assert.True(RowHidden(document, 5)); // South 400
    }

    [Fact]
    public void AutoFilter_unknown_column_is_invalid_args_with_header_candidates()
    {
        var file = CreateFilterTable();

        var criteria = new JsonObject { ["column"] = "Nope", ["criteria"] = ">1" };
        var envelope = EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", criteria)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("Nope", envelope.Error.Message, StringComparison.Ordinal);
        Assert.NotNull(envelope.Error.Candidates);
        Assert.Contains("Region", envelope.Error.Candidates!);
        Assert.Contains("Amount", envelope.Error.Candidates!);
    }

    [Fact]
    public void AutoFilter_column_index_outside_the_range_is_invalid_args()
    {
        var file = CreateFilterTable();

        var criteria = new JsonObject { ["column"] = 5, ["criteria"] = ">1" }; // only 2 columns
        var envelope = EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", criteria)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("outside the filter range", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AutoFilter_entry_needs_exactly_one_of_values_or_criteria()
    {
        var file = CreateFilterTable();

        var both = new JsonObject
        {
            ["column"] = "Region",
            ["values"] = new JsonArray("East"),
            ["criteria"] = ">1",
        };
        var envelope = EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", both)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("exactly one", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AutoFilter_false_clears_a_criteria_filter_and_unhides_rows()
    {
        var file = CreateFilterTable();
        var criteria = new JsonObject { ["column"] = "Amount", ["criteria"] = ">100" };
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", criteria))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", false))).IsOk);

        AssertValidatorClean(file);
        var sheetInfo = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Null(sheetInfo["autoFilter"]);
        Assert.Null(sheetInfo["autoFilterColumns"]);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Empty(RawSheet(document).Elements<S.AutoFilter>());
    }

    [Fact]
    public void AutoFilter_bool_form_still_reports_no_criteria_columns()
    {
        var file = CreateFilterTable();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B5", ("autoFilter", true))).IsOk);

        var sheetInfo = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Equal("A1:B5", sheetInfo["autoFilter"]!.GetValue<string>());
        // A plain enabled filter carries no per-column criteria.
        Assert.Null(sheetInfo["autoFilterColumns"]);
    }

    // ----- page setup -----------------------------------------------------------

    [Fact]
    public void Page_setup_props_write_raw_xml_and_get_reflects_them()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp(
            "/Sheet1",
            ("orientation", "landscape"), ("paperSize", "A4"), ("fitToWidth", 2), ("printArea", "A1:F40")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var setup = RawSheet(document).Elements<S.PageSetup>().Single();
            Assert.Equal(S.OrientationValues.Landscape, setup.Orientation?.Value);
            Assert.Equal(9u, setup.PaperSize?.Value); // 9 = ISO A4
            Assert.Equal(2u, setup.FitToWidth?.Value);

            // Fitting needs the sheetPr fitToPage flag too.
            var fitToPage = RawSheet(document).Elements<S.SheetProperties>()
                .SingleOrDefault()?.PageSetupProperties?.FitToPage;
            Assert.True(fitToPage?.Value);

            // The print area lives as the _xlnm.Print_Area defined name.
            var printArea = document.WorkbookPart!.Workbook!.Descendants<S.DefinedName>()
                .Single(n => n.Name?.Value == "_xlnm.Print_Area");
            Assert.Contains("$A$1:$F$40", printArea.Text, StringComparison.Ordinal);
        }

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        var pageSetup = data["pageSetup"]!;
        Assert.Equal("landscape", pageSetup["orientation"]!.GetValue<string>());
        Assert.Equal("A4", pageSetup["paperSize"]!.GetValue<string>());
        Assert.Equal(2, pageSetup["fitToWidth"]!.GetValue<int>());
        Assert.Equal("A1:F40", pageSetup["printArea"]!.GetValue<string>());
    }

    [Fact]
    public void Empty_printArea_clears_it()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("printArea", "A1:B2"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1", ("printArea", ""))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Null(data["pageSetup"]!["printArea"]);
    }

    [Fact]
    public void Unknown_paperSize_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("paperSize", "A7")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("A4", envelope.Error.Candidates!);
        Assert.Contains("Letter", envelope.Error.Candidates!);
    }

    [Fact]
    public void Unknown_orientation_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("orientation", "sideways")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Equal(["portrait", "landscape"], envelope.Error.Candidates!);
    }

    [Fact]
    public void Bad_printArea_is_invalid_args()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("printArea", "Sheet2!A1:B2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("A1:F40", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Page_setup_on_a_range_path_is_invalid_args()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1/A1:B2", ("orientation", "landscape")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    [Fact]
    public void All_sheet_props_survive_an_unrelated_edit_and_reopen()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1", ("freezeRows", 1)),
            SetOp("/Sheet1/A1:D3", ("autoFilter", true)),
            SetOp("/Sheet1", ("orientation", "landscape"), ("paperSize", "Letter"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/F9", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Equal(1, data["freezeRows"]!.GetValue<int>());
        Assert.Equal("A1:D3", data["autoFilter"]!.GetValue<string>());
        Assert.Equal("landscape", data["pageSetup"]!["orientation"]!.GetValue<string>());
        Assert.Equal("Letter", data["pageSetup"]!["paperSize"]!.GetValue<string>());
    }
}
