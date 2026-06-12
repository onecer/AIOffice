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
