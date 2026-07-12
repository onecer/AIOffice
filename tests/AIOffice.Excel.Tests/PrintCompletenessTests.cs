using System.Linq;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.7) Print completeness: print titles, manual page breaks, fit-to-page, the
/// center/gridlines/headings toggles, and print header/footer field codes. Every
/// mutating test reopens the file with the OpenXml SDK (the validator + raw part
/// inspection oracles) and round-trips through <c>get</c>.
/// </summary>
public sealed class PrintCompletenessTests : ExcelTestBase
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

    // ----- print titles -------------------------------------------------------

    [Fact]
    public void Print_titles_write_repeat_bands_and_get_reflects()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("printTitleRows", "1:1"), ("printTitleCols", "A:A")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            // Repeat titles live as the _xlnm.Print_Titles defined name (ClosedXML
            // writes the bands as Sheet1!A:A,Sheet1!1:1).
            var printTitles = document.WorkbookPart!.Workbook!.Descendants<S.DefinedName>()
                .Single(n => n.Name?.Value == "_xlnm.Print_Titles");
            Assert.Contains("1:1", printTitles.Text, StringComparison.Ordinal);
            Assert.Contains("A:A", printTitles.Text, StringComparison.Ordinal);
        }

        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        Assert.Equal("1:1", ps["printTitleRows"]!.GetValue<string>());
        Assert.Equal("A:A", ps["printTitleCols"]!.GetValue<string>());
    }

    [Fact]
    public void Print_title_rows_multi_row_band_round_trips()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("printTitleRows", "1:3"))).IsOk);
        AssertValidatorClean(file);
        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        Assert.Equal("1:3", ps["printTitleRows"]!.GetValue<string>());
    }

    [Fact]
    public void Empty_print_title_rows_clears_it()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("printTitleRows", "1:1"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1", ("printTitleRows", ""))).IsOk);

        AssertValidatorClean(file);
        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        Assert.Null(ps["printTitleRows"]);
    }

    [Fact]
    public void Bad_print_title_rows_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1", ("printTitleRows", "A:A")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    // ----- page breaks --------------------------------------------------------

    [Fact]
    public void Manual_page_breaks_write_row_and_col_breaks_and_get_reflects()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("pageBreaks", new JsonObject
        {
            ["rows"] = new JsonArray(20, 40),
            ["cols"] = new JsonArray("F"),
        })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var rowBreaks = RawSheet(document).Elements<S.RowBreaks>().Single();
            Assert.Equal(2, rowBreaks.Elements<S.Break>().Count());
            Assert.Contains(rowBreaks.Elements<S.Break>(), b => b.Id?.Value == 20u);
            Assert.Contains(rowBreaks.Elements<S.Break>(), b => b.Id?.Value == 40u);

            var colBreaks = RawSheet(document).Elements<S.ColumnBreaks>().Single();
            Assert.Single(colBreaks.Elements<S.Break>());
            Assert.Equal(6u, colBreaks.Elements<S.Break>().Single().Id?.Value); // F = column 6
        }

        var pb = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!["pageBreaks"]!;
        Assert.Equal([20, 40], pb["rows"]!.AsArray().Select(n => n!.GetValue<int>()).ToArray());
        Assert.Equal("F", pb["cols"]![0]!.GetValue<string>());
    }

    [Fact]
    public void Numeric_col_break_is_accepted()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("pageBreaks", new JsonObject
        {
            ["cols"] = new JsonArray(6),
        }))).IsOk);
        AssertValidatorClean(file);
        var pb = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!["pageBreaks"]!;
        Assert.Equal("F", pb["cols"]![0]!.GetValue<string>());
    }

    [Fact]
    public void Empty_page_break_array_clears_that_axis()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("pageBreaks", new JsonObject
        {
            ["rows"] = new JsonArray(20),
        }))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1", ("pageBreaks", new JsonObject
        {
            ["rows"] = new JsonArray(),
        }))).IsOk);

        AssertValidatorClean(file);
        Assert.Null(OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!["pageBreaks"]);
    }

    [Fact]
    public void Page_break_aliases_rowBreaks_colBreaks_work()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("pageBreaks", new JsonObject
        {
            ["rowBreaks"] = new JsonArray(15),
            ["colBreaks"] = new JsonArray("C"),
        }))).IsOk);
        AssertValidatorClean(file);
        var pb = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!["pageBreaks"]!;
        Assert.Equal(15, pb["rows"]![0]!.GetValue<int>());
        Assert.Equal("C", pb["cols"]![0]!.GetValue<string>());
    }

    // ----- fit to page --------------------------------------------------------

    [Fact]
    public void Fit_to_page_writes_pages_wide_tall_and_get_reflects()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("fitToPage", new JsonObject
        {
            ["fitToWidth"] = 2,
            ["fitToHeight"] = 3,
        })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var setup = RawSheet(document).Elements<S.PageSetup>().Single();
            Assert.Equal(2u, setup.FitToWidth?.Value);
            Assert.Equal(3u, setup.FitToHeight?.Value);
            var fitToPage = RawSheet(document).Elements<S.SheetProperties>()
                .SingleOrDefault()?.PageSetupProperties?.FitToPage;
            Assert.True(fitToPage?.Value);
        }

        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        Assert.Equal(2, ps["fitToWidth"]!.GetValue<int>());
        Assert.Equal(3, ps["fitToHeight"]!.GetValue<int>());
    }

    [Fact]
    public void Fit_to_page_scale_writes_scale_and_get_reflects()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("fitToPage", new JsonObject { ["scale"] = 80 })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var setup = RawSheet(document).Elements<S.PageSetup>().Single();
            Assert.Equal(80u, setup.Scale?.Value);
        }

        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        Assert.Equal(80, ps["scale"]!.GetValue<int>());
    }

    [Fact]
    public void Fit_to_page_scale_with_pages_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1", ("fitToPage", new JsonObject
        {
            ["scale"] = 80,
            ["fitToWidth"] = 1,
        })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("not both", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Fit_to_page_scale_out_of_range_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1", ("fitToPage", new JsonObject { ["scale"] = 5 })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    [Fact]
    public void Flat_fit_to_height_prop_works()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("fitToWidth", 1), ("fitToHeight", 3))).IsOk);
        AssertValidatorClean(file);
        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        Assert.Equal(1, ps["fitToWidth"]!.GetValue<int>());
        Assert.Equal(3, ps["fitToHeight"]!.GetValue<int>());
    }

    // ----- center / gridlines / headings --------------------------------------

    [Fact]
    public void Center_gridlines_headings_toggles_write_print_options_and_get_reflects()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp(
            "/Sheet1",
            ("centerHorizontally", true), ("centerVertically", true),
            ("printGridlines", true), ("printHeadings", true)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var printOptions = RawSheet(document).Elements<S.PrintOptions>().Single();
            Assert.True(printOptions.HorizontalCentered?.Value);
            Assert.True(printOptions.VerticalCentered?.Value);
            Assert.True(printOptions.GridLines?.Value);
            Assert.True(printOptions.Headings?.Value);
        }

        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        Assert.True(ps["centerHorizontally"]!.GetValue<bool>());
        Assert.True(ps["centerVertically"]!.GetValue<bool>());
        Assert.True(ps["printGridlines"]!.GetValue<bool>());
        Assert.True(ps["printHeadings"]!.GetValue<bool>());
    }

    [Fact]
    public void Center_toggles_off_clears_from_get()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("centerHorizontally", true))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1", ("centerHorizontally", false))).IsOk);

        AssertValidatorClean(file);
        Assert.Null(OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!["centerHorizontally"]);
    }

    // ----- print header / footer ----------------------------------------------

    [Fact]
    public void Print_header_and_footer_field_codes_write_oddHeader_oddFooter_and_get_reflects()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp(
            "/Sheet1",
            ("printHeader", new JsonObject { ["left"] = "&F", ["center"] = "&A", ["right"] = "&D" }),
            ("printFooter", new JsonObject { ["center"] = "Page &P of &N" })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var headerFooter = RawSheet(document).Elements<S.HeaderFooter>().Single();
            var oddHeader = headerFooter.OddHeader!.Text;
            Assert.Contains("&L&F", oddHeader, StringComparison.Ordinal);
            Assert.Contains("&C&A", oddHeader, StringComparison.Ordinal);
            Assert.Contains("&R&D", oddHeader, StringComparison.Ordinal);
            Assert.Contains("Page &P of &N", headerFooter.OddFooter!.Text, StringComparison.Ordinal);
        }

        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        Assert.Equal("&F", ps["printHeader"]!["left"]!.GetValue<string>());
        Assert.Equal("&A", ps["printHeader"]!["center"]!.GetValue<string>());
        Assert.Equal("&D", ps["printHeader"]!["right"]!.GetValue<string>());
        Assert.Equal("Page &P of &N", ps["printFooter"]!["center"]!.GetValue<string>());
    }

    [Fact]
    public void Print_header_section_can_be_changed_independently()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("printHeader",
            new JsonObject { ["left"] = "&F", ["right"] = "&D" }))).IsOk);

        // Change only the center; left/right survive.
        Assert.True(EditOps(file, SetOp("/Sheet1", ("printHeader",
            new JsonObject { ["center"] = "Report" }))).IsOk);

        AssertValidatorClean(file);
        var header = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!["printHeader"]!;
        Assert.Equal("&F", header["left"]!.GetValue<string>());
        Assert.Equal("Report", header["center"]!.GetValue<string>());
        Assert.Equal("&D", header["right"]!.GetValue<string>());
    }

    [Fact]
    public void Null_header_section_clears_it()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1", ("printHeader",
            new JsonObject { ["left"] = "&F", ["center"] = "&A" }))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1", ("printHeader",
            new JsonObject { ["left"] = null }))).IsOk);

        AssertValidatorClean(file);
        var header = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!["printHeader"]!;
        Assert.Null(header["left"]);
        Assert.Equal("&A", header["center"]!.GetValue<string>());
    }

    [Fact]
    public void Unknown_header_section_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1", ("printHeader",
            new JsonObject { ["middle"] = "x" })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    // ----- combined + survival ------------------------------------------------

    [Fact]
    public void All_print_props_survive_an_unrelated_edit_and_reopen()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, SetOp(
            "/Sheet1",
            ("printTitleRows", "1:1"),
            ("pageBreaks", new JsonObject { ["rows"] = new JsonArray(10), ["cols"] = new JsonArray("D") }),
            ("fitToPage", new JsonObject { ["fitToWidth"] = 1, ["fitToHeight"] = 0 }),
            ("printGridlines", true),
            ("printHeader", new JsonObject { ["center"] = "&A" }))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/F9", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        Assert.Equal("1:1", ps["printTitleRows"]!.GetValue<string>());
        Assert.Equal(10, ps["pageBreaks"]!["rows"]![0]!.GetValue<int>());
        Assert.Equal("D", ps["pageBreaks"]!["cols"]![0]!.GetValue<string>());
        Assert.Equal(1, ps["fitToWidth"]!.GetValue<int>());
        Assert.True(ps["printGridlines"]!.GetValue<bool>());
        Assert.Equal("&A", ps["printHeader"]!["center"]!.GetValue<string>());
    }

    [Fact]
    public void Print_completeness_props_on_a_range_path_are_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1/A1:B2", ("printTitleRows", "1:1")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    // ----- page margins (1.25) ------------------------------------------------

    [Fact]
    public void Margins_write_page_margins_in_inches_and_get_reflects()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, SetOp("/Sheet1", ("margins", new JsonObject
        {
            ["top"] = 1.0,
            ["bottom"] = 1.0,
            ["left"] = 0.75,
            ["right"] = 0.75,
            ["header"] = 0.5,
            ["footer"] = 0.5,
        })));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        // <pageMargins> stores the values verbatim in inches.
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var pm = RawSheet(document).GetFirstChild<S.PageMargins>()!;
            Assert.Equal(1.0, pm.Top!.Value);
            Assert.Equal(1.0, pm.Bottom!.Value);
            Assert.Equal(0.75, pm.Left!.Value);
            Assert.Equal(0.75, pm.Right!.Value);
            Assert.Equal(0.5, pm.Header!.Value);
            Assert.Equal(0.5, pm.Footer!.Value);
        }

        // save+reopen: get echoes the margins object.
        var m = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!["margins"]!;
        Assert.Equal(1.0, m["top"]!.GetValue<double>());
        Assert.Equal(1.0, m["bottom"]!.GetValue<double>());
        Assert.Equal(0.75, m["left"]!.GetValue<double>());
        Assert.Equal(0.75, m["right"]!.GetValue<double>());
        Assert.Equal(0.5, m["header"]!.GetValue<double>());
        Assert.Equal(0.5, m["footer"]!.GetValue<double>());
    }

    [Fact]
    public void Partial_margins_leave_the_other_edges_untouched()
    {
        var file = CreateDataWorkbook();

        // ClosedXML defaults: top 0.75, bottom 0.5, left 0.75, right 0.75,
        // header 0.5, footer 0.75. A partial object touches only top+bottom.
        Assert.True(EditOps(file, SetOp("/Sheet1", ("margins", new JsonObject
        {
            ["top"] = 1.25,
            ["bottom"] = 1.5,
        }))).IsOk);
        AssertValidatorClean(file);

        var m = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!["margins"]!;
        Assert.Equal(1.25, m["top"]!.GetValue<double>());
        Assert.Equal(1.5, m["bottom"]!.GetValue<double>());
        Assert.Equal(0.75, m["left"]!.GetValue<double>());
        Assert.Equal(0.75, m["right"]!.GetValue<double>());
        Assert.Equal(0.5, m["header"]!.GetValue<double>());
        Assert.Equal(0.75, m["footer"]!.GetValue<double>());
    }

    [Fact]
    public void Default_margins_emit_no_margins_key()
    {
        var file = CreateDataWorkbook();
        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        // Deviation-gated: the key is genuinely ABSENT (WhenWritingNull drops the null in the
        // anon PageSetupInfo) — assert on the key set so a future refactor to a null-serializing
        // Dictionary can't silently add a 'margins':null field to every legacy sheet.
        Assert.False(ps.AsObject().ContainsKey("margins"));
    }

    [Fact]
    public void Unrelated_edit_leaves_default_page_margins_byte_stable()
    {
        var file = CreateDataWorkbook();

        string RawMargins()
        {
            using var document = SpreadsheetDocument.Open(file, isEditable: false);
            return RawSheet(document).GetFirstChild<S.PageMargins>()!.OuterXml;
        }

        var golden = RawMargins();

        // An edit that never touches margins must not perturb <pageMargins>.
        Assert.True(EditOps(file, SetOp("/Sheet1/F9", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        Assert.Equal(golden, RawMargins());
        var ps = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["pageSetup"]!;
        Assert.False(ps.AsObject().ContainsKey("margins"));
    }

    [Fact]
    public void Non_object_margins_is_invalid_args_with_the_six_edge_candidates()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1", ("margins", 1.0)));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Equal(6, envelope.Error.Candidates!.Count);
        Assert.Contains("top", envelope.Error.Candidates!);
        Assert.Contains("footer", envelope.Error.Candidates!);
    }

    [Fact]
    public void Negative_margin_edge_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1", ("margins", new JsonObject { ["top"] = -0.5 })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Equal(6, envelope.Error.Candidates!.Count);
        Assert.Contains("top", envelope.Error.Candidates!);
    }

    [Fact]
    public void NaN_margin_edge_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1", ("margins", new JsonObject { ["left"] = double.NaN })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("left", envelope.Error.Candidates!);
    }

    [Fact]
    public void Unknown_margin_edge_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1", ("margins", new JsonObject { ["gutter"] = 0.5 })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Equal(6, envelope.Error.Candidates!.Count);
    }

    [Fact]
    public void Margins_on_a_range_path_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, SetOp("/Sheet1/A1:B2", ("margins", new JsonObject { ["top"] = 1.0 })));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }
}
