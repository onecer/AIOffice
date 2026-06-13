using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M9 conversion coverage for the xlsx handler's <see cref="INeutralConvertible"/>
/// surface. Exercises the round-trip law of <c>convert</c> at the content level:
/// a workbook (multiple sheets, formulas, number formats) projects to the
/// format-neutral model, that model imports into a fresh workbook, and sheet
/// names + values survive. Also pins the honest lossiness contract: charts/images
/// become <c>Dropped</c> notes on export; paragraph/heading blocks land in a
/// leading Notes sheet on import; the file stays validator-clean and deterministic.
/// </summary>
public sealed class ConvertTests : ExcelTestBase
{
    // ----- export: xlsx -> neutral --------------------------------------------

    [Fact]
    public void Export_projects_each_sheet_to_a_heading_then_table()
    {
        var file = CreateWorkbook(title: "Sales");
        Assert.True(EditOps(file, AddOp("/Costs", "sheet")).IsOk);
        Assert.True(EditOps(file, SetOp("/Sales/A1:B3", ("values", Grid(
            ["Region", "Total"],
            ["North", "100"],
            ["South", "200"])))).IsOk);
        Assert.True(EditOps(file, SetOp("/Costs/A1:B2", ("values", Grid(
            ["Item", "Cost"],
            ["Rent", "50"])))).IsOk);

        var doc = Handler.ExportNeutral(Ctx(file));

        // Two (Heading, Table) pairs, in sheet position order.
        var headings = doc.Blocks.Where(b => b.Kind == NeutralBlockKind.Heading).ToList();
        var tables = doc.Blocks.Where(b => b.Kind == NeutralBlockKind.Table).ToList();
        Assert.Equal(["Sales", "Costs"], headings.Select(h => h.Runs![0].Text));
        Assert.Equal(2, tables.Count);
        Assert.Equal(2, headings[0].Level); // sheet names are level-2 headings

        // First table: header row detected (text labels over numeric data).
        Assert.True(tables[0].HeaderRow);
        Assert.Equal(["Region", "Total"], tables[0].Rows![0]);
        Assert.Equal(["North", "100"], tables[0].Rows![1]);

        // Title comes from the core Title property.
        Assert.Equal("Sales", doc.Title);
    }

    [Fact]
    public void Export_exports_formula_cells_as_their_display_value()
    {
        var file = CreateWorkbook(title: "Nums");
        Assert.True(EditOps(file, SetOp("/Nums/A1", ("value", 4))).IsOk);
        Assert.True(EditOps(file, SetOp("/Nums/A2", ("value", 6))).IsOk);
        Assert.True(EditOps(file, SetOp("/Nums/A3", ("value", "=SUM(A1:A2)"))).IsOk);

        var doc = Handler.ExportNeutral(Ctx(file));
        var table = doc.Blocks.First(b => b.Kind == NeutralBlockKind.Table);

        // The SUM cell exports as its cached display value "10", not "=SUM(A1:A2)".
        Assert.Equal("10", table.Rows![2][0]);
        Assert.DoesNotContain(table.Rows!.SelectMany(r => r), v => v.StartsWith('='));
    }

    [Fact]
    public void Export_formats_numbers_via_the_cell_number_format()
    {
        var file = CreateWorkbook(title: "Fmt");
        Assert.True(EditOps(file, SetOp("/Fmt/A1", ("value", 0.25))).IsOk);
        Assert.True(EditOps(file, SetOp("/Fmt/A1", ("numberFormat", "0.00%"))).IsOk);

        var doc = Handler.ExportNeutral(Ctx(file));
        var table = doc.Blocks.First(b => b.Kind == NeutralBlockKind.Table);

        // Display value honours the number format: 0.25 -> "25.00%".
        Assert.Equal("25.00%", table.Rows![0][0]);
    }

    [Fact]
    public void Export_drops_a_chart_with_an_honest_note()
    {
        var file = CreateWorkbook(title: "Data");
        Assert.True(EditOps(file, SetOp("/Data/A1:B3", ("values", Grid(
            ["X", "Y"],
            ["1", "10"],
            ["2", "20"])))).IsOk);
        Assert.True(EditOps(file, AddOp(
            "/Data",
            "chart",
            ("kind", "bar"),
            ("dataRange", "A1:B3"),
            ("anchor", "D2"),
            ("title", "Trend"))).IsOk);

        var doc = Handler.ExportNeutral(Ctx(file));

        // The chart is not representable; the export emits an honest "[dropped]"
        // note paragraph naming it (so the loss survives into convert_lossy and
        // the converted document, instead of vanishing silently).
        Assert.Contains(
            doc.Blocks,
            b => b.Kind == NeutralBlockKind.Paragraph &&
                 b.Runs is { Count: > 0 } &&
                 b.Runs[0].Text.Contains("[dropped] chart", StringComparison.Ordinal) &&
                 b.Runs[0].Text.Contains("not representable", StringComparison.Ordinal));

        // Importing the neutral doc into a fresh xlsx: the data round-trips into a
        // Data sheet and the dropped-chart note lands in the Notes sheet.
        var clone = NewFile("clone.xlsx");
        Handler.Create(Ctx(clone));
        Handler.ImportNeutral(Ctx(clone), doc);

        using var workbook = new XLWorkbook(clone);
        Assert.True(workbook.TryGetWorksheet("Data", out _));
        Assert.True(workbook.TryGetWorksheet("Notes", out var notes));
        Assert.Contains(
            notes.Column(1).CellsUsed().Select(c => c.GetString()),
            s => s.Contains("[dropped] chart", StringComparison.Ordinal));
    }

    // ----- import: neutral -> xlsx --------------------------------------------

    [Fact]
    public void Import_writes_each_table_to_a_named_sheet_with_a_frozen_header()
    {
        var doc = new NeutralDoc("Report", new List<NeutralBlock>
        {
            new(NeutralBlockKind.Heading, Level: 2, Runs: [new NeutralRun("People")]),
            new(NeutralBlockKind.Table, HeaderRow: true, Rows: Rows(
                ["Name", "Age"],
                ["Ada", "36"])),
        });

        var file = NewFile("imported.xlsx");
        Handler.Create(Ctx(file));
        var result = Handler.ImportNeutral(Ctx(file), doc);

        Assert.Equal(1, result.BlocksWritten);
        AssertValidatorClean(file);

        using var workbook = new XLWorkbook(file);
        Assert.True(workbook.TryGetWorksheet("People", out var sheet));
        Assert.Equal("Name", sheet.Cell("A1").GetString());
        Assert.True(sheet.Cell("A1").Style.Font.Bold); // header bolded
        Assert.Equal(1, sheet.SheetView.SplitRow);      // header frozen
        Assert.Equal(36, sheet.Cell("B2").GetValue<int>()); // numbers typed
        Assert.Equal("Report", workbook.Properties.Title);
    }

    [Fact]
    public void Import_evaluates_a_leading_equals_as_a_formula()
    {
        var doc = new NeutralDoc(null, new List<NeutralBlock>
        {
            new(NeutralBlockKind.Heading, Level: 2, Runs: [new NeutralRun("Calc")]),
            new(NeutralBlockKind.Table, Rows: Rows(
                ["4", "6", "=A1+B1"])),
        });

        var file = NewFile("formula.xlsx");
        Handler.Create(Ctx(file));
        Handler.ImportNeutral(Ctx(file), doc);
        AssertValidatorClean(file);

        // The =A1+B1 cell becomes a real formula whose cached value evaluates to 10.
        var (formula, cached, _) = RawCell(file, "Calc", "C1");
        Assert.Equal("A1+B1", formula);
        Assert.Equal("10", cached);
    }

    [Fact]
    public void Import_collects_paragraphs_and_headings_into_a_notes_sheet()
    {
        var doc = new NeutralDoc(null, new List<NeutralBlock>
        {
            new(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Overview")]),
            new(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("A summary paragraph.")]),
            new(NeutralBlockKind.ListItem, Runs: [new NeutralRun("First point")]),
            new(NeutralBlockKind.Heading, Level: 2, Runs: [new NeutralRun("Q1")]),
            new(NeutralBlockKind.Table, HeaderRow: true, Rows: Rows(
                ["Month", "Sales"],
                ["Jan", "100"])),
        });

        var file = NewFile("notes.xlsx");
        Handler.Create(Ctx(file));
        var result = Handler.ImportNeutral(Ctx(file), doc);
        AssertValidatorClean(file);

        using var workbook = new XLWorkbook(file);

        // The standalone heading "Overview", the paragraph and the list item land
        // in a leading Notes sheet (one row each); "Q1" names the table sheet, so
        // it is NOT a note.
        Assert.True(workbook.TryGetWorksheet("Notes", out var notes));
        var noteTexts = notes.Column(1).CellsUsed().Select(c => c.GetString()).ToList();
        Assert.Equal(["Overview", "A summary paragraph.", "First point"], noteTexts);

        Assert.True(workbook.TryGetWorksheet("Q1", out var q1));
        Assert.Equal("Month", q1.Cell("A1").GetString());
    }

    [Fact]
    public void Import_names_inline_formatting_loss_in_dropped()
    {
        var doc = new NeutralDoc(null, new List<NeutralBlock>
        {
            new(NeutralBlockKind.Paragraph, Runs:
            [
                new NeutralRun("bold bit", Bold: true),
                new NeutralRun(" and a link", Href: "https://example.com"),
            ]),
        });

        var file = NewFile("lossy.xlsx");
        Handler.Create(Ctx(file));
        var result = Handler.ImportNeutral(Ctx(file), doc);

        Assert.Contains(result.Dropped, d => d.Contains("inline run formatting", StringComparison.Ordinal));
    }

    // ----- full round-trip ----------------------------------------------------

    [Fact]
    public void Multi_sheet_workbook_round_trips_sheet_names_and_values()
    {
        var source = CreateWorkbook(name: "src.xlsx", title: "Sales");
        Assert.True(EditOps(source, AddOp("/Totals", "sheet")).IsOk);
        Assert.True(EditOps(source, SetOp("/Sales/A1:B3", ("values", Grid(
            ["Region", "Units"],
            ["North", "10"],
            ["South", "20"])))).IsOk);
        // A formula on the second sheet: exports as its cached value, re-imports
        // as that same value (documented: convert transfers DISPLAY values).
        Assert.True(EditOps(source, SetOp("/Totals/A1", ("value", "=SUM(Sales!B2:B3)"))).IsOk);

        var doc = Handler.ExportNeutral(Ctx(source));
        Assert.Equal("30", doc.Blocks.Last(b => b.Kind == NeutralBlockKind.Table).Rows![0][0]);

        var dest = NewFile("dest.xlsx");
        Handler.Create(Ctx(dest));
        var result = Handler.ImportNeutral(Ctx(dest), doc);
        AssertValidatorClean(dest);
        Assert.Equal(2, result.BlocksWritten);

        using var workbook = new XLWorkbook(dest);
        Assert.True(workbook.TryGetWorksheet("Sales", out var sales));
        Assert.True(workbook.TryGetWorksheet("Totals", out var totals));
        Assert.Equal("Region", sales.Cell("A1").GetString());
        Assert.Equal(20, sales.Cell("B3").GetValue<int>());
        // The formula crossed as its cached value 30 (a number cell, not a formula).
        Assert.False(totals.Cell("A1").HasFormula);
        Assert.Equal(30, totals.Cell("A1").GetValue<int>());
        Assert.Equal("Sales", workbook.Properties.Title);
    }

    [Fact]
    public void Round_trip_is_deterministic()
    {
        var source = CreateWorkbook(name: "det.xlsx", title: "D");
        Assert.True(EditOps(source, SetOp("/D/A1:B3", ("values", Grid(
            ["K", "V"],
            ["a", "1"],
            ["b", "2"])))).IsOk);

        var first = NewFile("d1.xlsx");
        var second = NewFile("d2.xlsx");
        Handler.Create(Ctx(first));
        Handler.Create(Ctx(second));
        Handler.ImportNeutral(Ctx(first), Handler.ExportNeutral(Ctx(source)));
        Handler.ImportNeutral(Ctx(second), Handler.ExportNeutral(Ctx(source)));

        // Same input -> same cell content on every platform (sheetData equal).
        Assert.Equal(SheetDataXml(first, "D"), SheetDataXml(second, "D"));
    }

    [Fact]
    public void Empty_neutral_doc_still_produces_a_valid_workbook()
    {
        var doc = new NeutralDoc(null, new List<NeutralBlock>());
        var file = NewFile("empty.xlsx");
        Handler.Create(Ctx(file));
        var result = Handler.ImportNeutral(Ctx(file), doc);

        Assert.Equal(0, result.BlocksWritten);
        AssertValidatorClean(file);
        using var workbook = new XLWorkbook(file);
        Assert.True(workbook.Worksheets.Any());
    }

    // ----- helpers ------------------------------------------------------------

    private static JsonArray Grid(params string[][] rows)
    {
        var grid = new JsonArray();
        foreach (var row in rows)
        {
            var jsonRow = new JsonArray();
            foreach (var cell in row)
            {
                jsonRow.Add(cell);
            }

            grid.Add(jsonRow);
        }

        return grid;
    }

    private static IReadOnlyList<IReadOnlyList<string>> Rows(params string[][] rows) =>
        rows.Select(r => (IReadOnlyList<string>)[.. r]).ToList();

    /// <summary>Reads one sheet's raw <c>sheetData</c> XML (CRLF-normalized) for a deterministic compare.</summary>
    private static string SheetDataXml(string file, string sheetName)
    {
        using var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!
            .Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        var worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var sheetData = worksheetPart.Worksheet!
            .Descendants<DocumentFormat.OpenXml.Spreadsheet.SheetData>()
            .First();
        return sheetData.OuterXml.Replace("\r\n", "\n", StringComparison.Ordinal);
    }
}
