using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M6 flagship — in-place streaming write. The headline gate: the SAME set
/// of edits applied via BOTH the streaming path (a large file) and the DOM path
/// (a small twin) yields semantically identical workbooks. Plus the unit-level
/// correctness proofs: overwrite an existing cell, insert a new cell in column
/// order, insert a new row in row order, append shared strings, preserve
/// untouched cells/parts byte-for-byte, set the recalc flag for streamed
/// formulas, fall back transparently to the DOM when an op is not streamable,
/// and stay validator-clean throughout.
/// </summary>
[Collection("FileSizeEnv")]
public sealed class InPlaceStreamingWriteTests : ExcelTestBase
{
    // ----- helpers ------------------------------------------------------------

    /// <summary>The same edit batch the equality gate replays on both paths.</summary>
    private static EditOp[] EqualityEdits() =>
    [
        // overwrite an existing numeric cell
        SetOp("/Sheet1/A100", ("value", 123456)),
        // overwrite an existing shared-string cell with a NEW string (append to sst)
        SetOp("/Sheet1/D200", ("value", "BrandNewLabel")),
        // a brand-new cell to the RIGHT of the used columns (new column order)
        SetOp("/Sheet1/H300", ("value", 7.5)),
        // a brand-new cell to the LEFT-mid, forcing column-order insertion
        SetOp("/Sheet1/B305", ("value", "mid")),
        // a bulk 2D block overwriting + extending existing rows
        SetOp("/Sheet1/A400", ("values", new JsonArray(
            new JsonArray(1, "two", true),
            new JsonArray(4.5, "five", false)))),
        // brand-new rows far below the used range (new row order, out of band)
        SetOp("/Sheet1/A999000", ("values", new JsonArray(
            new JsonArray("far", 1),
            new JsonArray("down", 2)))),
        // a date (streamable in-place because the styles part exists)
        SetOp("/Sheet1/C500", ("value", "2024-05-01")),
        // a boolean
        SetOp("/Sheet1/C501", ("value", false)),
    ];

    private static double DomNumber(IXLWorksheet sheet, string address) => sheet.Cell(address).Value.GetNumber();

    /// <summary>Edits with an explicit context (carrying stream/dryRun args) and one-or-more ops.</summary>
    private Envelope EditCtx(CommandContext ctx, params EditOp[] ops) => Handler.Edit(ctx, ops);

    // ----- the equality gate (≥50 MB) -----------------------------------------

    [Fact]
    public void Streaming_and_dom_paths_produce_identical_workbooks_on_a_large_file()
    {
        using var guard = new EnvScope(FileSizeGuard.EnvVar, "100000");

        // A ~50 MB workbook, comfortably above the 15 MB in-place streaming-write
        // threshold so size alone activates the streaming path. The compressed size
        // varies a little by platform/runtime DEFLATE (~49.6 MB on CI x64, ~50+ MB
        // locally), so the precondition asserts well clear of the 15 MB threshold
        // rather than an exact round number.
        var streamed = Path.Combine(Dir, "streamed.xlsx");
        BigWorkbookGenerator.Generate(streamed, 400_000);
        Assert.True(
            new FileInfo(streamed).Length > 40L * 1024 * 1024,
            $"fixture is {new FileInfo(streamed).Length / (1024.0 * 1024.0):F1} MB; the streaming gate needs > 40 MB (>> the 15 MB activation threshold)");

        // The DOM twin: identical source bytes, but edited through the DOM path
        // by forcing stream=false on a copy the size guard would normally stream.
        var dom = Path.Combine(Dir, "dom.xlsx");
        File.Copy(streamed, dom);

        var edits = EqualityEdits();
        var streamedEnvelope = EditCtx(Ctx(streamed), edits); // size alone activates streaming
        Assert.True(streamedEnvelope.IsOk, streamedEnvelope.ToJson());
        Assert.True(Json(streamedEnvelope)["data"]!["streamed"]!.GetValue<bool>());

        var domEnvelope = EditCtx(Ctx(dom, ("stream", false)), edits); // forced DOM path
        Assert.True(domEnvelope.IsOk, domEnvelope.ToJson());
        Assert.Null(Json(domEnvelope)["data"]!["streamed"]);

        // Semantic equality: read both back through ClosedXML (the trusted
        // engine) and compare the edited cells plus a sample of untouched ones.
        using var streamedWb = new XLWorkbook(streamed);
        using var domWb = new XLWorkbook(dom);
        var s = streamedWb.Worksheet("Sheet1");
        var d = domWb.Worksheet("Sheet1");

        foreach (var address in new[] { "A100", "D200", "H300", "B305", "C500", "C501" })
        {
            Assert.True(
                d.Cell(address).Value.Equals(s.Cell(address).Value),
                $"{address}: dom={d.Cell(address).Value} streamed={s.Cell(address).Value}");
        }

        // The bulk block.
        for (var r = 400; r <= 401; r++)
        {
            for (var c = 1; c <= 3; c++)
            {
                Assert.True(
                    d.Cell(r, c).Value.Equals(s.Cell(r, c).Value),
                    $"({r},{c}): dom={d.Cell(r, c).Value} streamed={s.Cell(r, c).Value}");
            }
        }

        // Out-of-band new rows.
        Assert.Equal(d.Cell("A999000").GetString(), s.Cell("A999000").GetString());
        Assert.Equal(d.Cell("B999001").Value.GetNumber(), s.Cell("B999001").Value.GetNumber());

        // A wide sample of UNTOUCHED cells must be byte-equal in value.
        var random = new Random(20260613);
        for (var i = 0; i < 200; i++)
        {
            var row = random.Next(1, 399_999);
            if (row is 100 or 200 or 300 or 305 or 400 or 401 or 500 or 501)
            {
                continue;
            }

            Assert.Equal(DomNumber(d, "A" + row), DomNumber(s, "A" + row));
            Assert.Equal(d.Cell("D" + row).GetString(), s.Cell("D" + row).GetString());
            Assert.Equal(d.Cell("E" + row).FormulaA1, s.Cell("E" + row).FormulaA1);
        }

        AssertValidatorClean(streamed);
        AssertValidatorClean(dom);
    }

    // ----- unit correctness (small files, stream=true) ------------------------

    private string SeedSmall(string name = "book.xlsx")
    {
        var file = CreateWorkbook(name);
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray("h1", "h2", "h3"),
                new JsonArray(10, 20, 30),
                new JsonArray(40, 50, 60))))).IsOk);
        return file;
    }

    [Fact]
    public void Stream_arg_overwrites_an_existing_cell_in_place()
    {
        var file = SeedSmall();

        var envelope = EditCtx(Ctx(file, ("stream", true)), SetOp("/Sheet1/A2", ("value", 999)));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.True(Json(envelope)["data"]!["streamed"]!.GetValue<bool>());

        Assert.Equal(999.0, OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A2"))))["value"]!.GetValue<double>());
        // untouched neighbors survive
        Assert.Equal("h1", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<string>());
        Assert.Equal(60.0, OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C3"))))["value"]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Stream_inserts_a_new_cell_in_correct_column_order()
    {
        var file = SeedSmall();

        // Insert E2 (past C) and a new B-column already exists; add a cell at a
        // gap column (we add F2 then read the raw row to confirm column order).
        var envelope = EditCtx(Ctx(file, ("stream", true)),
            SetOp("/Sheet1/F2", ("value", "f")),
            SetOp("/Sheet1/D2", ("value", "d")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        // Raw cell order in row 2 must be ascending by column ref.
        var order = RawRowColumnRefs(file, "Sheet1", 2);
        Assert.Equal(["A2", "B2", "C2", "D2", "F2"], order);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Stream_inserts_new_rows_in_correct_row_order()
    {
        var file = SeedSmall();

        // Add a row ABOVE the data (row count starts at 1) and one far below.
        var envelope = EditCtx(Ctx(file, ("stream", true)),
            SetOp("/Sheet1/A100", ("value", "low")),
            SetOp("/Sheet1/A5", ("value", "mid")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var rows = RawRowIndexes(file, "Sheet1");
        Assert.Equal(rows.OrderBy(r => r).ToList(), rows); // strictly ascending
        Assert.Contains(5, rows);
        Assert.Contains(100, rows);
        Assert.Equal("mid", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A5"))))["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Stream_appends_new_shared_strings_and_reuses_existing_ones()
    {
        var file = SeedSmall();

        var envelope = EditCtx(Ctx(file, ("stream", true)),
            SetOp("/Sheet1/A2", ("value", "h1")),     // reuse existing sst entry
            SetOp("/Sheet1/A3", ("value", "brandNew"))); // append a new one
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Equal("h1", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A2"))))["value"]!.GetValue<string>());
        Assert.Equal("brandNew", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A3"))))["value"]!.GetValue<string>());

        // Both cells are stored as shared strings (the DOM-path representation):
        // raw t="s". The new label appended to the sst, the reused one did not.
        Assert.Equal("s", RawCellType(file, "Sheet1", "A2"));
        Assert.Equal("s", RawCellType(file, "Sheet1", "A3"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Streamed_formula_carries_no_cached_value_and_warns_formula_not_evaluated()
    {
        var file = SeedSmall();

        var envelope = EditCtx(Ctx(file, ("stream", true)), SetOp("/Sheet1/A4", ("value", "=A2+A3")));
        Assert.True(envelope.IsOk, envelope.ToJson());

        // Honest behavior: formula text saved, NO cached <v>, recalc flag set,
        // and a formula_not_evaluated warning.
        var raw = RawCell(file, "Sheet1", "A4");
        Assert.Equal("A2+A3", raw.Formula);
        Assert.Null(raw.CachedValue);
        Assert.Contains(
            envelope.Meta.Warnings ?? [],
            w => w.Code == ErrorCodes.FormulaNotEvaluated);
        Assert.True(FullRecalcOnLoad(file));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Streamed_bulk_2d_write_overwrites_and_extends()
    {
        var file = SeedSmall();

        var envelope = EditCtx(Ctx(file, ("stream", true)), SetOp("/Sheet1/B2", ("values", new JsonArray(
            new JsonArray(111, 222, 333),  // overwrites B2,C2 and adds D2
            new JsonArray(444, 555, 666)))));
        Assert.True(envelope.IsOk, envelope.ToJson());

        Assert.Equal(111.0, OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))))["value"]!.GetValue<double>());
        Assert.Equal(333.0, OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D2"))))["value"]!.GetValue<double>());
        Assert.Equal(666.0, OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D3"))))["value"]!.GetValue<double>());
        Assert.Equal(10.0, OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A2"))))["value"]!.GetValue<double>()); // A2 untouched
        AssertValidatorClean(file);
    }

    [Fact]
    public void Non_streamable_op_falls_back_to_the_dom_transparently()
    {
        var file = SeedSmall();

        // A bold style prop is not streamable; even with stream=true the batch
        // routes through the DOM and STILL succeeds (no "streamed":true marker).
        var envelope = EditCtx(Ctx(file, ("stream", true)),
            SetOp("/Sheet1/A2", ("value", 7), ("bold", true)));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Null(Json(envelope)["data"]!["streamed"]);

        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A2"))));
        Assert.Equal(7.0, cell["value"]!.GetValue<double>());
        Assert.True(cell["bold"]!.GetValue<bool>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Stream_arg_on_unknown_sheet_falls_back_for_the_canonical_error()
    {
        var file = SeedSmall();

        // The SAX planner can't resolve an unknown sheet, so it falls back to the
        // DOM path, which raises the canonical invalid_path with candidates.
        var envelope = EditCtx(Ctx(file, ("stream", true)), SetOp("/Ghost/A1", ("value", 1)));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1", envelope.Error.Candidates ?? [], StringComparer.Ordinal);
    }

    [Fact]
    public void Streamed_dry_run_reports_without_touching_the_file()
    {
        var file = SeedSmall();
        var before = File.ReadAllBytes(file);

        var envelope = EditCtx(Ctx(file, ("stream", true), ("dryRun", true)),
            SetOp("/Sheet1/A2", ("value", 999)));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.True(Json(envelope)["data"]!["dryRun"]!.GetValue<bool>());
        Assert.True(Json(envelope)["data"]!["streamed"]!.GetValue<bool>());

        Assert.True(before.AsSpan().SequenceEqual(File.ReadAllBytes(file))); // untouched
    }

    [Fact]
    public void Streamed_edit_is_undoable_via_snapshot()
    {
        var file = SeedSmall();
        Assert.True(EditCtx(Ctx(file, ("stream", true)), SetOp("/Sheet1/A2", ("value", 999))).IsOk);

        // A snapshot was taken (the file is restorable to its pre-edit state):
        // the newest snapshot captures the seeded state just before this edit.
        var snapshots = Snapshots.List(file);
        Assert.NotEmpty(snapshots);
        var snapshotCopy = Path.Combine(Dir, "snapshot-readback.xlsx"); // ClosedXML wants a .xlsx ext
        File.Copy(snapshots[^1].Path, snapshotCopy, overwrite: true);
        using var snapshotWb = new XLWorkbook(snapshotCopy);
        Assert.Equal(10.0, snapshotWb.Worksheet("Sheet1").Cell("A2").Value.GetNumber());
    }

    // ----- raw inspection helpers ---------------------------------------------

    /// <summary>The raw <c>t</c> attribute of a cell ("s" = shared string, "b" = bool, null = number).</summary>
    private static string? RawCellType(string file, string sheetName, string address)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var part = WorksheetPartOf(document, sheetName);
        var cell = part.Worksheet!.Descendants<S.Cell>().FirstOrDefault(c => c.CellReference?.Value == address);
        if (cell?.DataType is not { } dataType)
        {
            return null;
        }

        if (dataType.Value == S.CellValues.SharedString)
        {
            return "s";
        }

        if (dataType.Value == S.CellValues.Boolean)
        {
            return "b";
        }

        if (dataType.Value == S.CellValues.InlineString)
        {
            return "inlineStr";
        }

        return dataType.Value == S.CellValues.String ? "str" : dataType.InnerText;
    }

    private static List<string> RawRowColumnRefs(string file, string sheetName, int rowNumber)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var part = WorksheetPartOf(document, sheetName);
        var row = part.Worksheet!.Descendants<S.Row>().First(r => r.RowIndex?.Value == rowNumber);
        return row.Elements<S.Cell>().Select(c => c.CellReference!.Value!).ToList();
    }

    private static List<int> RawRowIndexes(string file, string sheetName)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var part = WorksheetPartOf(document, sheetName);
        return part.Worksheet!.Descendants<S.Row>()
            .Where(r => r.RowIndex?.Value is not null)
            .Select(r => (int)r.RowIndex!.Value)
            .ToList();
    }

    private static bool FullRecalcOnLoad(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        return document.WorkbookPart?.Workbook?.CalculationProperties?.FullCalculationOnLoad?.Value == true;
    }

    private static WorksheetPart WorksheetPartOf(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    }
}
