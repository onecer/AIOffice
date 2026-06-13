using System.Globalization;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The <see cref="INeutralConvertible"/> half of the xlsx handler (M9): the
/// content bridge <c>convert</c> uses to move a workbook to/from the
/// format-neutral document model.
///
/// <para>EXPORT — each worksheet becomes a <see cref="NeutralBlockKind.Heading"/>
/// block (the sheet name, level 2) followed by a <see cref="NeutralBlockKind.Table"/>
/// block built from the sheet's used range. Cells export their DISPLAY value
/// (number-format applied; formulas as their cached value text), so a converted
/// docx/pptx shows what Excel would show, not raw formula text. The first row is
/// marked <see cref="NeutralBlock.HeaderRow"/> only when it looks like a header
/// (all-text, no blanks, the rows below carry at least one non-text cell). The
/// title comes from the core Title property, falling back to the first sheet
/// name. Charts, pivots and images are not representable in the neutral model;
/// each becomes a <c>Dropped</c> note. A very large sheet is capped at
/// <see cref="MaxConvertRows"/> rows with a <c>Dropped</c> note.</para>
///
/// <para>IMPORT — writes a <see cref="NeutralDoc"/> INTO a freshly-created, empty
/// workbook. Each Table block becomes a worksheet (named from the preceding
/// Heading block when there is one, else <c>Sheet1</c>, <c>Sheet2</c>…); a
/// header row is bolded and frozen. Cell text is typed exactly like a bulk
/// write (numbers/bools/ISO dates/strings; a leading <c>=</c> becomes an
/// evaluated formula). Heading/Paragraph/ListItem/Image blocks that are NOT a
/// table are collected into a leading <c>Notes</c> sheet — one row per block —
/// so nothing is silently lost; inline formatting a flat cell cannot carry is
/// named in <c>Dropped</c>. The workbook's Title property is set from
/// <see cref="NeutralDoc.Title"/>.</para>
///
/// <para>Conversion is content-transfer and inherently lossy: <c>Dropped</c>
/// honestly names what the xlsx surface dropped (charts/pivots/images on export;
/// run-level bold/italic/links/heading levels on import).</para>
/// </summary>
public sealed partial class ExcelHandler : INeutralConvertible
{
    /// <summary>Row cap per sheet on export; rows beyond this are dropped with a note (keeps a huge workbook convertible).</summary>
    internal const int MaxConvertRows = 1000;

    // ----- export: xlsx -> neutral model --------------------------------------

    public NeutralDoc ExportNeutral(CommandContext ctx) => ExportNeutral(ctx, out _);

    /// <summary>
    /// Projects the workbook to the neutral model and, via <paramref name="dropped"/>,
    /// names the export-side losses the neutral model cannot carry (charts, pivot
    /// tables, images, and over-cap rows). The command layer folds these into the
    /// same <c>convert_lossy</c> report as the import-side
    /// <see cref="ImportResult.Dropped"/> — mirroring the pptx exporter — so an
    /// xlsx-source conversion no longer hides its losses behind body-only notes.
    /// </summary>
    public NeutralDoc ExportNeutral(CommandContext ctx, out IReadOnlyList<string> dropped)
    {
        var file = RequireFile(ctx, mustExist: true);
        using var workbook = OpenWorkbook(file);

        // Charts live in raw parts ClosedXML cannot enumerate; read them once.
        List<ChartInfo> charts;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(file, isEditable: false))
        {
            charts = ExcelCharts.Read(document);
        }

        var chartsBySheet = charts.ToLookup(c => c.SheetName, StringComparer.OrdinalIgnoreCase);
        var blocks = new List<NeutralBlock>();
        var lost = new List<string>();

        foreach (var sheet in workbook.Worksheets.OrderBy(ws => ws.Position))
        {
            blocks.Add(new NeutralBlock(
                NeutralBlockKind.Heading,
                Level: 2,
                Runs: [new NeutralRun(sheet.Name)]));

            var used = sheet.RangeUsed();
            if (used is not null)
            {
                var (rows, headerRow, capped) = BuildSheetGrid(sheet, used);
                blocks.Add(new NeutralBlock(
                    NeutralBlockKind.Table,
                    Rows: rows,
                    HeaderRow: headerRow));

                if (capped)
                {
                    var note = $"sheet '{sheet.Name}' exceeded {MaxConvertRows} rows; only the first {MaxConvertRows} were converted";
                    blocks.Add(DropNote(note));
                    lost.Add(note);
                }
            }

            // Charts/pivots/images are not representable in the neutral model.
            // Each becomes BOTH an honest note block in the converted body AND a
            // convert_lossy entry so an agent watching meta.warnings sees the loss.
            foreach (var chart in chartsBySheet[sheet.Name])
            {
                var label = chart.Title is { Length: > 0 } t ? $"'{t}' ({chart.Kind})" : chart.Kind;
                var note = $"chart {label} on sheet '{sheet.Name}' is not representable in the neutral model";
                blocks.Add(DropNote(note));
                lost.Add(note);
            }

            foreach (var pivot in sheet.PivotTables)
            {
                var note = $"pivot table '{pivot.Name}' on sheet '{sheet.Name}' is not representable in the neutral model";
                blocks.Add(DropNote(note));
                lost.Add(note);
            }

            if (sheet.Pictures.Count > 0)
            {
                var note = $"{sheet.Pictures.Count} image(s) on sheet '{sheet.Name}' are not representable in the neutral model";
                blocks.Add(DropNote(note));
                lost.Add(note);
            }
        }

        dropped = lost.Distinct(StringComparer.Ordinal).ToList();

        var title = NonEmpty(workbook.Properties.Title)
            ?? workbook.Worksheets.OrderBy(ws => ws.Position).FirstOrDefault()?.Name;

        return new NeutralDoc(title, blocks);
    }

    /// <summary>A neutral Paragraph block carrying an honest "this was dropped" note (prefixed for discoverability).</summary>
    private static NeutralBlock DropNote(string text) =>
        new(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("[dropped] " + text)]);

    /// <summary>
    /// Reads one sheet's used range into a text grid (display values), capped at
    /// <see cref="MaxConvertRows"/>, and decides whether the first row is a header.
    /// <c>Capped</c> is true when rows beyond the cap were dropped.
    /// </summary>
    private static (List<IReadOnlyList<string>> Rows, bool HeaderRow, bool Capped) BuildSheetGrid(
        IXLWorksheet sheet, IXLRange used)
    {
        var address = used.RangeAddress;
        var firstRow = address.FirstAddress.RowNumber;
        var lastRow = address.LastAddress.RowNumber;
        var firstColumn = address.FirstAddress.ColumnNumber;
        var lastColumn = address.LastAddress.ColumnNumber;

        var rowCount = lastRow - firstRow + 1;
        var capped = rowCount > MaxConvertRows;
        var cappedLastRow = capped ? firstRow + MaxConvertRows - 1 : lastRow;

        var rows = new List<IReadOnlyList<string>>();
        for (var r = firstRow; r <= cappedLastRow; r++)
        {
            var cells = new List<string>(lastColumn - firstColumn + 1);
            for (var c = firstColumn; c <= lastColumn; c++)
            {
                cells.Add(ExcelValues.SafeFormatted(sheet.Cell(r, c)));
            }

            rows.Add(cells);
        }

        return (rows, LooksLikeHeader(sheet, firstRow, cappedLastRow, firstColumn, lastColumn), capped);
    }

    /// <summary>
    /// True when the first row reads like a header: every cell is non-blank text,
    /// and at least one row below carries a non-text (number/bool/date) value —
    /// the classic "labels over data" shape. A single-row or all-text block is
    /// NOT treated as a header (HeaderRow=false), matching the contract.
    /// </summary>
    private static bool LooksLikeHeader(IXLWorksheet sheet, int firstRow, int lastRow, int firstColumn, int lastColumn)
    {
        if (lastRow <= firstRow)
        {
            return false; // a lone row is data, not a header
        }

        for (var c = firstColumn; c <= lastColumn; c++)
        {
            var cell = sheet.Cell(firstRow, c);
            if (cell.IsEmpty() || cell.DataType != XLDataType.Text || string.IsNullOrWhiteSpace(cell.GetText()))
            {
                return false;
            }
        }

        for (var r = firstRow + 1; r <= lastRow; r++)
        {
            for (var c = firstColumn; c <= lastColumn; c++)
            {
                var type = sheet.Cell(r, c).DataType;
                if (type is XLDataType.Number or XLDataType.Boolean or XLDataType.DateTime or XLDataType.TimeSpan)
                {
                    return true;
                }
            }
        }

        return false;
    }

    // ----- import: neutral model -> xlsx --------------------------------------

    public ImportResult ImportNeutral(CommandContext ctx, NeutralDoc doc)
    {
        var file = RequireFile(ctx, mustExist: true);
        var dropped = new List<string>();
        var blocksWritten = 0;

        // Pair each Table with the nearest preceding Heading for its sheet name;
        // everything else (headings without a table, paragraphs, list items,
        // images) goes to a leading Notes sheet so no content is lost.
        var notes = new List<string>();
        var sheets = new List<(string? Name, NeutralBlock Table)>();
        string? pendingHeading = null;

        foreach (var block in doc.Blocks)
        {
            switch (block.Kind)
            {
                case NeutralBlockKind.Table:
                    sheets.Add((pendingHeading, block));
                    pendingHeading = null;
                    blocksWritten++;
                    break;

                case NeutralBlockKind.Heading:
                    // A heading immediately followed by a table names that sheet;
                    // a heading with no table after it is kept as a note instead
                    // of being lost.
                    if (pendingHeading is not null)
                    {
                        notes.Add(pendingHeading);
                    }

                    pendingHeading = RunsToText(block.Runs);
                    break;

                case NeutralBlockKind.Paragraph:
                case NeutralBlockKind.ListItem:
                    FlushPendingHeading(ref pendingHeading, notes);
                    notes.Add(RunsToText(block.Runs));
                    blocksWritten++;
                    break;

                case NeutralBlockKind.Image:
                    FlushPendingHeading(ref pendingHeading, notes);
                    notes.Add(block.Alt is { Length: > 0 } alt
                        ? $"[image: {alt}] {block.Source}".TrimEnd()
                        : $"[image] {block.Source}".TrimEnd());
                    dropped.Add($"image '{block.Source}' written as a Notes-sheet reference (xlsx convert does not re-embed images)");
                    blocksWritten++;
                    break;
            }
        }

        FlushPendingHeading(ref pendingHeading, notes);

        // Any inline run formatting (bold/italic/underline/color/links) cannot
        // ride on a flat cell — name it once, honestly, if it was present.
        if (HasInlineFormatting(doc.Blocks))
        {
            dropped.Add("inline run formatting (bold/italic/underline/color/hyperlinks) is dropped: cells carry plain text");
        }

        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var warnings = new List<Warning>();

        using (var workbook = new XLWorkbook())
        {
            if (NonEmpty(doc.Title) is { } title)
            {
                workbook.Properties.Title = title;
            }

            if (notes.Count > 0)
            {
                WriteNotesSheet(workbook, notes, usedNames);
            }

            var autoIndex = 0;
            foreach (var (name, table) in sheets)
            {
                var sheetName = UniqueSheetName(name, ref autoIndex, usedNames);
                WriteTableSheet(workbook, sheetName, table, dropped);
            }

            // A workbook needs at least one sheet; an empty neutral doc still
            // produces a valid (single bare sheet) file.
            if (!workbook.Worksheets.Any())
            {
                AddSheetOrThrow(workbook, "Sheet1");
            }

            if (SaveWithCachedValues(workbook, file) is { } saveWarnings)
            {
                warnings.AddRange(saveWarnings);
            }
        }

        return new ImportResult(blocksWritten, dropped);
    }

    private static void FlushPendingHeading(ref string? pendingHeading, List<string> notes)
    {
        if (pendingHeading is not null)
        {
            notes.Add(pendingHeading);
            pendingHeading = null;
        }
    }

    /// <summary>Writes the leading Notes sheet: one row per non-table block text.</summary>
    private static void WriteNotesSheet(XLWorkbook workbook, List<string> notes, HashSet<string> usedNames)
    {
        var sheet = AddSheetOrThrowForImport(workbook, "Notes");
        usedNames.Add("Notes");
        for (var i = 0; i < notes.Count; i++)
        {
            sheet.Cell(i + 1, 1).Value = notes[i];
        }
    }

    /// <summary>Writes one Table block as a worksheet; a header row is bolded and frozen.</summary>
    private static void WriteTableSheet(XLWorkbook workbook, string sheetName, NeutralBlock table, List<string> dropped)
    {
        var sheet = AddSheetOrThrowForImport(workbook, sheetName);
        var rows = table.Rows ?? [];
        for (var r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            for (var c = 0; c < row.Count; c++)
            {
                if (c >= MaxSheetColumns || r >= MaxSheetRows)
                {
                    dropped.Add($"sheet '{sheetName}' content beyond row {MaxSheetRows} / column {MaxSheetColumns} was clipped");
                    continue;
                }

                WriteParsed(sheet.Cell(r + 1, c + 1), ExcelValues.Parse(row[c]));
            }
        }

        if (table.HeaderRow && rows.Count > 0)
        {
            var columns = rows[0].Count;
            if (columns > 0)
            {
                sheet.Range(1, 1, 1, Math.Min(columns, MaxSheetColumns)).Style.Font.Bold = true;
            }

            sheet.SheetView.FreezeRows(1);
        }
    }

    /// <summary>The sheet name for a table: the heading text (clamped to a legal name) or Sheet{n}; deduped.</summary>
    private static string UniqueSheetName(string? heading, ref int autoIndex, HashSet<string> usedNames)
    {
        var baseName = SanitizeSheetName(heading) ?? "Sheet" + (++autoIndex).ToString(CultureInfo.InvariantCulture);
        var candidate = baseName;
        var suffix = 2;
        while (!usedNames.Add(candidate))
        {
            // Keep the 31-char limit while appending the disambiguator.
            var tag = " (" + suffix.ToString(CultureInfo.InvariantCulture) + ")";
            var head = baseName.Length + tag.Length > 31 ? baseName[..(31 - tag.Length)] : baseName;
            candidate = head + tag;
            suffix++;
        }

        return candidate;
    }

    /// <summary>
    /// Clamps an arbitrary heading into a legal Excel sheet name (1-31 chars,
    /// none of <c>: \ / ? * [ ]</c>). Returns null for an empty/whitespace name
    /// so the caller falls back to Sheet{n}.
    /// </summary>
    private static string? SanitizeSheetName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return null;
        }

        var cleaned = new string(name.Select(c => c is ':' or '\\' or '/' or '?' or '*' or '[' or ']' ? ' ' : c).ToArray())
            .Trim();
        if (cleaned.Length == 0)
        {
            return null;
        }

        return cleaned.Length > 31 ? cleaned[..31].Trim() : cleaned;
    }

    private static string RunsToText(IReadOnlyList<NeutralRun>? runs) =>
        runs is null ? string.Empty : string.Concat(runs.Select(run => run.Text));

    private static bool HasInlineFormatting(IReadOnlyList<NeutralBlock> blocks) =>
        blocks.Any(b => b.Runs is { } runs && runs.Any(r =>
            r.Bold || r.Italic || r.Underline || r.Color is not null || r.Href is not null));

    private static string? NonEmpty(string? text) => string.IsNullOrEmpty(text) ? null : text;
}
