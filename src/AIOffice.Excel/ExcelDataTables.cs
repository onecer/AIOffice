using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// (1.4) What-if data tables. <c>{op:add, type:dataTable, path:/Sheet1/A1:C10,
/// props:{rowInput:"B1", colInput:"B2"}}</c> turns a rectangular range into a
/// one- or two-variable data table: the corner formula (top-left cell of the
/// range) is recomputed for every combination of the row-input values (across
/// the first row) and column-input values (down the first column), and the body
/// is filled with the computed results.
///
/// <para>The body carries the Excel data-table array construct
/// (<c>&lt;f t="dataTable" ref="…" r1="…" r2="…" dt2D="…" dtr="…"&gt;</c>) on its
/// top-left cell plus a cached value in every body cell, so Excel recomputes on
/// open while headless readers see results immediately. ClosedXML 0.105 has no
/// data-table model, so the construct + cached values are authored raw in a
/// post-save pass.</para>
///
/// <para>Addressing for get/remove is <c>/Sheet1/dataTable[i]</c> (1-based per
/// sheet, in row-major body order).</para>
/// </summary>
internal static class ExcelDataTables
{
    private static readonly IReadOnlyList<string> AddProps = ["rowInput", "colInput"];

    /// <summary>A computed data table queued for the post-save raw write.</summary>
    public sealed record Pending(
        string Sheet,
        string BodyRef,
        string? RowInput,
        string? ColInput,
        bool TwoDimensional,
        int FirstRow,
        int FirstColumn,
        IReadOnlyList<IReadOnlyList<XLCellValue>> Body);

    // ----- add ---------------------------------------------------------------

    /// <summary>Validates and applies an <c>add dataTable</c> op; returns the details entry and queues the raw write.</summary>
    public static object Add(XLWorkbook workbook, ExcelTarget target, EditOp op, int opIndex, List<Pending> queue)
    {
        if (target.Kind != ExcelTargetKind.Range)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add dataTable targets the table's full rectangular range like /Sheet1/A1:C10.",
                "The top-left cell holds the formula; the first row/column hold the input values. " +
                "Use {op:add, type:dataTable, path:/Sheet1/A1:C10, props:{rowInput:\"B1\", colInput:\"B2\"}}.");
        }

        var props = op.Props ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add dataTable needs props with rowInput and/or colInput.",
            "Use {op:add, type:dataTable, path:/Sheet1/A1:C10, props:{rowInput:\"B1\", colInput:\"B2\"}} for a two-variable table.");
        GuardProps(props, opIndex);

        var rowInput = StringProp(props, "rowInput");
        var colInput = StringProp(props, "colInput");
        if (rowInput is null && colInput is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add dataTable needs at least one of rowInput / colInput.",
                "A one-variable column table uses colInput; a one-variable row table uses rowInput; " +
                "a two-variable table uses both.");
        }

        var range = target.Range!;
        var address = range.RangeAddress;
        var rows = address.LastAddress.RowNumber - address.FirstAddress.RowNumber + 1;
        var columns = address.LastAddress.ColumnNumber - address.FirstAddress.ColumnNumber + 1;
        if (rows < 2 || columns < 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a data table needs at least a 2x2 range (got {rows}x{columns}).",
                "Row 1 / column A of the range hold the input axis; the body is everything else.");
        }

        var twoDimensional = rowInput is not null && colInput is not null;
        var sheet = target.Sheet;
        var firstRow = address.FirstAddress.RowNumber;
        var firstColumn = address.FirstAddress.ColumnNumber;

        // The body is the range minus its first row (row inputs) and first column
        // (column inputs); the corner holds the analyzed formula.
        var bodyFirstRow = firstRow + 1;
        var bodyFirstColumn = firstColumn + 1;
        var body = Compute(
            workbook, sheet, firstRow, firstColumn, rows, columns, rowInput, colInput, twoDimensional, opIndex);

        var bodyRef = string.Create(
            CultureInfo.InvariantCulture,
            $"{ExcelCharts.ColumnLetters(bodyFirstColumn)}{bodyFirstRow}:" +
            $"{ExcelCharts.ColumnLetters(address.LastAddress.ColumnNumber)}{address.LastAddress.RowNumber}");

        queue.Add(new Pending(
            sheet.Name, bodyRef, rowInput, colInput, twoDimensional, bodyFirstRow, bodyFirstColumn, body));

        // The index this table will get is its position among queued tables on the
        // sheet (1-based), matching the row-major body order the reader rebuilds.
        var indexOnSheet = queue.Count(p => string.Equals(p.Sheet, sheet.Name, StringComparison.OrdinalIgnoreCase));
        return new
        {
            op = "add",
            type = "dataTable",
            path = ExcelPaths.DataTablePath(sheet, indexOnSheet),
            range = address.ToString(),
            rowInput,
            colInput,
            twoDimensional,
        };
    }

    /// <summary>
    /// Computes the body by substituting each axis value into the input cell(s)
    /// and re-evaluating the corner formula IN PLACE (a full recalc per probe, so
    /// transitive dependencies of the inputs are picked up). The input cells (and
    /// the corner's cached value) are restored when done, so the live model is
    /// left exactly as it was — only the queued cached values carry the results.
    /// </summary>
    private static List<List<XLCellValue>> Compute(
        XLWorkbook workbook, IXLWorksheet sheet, int firstRow, int firstColumn, int rows, int columns,
        string? rowInput, string? colInput, bool twoDimensional, int opIndex)
    {
        var corner = sheet.Cell(firstRow, firstColumn);
        if (!corner.HasFormula)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: the data table's corner cell {corner.Address} must hold the formula to analyze.",
                "Set the top-left cell of the range to the formula (e.g. =PMT(B2/12,B1,-A1)) before adding the table.");
        }

        // Snapshot the axis input cells (and read their axis values up front so the
        // probing loop never reads a cell we are about to overwrite).
        var rowCell = rowInput is null ? null : sheet.Cell(rowInput);
        var colCell = colInput is null ? null : sheet.Cell(colInput);
        var rowOriginal = rowCell?.Value;
        var colOriginal = colCell?.Value;

        var colAxis = new List<XLCellValue>(rows - 1);
        for (var r = 1; r < rows; r++)
        {
            colAxis.Add(sheet.Cell(firstRow + r, firstColumn).Value);
        }

        var rowAxis = new List<XLCellValue>(columns - 1);
        for (var c = 1; c < columns; c++)
        {
            rowAxis.Add(sheet.Cell(firstRow, firstColumn + c).Value);
        }

        var body = new List<List<XLCellValue>>(rows - 1);
        try
        {
            for (var r = 0; r < rows - 1; r++)
            {
                var bodyRow = new List<XLCellValue>(columns - 1);
                for (var c = 0; c < columns - 1; c++)
                {
                    if (twoDimensional)
                    {
                        rowCell!.Value = rowAxis[c];
                        colCell!.Value = colAxis[r];
                    }
                    else if (colInput is not null)
                    {
                        // One-variable column table: the corner formula varies with
                        // the column-input cell down column A.
                        colCell!.Value = colAxis[r];
                    }
                    else
                    {
                        // One-variable row table: it varies with the row-input cell
                        // across row 1.
                        rowCell!.Value = rowAxis[c];
                    }

                    bodyRow.Add(EvaluateCorner(workbook, corner));
                }

                body.Add(bodyRow);
            }
        }
        finally
        {
            if (rowCell is not null && rowOriginal is { } ro)
            {
                rowCell.Value = ro;
            }

            if (colCell is not null && colOriginal is { } co)
            {
                colCell.Value = co;
            }

            workbook.RecalculateAllFormulas(); // restore the corner's original cached value
        }

        return body;
    }

    /// <summary>Forces a recalc and reads the corner cell's value; #N/A on any evaluation failure.</summary>
    private static XLCellValue EvaluateCorner(XLWorkbook workbook, IXLCell corner)
    {
        try
        {
            workbook.RecalculateAllFormulas();
            return corner.Value;
        }
        catch (Exception)
        {
            return XLError.NoValueAvailable;
        }
    }

    // ----- post-save raw write ----------------------------------------------

    /// <summary>
    /// Authors the queued data tables raw: the <c>{=TABLE(...)}</c> array construct
    /// on each body's top-left cell plus cached values across the body. ClosedXML
    /// left the body cells untouched, so this is the single writer.
    /// </summary>
    public static void ApplyAfterSave(string file, IReadOnlyList<Pending> tables)
    {
        if (tables.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart!;
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;

        foreach (var group in tables.GroupBy(t => t.Sheet, StringComparer.OrdinalIgnoreCase))
        {
            if (ExcelFormulaParts.SheetDataFor(workbookPart, group.Key) is not { } sheetData)
            {
                continue;
            }

            foreach (var table in group)
            {
                WriteTable(sheetData, table, ref sharedStrings, workbookPart);
            }

            sheetData.Ancestors<S.Worksheet>().First().Save();
        }

        ExcelFormulaParts.SetFullRecalc(workbookPart);
        workbookPart.Workbook!.Save();
    }

    private static void WriteTable(
        S.SheetData sheetData, Pending table, ref S.SharedStringTable? sharedStrings, WorkbookPart workbookPart)
    {
        for (var r = 0; r < table.Body.Count; r++)
        {
            var rowNumber = (uint)(table.FirstRow + r);
            var row = ExcelFormulaParts.EnsureRow(sheetData, rowNumber);
            for (var c = 0; c < table.Body[r].Count; c++)
            {
                var columnNumber = table.FirstColumn + c;
                var reference = ExcelFormulaParts.CellRef(columnNumber, (int)rowNumber);
                var cell = ExcelFormulaParts.EnsureCell(row, reference, columnNumber);
                WriteValue(cell, table.Body[r][c], ref sharedStrings, workbookPart);

                if (r == 0 && c == 0)
                {
                    // The data-table construct lives on the body's top-left cell;
                    // it carries no formula text (the t="dataTable" attributes drive
                    // the recompute), with the input-cell references and 1D/2D mode.
                    var formula = new S.CellFormula
                    {
                        FormulaType = S.CellFormulaValues.DataTable,
                        Reference = table.BodyRef,
                        CalculateCell = true,
                    };
                    if (table.TwoDimensional)
                    {
                        formula.DataTable2D = true;
                        formula.Input1Deleted = false;
                        formula.Input2Deleted = false;
                        formula.R1 = table.RowInput;
                        formula.R2 = table.ColInput;
                    }
                    else if (table.ColInput is not null)
                    {
                        // One-variable column table: row=false, r1 is the column input.
                        formula.DataTableRow = false;
                        formula.R1 = table.ColInput;
                    }
                    else
                    {
                        // One-variable row table.
                        formula.DataTableRow = true;
                        formula.R1 = table.RowInput;
                    }

                    cell.CellFormula = formula;
                }
                else
                {
                    cell.CellFormula = null; // body cells carry only the cached value
                }
            }
        }
    }

    private static void WriteValue(
        S.Cell cell, XLCellValue value, ref S.SharedStringTable? sharedStrings, WorkbookPart workbookPart)
    {
        switch (value.Type)
        {
            case XLDataType.Number:
                cell.DataType = S.CellValues.Number;
                cell.CellValue = new S.CellValue(value.GetNumber().ToString("R", CultureInfo.InvariantCulture));
                break;
            case XLDataType.Boolean:
                cell.DataType = S.CellValues.Boolean;
                cell.CellValue = new S.CellValue(value.GetBoolean() ? "1" : "0");
                break;
            case XLDataType.DateTime:
                cell.DataType = S.CellValues.Number;
                cell.CellValue = new S.CellValue(
                    value.GetDateTime().ToOADate().ToString("R", CultureInfo.InvariantCulture));
                break;
            case XLDataType.Error:
                cell.DataType = S.CellValues.Error;
                cell.CellValue = new S.CellValue(value.ToString(CultureInfo.InvariantCulture));
                break;
            case XLDataType.Blank:
                cell.DataType = null;
                cell.CellValue = null;
                break;
            default:
                cell.DataType = S.CellValues.SharedString;
                cell.CellValue = new S.CellValue(
                    ExcelFormulaParts.SharedStringIndex(value.GetText(), ref sharedStrings, workbookPart)
                        .ToString(CultureInfo.InvariantCulture));
                break;
        }
    }

    // ----- read back (get / remove) -----------------------------------------

    /// <summary>One data table read raw from the saved bytes (the t="dataTable" anchor + its cached body).</summary>
    public sealed record Info(
        int Index, string Sheet, string BodyRef, string? RowInput, string? ColInput, bool TwoDimensional);

    /// <summary>Reads every data table on a sheet, in row-major anchor order (1-based index).</summary>
    public static List<Info> ReadOnSheet(string file, string sheetName)
    {
        // Open through a shared read stream so this can run while ClosedXML still
        // holds the workbook open during op-application (the remove path).
        using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        return ReadOnSheet(document.WorkbookPart!, sheetName);
    }

    /// <summary>Reads the data tables on a sheet from an already-open workbook part (no second file open).</summary>
    private static List<Info> ReadOnSheet(WorkbookPart workbookPart, string sheetName)
    {
        if (ExcelFormulaParts.SheetDataFor(workbookPart, sheetName) is not { } sheetData)
        {
            return [];
        }

        var anchors = new List<(int Row, int Column, S.CellFormula Formula)>();
        foreach (var cell in sheetData.Descendants<S.Cell>())
        {
            if (cell.CellFormula is { } cf &&
                cf.FormulaType is not null && cf.FormulaType.Value == S.CellFormulaValues.DataTable &&
                cell.CellReference?.Value is { } reference)
            {
                anchors.Add((RowOf(reference), ExcelFormulaParts.ColumnOf(reference), cf));
            }
        }

        anchors.Sort((a, b) => a.Row != b.Row ? a.Row.CompareTo(b.Row) : a.Column.CompareTo(b.Column));
        var result = new List<Info>(anchors.Count);
        for (var i = 0; i < anchors.Count; i++)
        {
            var cf = anchors[i].Formula;
            var twoD = cf.DataTable2D?.Value == true;
            result.Add(new Info(
                i + 1,
                sheetName,
                cf.Reference?.Value ?? string.Empty,
                twoD ? cf.R1?.Value : (cf.DataTableRow?.Value == true ? cf.R1?.Value : null),
                twoD ? cf.R2?.Value : (cf.DataTableRow?.Value == false ? cf.R1?.Value : null),
                twoD));
        }

        return result;
    }

    /// <summary>Describes one data table for <c>get</c>.</summary>
    public static object Describe(Info info) => new
    {
        path = string.Create(CultureInfo.InvariantCulture, $"/{ExcelPaths.QuoteSheet(info.Sheet)}/dataTable[{info.Index}]"),
        kind = "dataTable",
        sheet = info.Sheet,
        body = info.BodyRef,
        rowInput = info.RowInput,
        colInput = info.ColInput,
        twoDimensional = info.TwoDimensional,
    };

    /// <summary>
    /// Removes the queued data tables from the saved bytes: strips the
    /// <c>t="dataTable"</c> construct from each anchor and clears every body cell's
    /// cached value. Indices are resolved per sheet against the on-disk file and
    /// applied highest-first so earlier removals never shift a later one.
    /// </summary>
    public static void RemoveAfterSave(string file, IReadOnlyList<(string Sheet, int Index)> removals)
    {
        if (removals.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart!;
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        _ = sharedStrings;

        foreach (var group in removals.GroupBy(r => r.Sheet, StringComparer.OrdinalIgnoreCase))
        {
            if (ExcelFormulaParts.SheetDataFor(workbookPart, group.Key) is not { } sheetData)
            {
                continue;
            }

            // Re-read once from the OPEN part (never re-open the file we hold), and
            // clear the requested bodies; the anchor order is stable within a sheet.
            var tables = ReadOnSheet(workbookPart, group.Key);
            var wantedRefs = group
                .Select(r => tables.FirstOrDefault(t => t.Index == r.Index)?.BodyRef)
                .Where(r => r is not null)
                .Select(r => r!)
                .ToList();

            foreach (var bodyRef in wantedRefs)
            {
                var (firstColumn, firstRow, lastColumn, lastRow) = ParseRef(bodyRef);
                foreach (var cell in sheetData.Descendants<S.Cell>().ToList())
                {
                    if (cell.CellReference?.Value is not { } reference)
                    {
                        continue;
                    }

                    var column = ExcelFormulaParts.ColumnOf(reference);
                    var row = RowOf(reference);
                    if (column >= firstColumn && column <= lastColumn && row >= firstRow && row <= lastRow)
                    {
                        cell.CellFormula = null;
                        cell.CellValue = null;
                        cell.DataType = null;
                    }
                }
            }

            sheetData.Ancestors<S.Worksheet>().First().Save();
        }

        workbookPart.Workbook!.Save();
    }

    // ----- helpers -----------------------------------------------------------

    private static void GuardProps(JsonObject props, int opIndex)
    {
        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown dataTable prop '{key}'.",
                    "Supported props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }
    }

    private static string? StringProp(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == System.Text.Json.JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static int RowOf(string cellReference)
    {
        var i = 0;
        while (i < cellReference.Length && char.IsLetter(cellReference[i]))
        {
            i++;
        }

        return int.TryParse(cellReference[i..], NumberStyles.Integer, CultureInfo.InvariantCulture, out var row)
            ? row
            : 0;
    }

    private static (int FirstColumn, int FirstRow, int LastColumn, int LastRow) ParseRef(string reference)
    {
        var parts = reference.Split(':');
        var first = parts[0];
        var last = parts.Length > 1 ? parts[1] : parts[0];
        return (
            ExcelFormulaParts.ColumnOf(first), RowOf(first),
            ExcelFormulaParts.ColumnOf(last), RowOf(last));
    }
}
