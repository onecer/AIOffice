using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// (1.4) Shared raw-OOXML helpers for the write-time formula evaluators
/// (dynamic-array spill, what-if data tables): locate/insert rows and cells in
/// document order, intern shared strings, write typed cached values, and set the
/// workbook's full-recalc-on-load flag. ClosedXML cannot model these constructs,
/// so they are authored directly on the saved bytes in a post-save pass — this is
/// the one place that shared cell-plumbing lives.
/// </summary>
internal static class ExcelFormulaParts
{
    /// <summary>
    /// Sets the workbook's full-recalc-on-load flag so Excel recomputes (and, for
    /// spills, re-expands) every formula on open while our cached values serve
    /// headless readers. Idempotent.
    /// </summary>
    public static void SetFullRecalc(WorkbookPart workbookPart)
    {
        if (workbookPart.Workbook is not { } workbook)
        {
            return;
        }

        workbook.CalculationProperties ??= new S.CalculationProperties { CalculationId = 0 };
        workbook.CalculationProperties.FullCalculationOnLoad = true;
        workbook.CalculationProperties.ForceFullCalculation = true;
    }

    /// <summary>Finds the sheetData for a sheet by name, or null when absent.</summary>
    public static S.SheetData? SheetDataFor(WorkbookPart workbookPart, string sheetName)
    {
        var sheetElement = workbookPart.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheetElement?.Id?.Value is not { } relId ||
            workbookPart.GetPartById(relId) is not WorksheetPart worksheetPart)
        {
            return null;
        }

        return worksheetPart.Worksheet?.GetFirstChild<S.SheetData>();
    }

    /// <summary>Locates an existing row by 1-based index or inserts a new one in document order.</summary>
    public static S.Row EnsureRow(S.SheetData sheetData, uint rowNumber)
    {
        var existing = sheetData.Elements<S.Row>().FirstOrDefault(r => r.RowIndex?.Value == rowNumber);
        if (existing is not null)
        {
            return existing;
        }

        var row = new S.Row { RowIndex = rowNumber };
        var after = sheetData.Elements<S.Row>().FirstOrDefault(r => r.RowIndex?.Value > rowNumber);
        if (after is null)
        {
            sheetData.Append(row);
        }
        else
        {
            sheetData.InsertBefore(row, after);
        }

        return row;
    }

    /// <summary>Locates an existing cell by reference or inserts a new one in column order.</summary>
    public static S.Cell EnsureCell(S.Row row, string reference, int columnNumber)
    {
        var existing = row.Elements<S.Cell>().FirstOrDefault(c => c.CellReference?.Value == reference);
        if (existing is not null)
        {
            return existing;
        }

        var cell = new S.Cell { CellReference = reference };
        var after = row.Elements<S.Cell>().FirstOrDefault(c => ColumnOf(c.CellReference?.Value) > columnNumber);
        if (after is null)
        {
            row.Append(cell);
        }
        else
        {
            row.InsertBefore(cell, after);
        }

        return cell;
    }

    /// <summary>Interns a string in the shared-string table (creating the part if needed), returning its index.</summary>
    public static int SharedStringIndex(string text, ref S.SharedStringTable? table, WorkbookPart workbookPart)
    {
        if (table is null)
        {
            var part = workbookPart.SharedStringTablePart ?? workbookPart.AddNewPart<SharedStringTablePart>();
            part.SharedStringTable ??= new S.SharedStringTable();
            table = part.SharedStringTable;
        }

        var index = 0;
        foreach (var item in table.Elements<S.SharedStringItem>())
        {
            if (item.InnerText == text)
            {
                return index;
            }

            index++;
        }

        table.Append(new S.SharedStringItem(new S.Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return index;
    }

    /// <summary>1-based column number for a cell reference (max value for null/blank, so it sorts last).</summary>
    public static int ColumnOf(string? cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
        {
            return int.MaxValue;
        }

        var column = 0;
        foreach (var ch in cellReference)
        {
            if (!char.IsLetter(ch))
            {
                break;
            }

            column = (column * 26) + (char.ToUpperInvariant(ch) - 'A' + 1);
        }

        return column;
    }

    /// <summary>A1-style reference from a 1-based column and row.</summary>
    public static string CellRef(int column, int row) =>
        ExcelCharts.ColumnLetters(column) + row.ToString(CultureInfo.InvariantCulture);
}
