using System.Globalization;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// The M4 bulk-write performance path: when one <c>set values</c> op writes
/// more than <see cref="StreamingThresholdCells"/> cells into a BARE sheet,
/// the grid is not pushed through the ClosedXML DOM. Instead, after the normal
/// ClosedXML save, the sheet's part is rewritten with a SAX writer
/// (<see cref="OpenXmlWriter"/>): every non-sheetData element of the part
/// (sheetViews, sheetFormatPr, pageMargins, even a drawing reference) is
/// preserved verbatim and a fresh sheetData is streamed in bounded memory.
///
/// Honest capability notes:
/// <list type="bullet">
/// <item>Strings are written as inline strings (no shared-string table
/// management); numbers and booleans as native cell values; formulas as
/// formula text. DATE/TIME values are NOT streamable (they need a styles-part
/// number format) — a grid containing them falls back to the DOM path, which
/// is correct but slower. An equality test pins that both paths produce the
/// same workbook content for the same input.</item>
/// <item>Grids containing formulas get their cached values from one extra
/// pass through the normal ClosedXML pipeline after the SAX write (that pass
/// also folds inline strings into the shared-string table).</item>
/// <item>This is a WRITE-INTO-BARE-SHEET fast path only. In-place streaming
/// edits of existing big sheets remain M5; they still load the full DOM.</item>
/// </list>
/// </summary>
internal static class ExcelBulkWrites
{
    /// <summary>Cell count above which a bare-sheet bulk write streams.</summary>
    public const int StreamingThresholdCells = 50_000;

    /// <summary>One queued streamed write (the sheet reference survives batch renames).</summary>
    internal sealed record Pending(
        IXLWorksheet Sheet, int FirstRow, int FirstColumn, List<List<ExcelValues.ParsedValue>> Grid)
    {
        public bool HasFormulas => Grid.Any(row => row.Any(v => v.IsFormula));
    }

    public static int CellCount(List<List<ExcelValues.ParsedValue>> grid) => grid.Sum(row => row.Count);

    /// <summary>
    /// True when the SAX path can represent every value: blanks, numbers,
    /// booleans, text and formulas. Dates/timespans need styles → DOM path.
    /// </summary>
    public static bool IsStreamable(List<List<ExcelValues.ParsedValue>> grid) =>
        grid.All(row => row.All(v =>
            v.IsFormula ||
            v.Value.Type is XLDataType.Blank or XLDataType.Number or XLDataType.Boolean or XLDataType.Text));

    /// <summary>Rewrites each target sheet's part with the queued grids (merged per sheet, in batch order).</summary>
    public static void ApplyAfterSave(string file, IReadOnlyList<Pending> writes)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart ?? throw Corrupt("the package has no workbook part");

        foreach (var group in writes.GroupBy(w => w.Sheet.Name, StringComparer.OrdinalIgnoreCase))
        {
            var sheetElement = workbookPart.Workbook
                ?.Descendants<S.Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, group.Key, StringComparison.OrdinalIgnoreCase))
                ?? throw Corrupt($"sheet '{group.Key}' is missing from workbook.xml");
            if (sheetElement.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart part)
            {
                throw Corrupt($"sheet '{group.Key}' points at a missing worksheet part");
            }

            RewriteSheetData(part, Merge(group));
        }
    }

    private static AiofficeException Corrupt(string what) => new(
        ErrorCodes.InternalError,
        $"Streamed bulk write failed: {what}.",
        "The edit was rolled back; re-run it. If this persists, report a bug with the workbook.");

    /// <summary>Later writes win cell-by-cell, mirroring sequential DOM semantics.</summary>
    private static SortedDictionary<int, SortedDictionary<int, ExcelValues.ParsedValue>> Merge(
        IEnumerable<Pending> writes)
    {
        var rows = new SortedDictionary<int, SortedDictionary<int, ExcelValues.ParsedValue>>();
        foreach (var write in writes)
        {
            for (var r = 0; r < write.Grid.Count; r++)
            {
                var rowNumber = write.FirstRow + r;
                if (!rows.TryGetValue(rowNumber, out var cells))
                {
                    rows[rowNumber] = cells = [];
                }

                var row = write.Grid[r];
                for (var c = 0; c < row.Count; c++)
                {
                    cells[write.FirstColumn + c] = row[c];
                }
            }
        }

        return rows;
    }

    /// <summary>
    /// Replaces the part's sheetData while preserving every other child of
    /// <c>&lt;worksheet&gt;</c> byte-for-byte (cloned through the part reader —
    /// the typed DOM is never loaded, so nothing re-saves over the SAX output).
    /// </summary>
    private static void RewriteSheetData(
        WorksheetPart part, SortedDictionary<int, SortedDictionary<int, ExcelValues.ParsedValue>> rows)
    {
        List<OpenXmlAttribute> rootAttributes = [];
        List<KeyValuePair<string, string>> rootNamespaces = [];
        var preElements = new List<OpenXmlElement>();
        var postElements = new List<OpenXmlElement>();

        using (var reader = new OpenXmlPartReader(part))
        {
            while (reader.Read())
            {
                if (reader.ElementType != typeof(S.Worksheet) || !reader.IsStartElement)
                {
                    continue;
                }

                rootAttributes = [.. reader.Attributes];
                rootNamespaces = [.. reader.NamespaceDeclarations];
                if (!reader.ReadFirstChild())
                {
                    break;
                }

                var seenSheetData = false;
                do
                {
                    if (!reader.IsStartElement)
                    {
                        continue;
                    }

                    if (reader.ElementType == typeof(S.SheetData))
                    {
                        seenSheetData = true; // skipped: ReadNextSibling jumps the subtree
                        continue;
                    }

                    (seenSheetData ? postElements : preElements).Add(reader.LoadCurrentElement()!);
                }
                while (reader.ReadNextSibling());

                break;
            }
        }

        UpdateDimension(preElements, rows);

        using var writer = OpenXmlWriter.Create(part);
        writer.WriteStartElement(new S.Worksheet(), rootAttributes, rootNamespaces);
        foreach (var element in preElements)
        {
            writer.WriteElement(element);
        }

        writer.WriteStartElement(new S.SheetData());
        foreach (var (rowNumber, cells) in rows)
        {
            writer.WriteStartElement(new S.Row { RowIndex = (uint)rowNumber });
            foreach (var (columnNumber, value) in cells)
            {
                if (BuildCell(rowNumber, columnNumber, value) is { } cell)
                {
                    writer.WriteElement(cell);
                }
            }

            writer.WriteEndElement(); // row
        }

        writer.WriteEndElement(); // sheetData
        foreach (var element in postElements)
        {
            writer.WriteElement(element);
        }

        writer.WriteEndElement(); // worksheet
    }

    /// <summary>Points the cloned <c>&lt;dimension&gt;</c> at the written extent (the sheet was bare).</summary>
    private static void UpdateDimension(
        List<OpenXmlElement> preElements,
        SortedDictionary<int, SortedDictionary<int, ExcelValues.ParsedValue>> rows)
    {
        if (preElements.OfType<S.SheetDimension>().FirstOrDefault() is not { } dimension || rows.Count == 0)
        {
            return;
        }

        var firstRow = rows.Keys.First();
        var lastRow = rows.Keys.Last();
        var firstColumn = int.MaxValue;
        var lastColumn = 0;
        foreach (var cells in rows.Values)
        {
            firstColumn = Math.Min(firstColumn, cells.Keys.First());
            lastColumn = Math.Max(lastColumn, cells.Keys.Last());
        }

        var start = ExcelCharts.ColumnLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture);
        var end = ExcelCharts.ColumnLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture);
        dimension.Reference = start == end ? start : start + ":" + end;
    }

    /// <summary>One raw cell; null for blanks (a bare sheet has nothing to clear).</summary>
    private static S.Cell? BuildCell(int rowNumber, int columnNumber, ExcelValues.ParsedValue value)
    {
        var reference = ExcelCharts.ColumnLetters(columnNumber) + rowNumber.ToString(CultureInfo.InvariantCulture);
        if (value.IsFormula)
        {
            var text = value.Formula!;
            return new S.Cell
            {
                CellReference = reference,
                CellFormula = new S.CellFormula(text.StartsWith('=') ? text[1..] : text),
            };
        }

        switch (value.Value.Type)
        {
            case XLDataType.Number:
                return new S.Cell { CellReference = reference, CellValue = new S.CellValue(value.Value.GetNumber()) };

            case XLDataType.Boolean:
                return new S.Cell
                {
                    CellReference = reference,
                    DataType = S.CellValues.Boolean,
                    CellValue = new S.CellValue(value.Value.GetBoolean() ? "1" : "0"),
                };

            case XLDataType.Text:
            {
                var text = new S.Text(value.Value.GetText());
                if (text.Text.Length > 0 && (char.IsWhiteSpace(text.Text[0]) || char.IsWhiteSpace(text.Text[^1])))
                {
                    text.Space = SpaceProcessingModeValues.Preserve;
                }

                return new S.Cell
                {
                    CellReference = reference,
                    DataType = S.CellValues.InlineString,
                    InlineString = new S.InlineString(text),
                };
            }

            default: // blank — IsStreamable already excluded dates/timespans
                return null;
        }
    }
}
