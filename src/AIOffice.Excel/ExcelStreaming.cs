using System.Globalization;
using System.Text;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// The M3 big-workbook read path: a SAX scan over the raw worksheet XML
/// (<see cref="OpenXmlPartReader"/>) that never materializes the workbook DOM,
/// so multi-hundred-MB files answer <c>read --view stats|text</c> and
/// <c>get</c> of a cell/range in bounded memory, stopping as soon as the
/// requested window has been served.
///
/// Honest capability notes:
/// <list type="bullet">
/// <item>Streaming is READ-ONLY. Mutating ops (edit/template) still load the
/// full workbook through ClosedXML — editing huge files works but is slow and
/// memory-hungry.</item>
/// <item>Streamed reads report values exactly as stored: shared strings are
/// resolved (only the needed entries are read, stopping early), formula cells
/// return their cached <c>&lt;v&gt;</c> value (nothing is evaluated), and
/// date-formatted numbers appear as raw serial numbers because number formats
/// live in the styles part the scan never opens. Styling fields
/// (numberFormat/bold/italic/merged) are omitted.</item>
/// <item>The whole-workbook text view materializes the shared-string table as
/// a flat list (lazily, on the first shared-string cell) — still far cheaper
/// than the DOM, but proportional to the table size.</item>
/// </list>
/// Activation: file size over <see cref="ThresholdBytes"/>, or the explicit
/// <c>stream=true</c> arg. Streamed envelopes carry <c>"streamed": true</c>.
/// </summary>
internal static class ExcelStreaming
{
    /// <summary>Files larger than this stream automatically (compressed package size).</summary>
    public const long ThresholdBytes = 20L * 1024 * 1024;

    public static bool IsLarge(string file) => new FileInfo(file).Length > ThresholdBytes;

    // ----- sheet catalog (workbook.xml is tiny; loading it is fine) ----------

    /// <summary>One sheet as listed in workbook.xml, in workbook order.</summary>
    internal sealed record SheetEntry(string Name, string RelationshipId, int Position, bool Visible);

    private static List<SheetEntry> Sheets(SpreadsheetDocument document)
    {
        var result = new List<SheetEntry>();
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return result;
        }

        var position = 0;
        foreach (var sheet in workbookPart.Workbook.Descendants<S.Sheet>())
        {
            if (sheet.Name?.Value is not { } name || sheet.Id?.Value is not { } relationshipId)
            {
                continue;
            }

            position++;
            var visible = sheet.State is null || sheet.State.Value == S.SheetStateValues.Visible;
            result.Add(new SheetEntry(name, relationshipId, position, visible));
        }

        return result;
    }

    private static SheetEntry SheetOrThrow(List<SheetEntry> sheets, string name)
    {
        var found = sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
        if (found is not null)
        {
            return found;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No sheet named '{name}' exists in the workbook.",
            "Sheet names are matched case-insensitively; pick one of the candidates.",
            candidates: [.. sheets
                .OrderBy(s => ExcelPaths.Levenshtein(name, s.Name))
                .ThenBy(s => s.Position)
                .Select(s => "/" + ExcelPaths.QuoteSheet(s.Name))
                .Take(5)]);
    }

    private static WorksheetPart PartOf(SpreadsheetDocument document, SheetEntry sheet)
    {
        if (document.WorkbookPart!.GetPartById(sheet.RelationshipId) is WorksheetPart part)
        {
            return part;
        }

        throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            $"Sheet '{sheet.Name}' points at a missing worksheet part.",
            "The package is internally inconsistent; re-export the file from its source application.");
    }

    // ----- raw cell model -----------------------------------------------------

    private enum RawKind
    {
        Blank,
        Number,
        Text,
        SharedString,
        Boolean,
        Error,
        DateIso,
    }

    /// <summary>One cell exactly as stored in sheet XML (shared strings unresolved).</summary>
    private readonly record struct RawCell(int Row, int Column, RawKind Kind, string? Value, string? Formula)
    {
        public bool HasContent => Formula is not null || Value is not null || Kind is RawKind.SharedString;
    }

    // ----- the scan core ------------------------------------------------------

    /// <summary>
    /// Streams one worksheet's cells within a 1-based inclusive window, in
    /// document order, without loading the sheet DOM. Stops as soon as the scan
    /// passes <paramref name="lastRow"/> or <paramref name="onCell"/> returns
    /// false. Returns the <c>&lt;dimension&gt;</c> ref when the sheet has one.
    /// </summary>
    private static string? Scan(
        WorksheetPart part, int firstRow, int lastRow, int firstColumn, int lastColumn, Func<RawCell, bool> onCell)
    {
        string? dimension = null;
        using var reader = new OpenXmlPartReader(part);
        while (reader.Read())
        {
            if (reader.ElementType == typeof(S.SheetDimension) && reader.IsStartElement)
            {
                dimension = Attribute(reader, "ref");
                continue;
            }

            if (reader.ElementType == typeof(S.SheetData) && reader.IsStartElement)
            {
                ScanRows(reader, firstRow, lastRow, firstColumn, lastColumn, onCell);
                break; // the window is served; everything after sheetData is irrelevant
            }
        }

        return dimension;
    }

    private static void ScanRows(
        OpenXmlPartReader reader, int firstRow, int lastRow, int firstColumn, int lastColumn, Func<RawCell, bool> onCell)
    {
        if (!reader.ReadFirstChild())
        {
            return;
        }

        var impliedRow = 0;
        do
        {
            if (reader.ElementType != typeof(S.Row) || !reader.IsStartElement)
            {
                continue;
            }

            var rowNumber = IntAttribute(reader, "r") ?? impliedRow + 1;
            impliedRow = rowNumber;
            if (rowNumber > lastRow)
            {
                return; // rows are stored in ascending order; the window is done
            }

            if (rowNumber < firstRow || !reader.ReadFirstChild())
            {
                continue; // ReadNextSibling skips the row's subtree
            }

            var impliedColumn = 0;
            do
            {
                if (reader.ElementType != typeof(S.Cell) || !reader.IsStartElement)
                {
                    continue;
                }

                var cell = (S.Cell)reader.LoadCurrentElement()!;
                var column = ColumnOfRef(cell.CellReference?.Value) ?? impliedColumn + 1;
                impliedColumn = column;
                if (column < firstColumn || column > lastColumn)
                {
                    continue;
                }

                if (!onCell(Capture(rowNumber, column, cell)))
                {
                    return;
                }
            }
            while (reader.ReadNextSibling());
        }
        while (reader.ReadNextSibling());
    }

    private static RawCell Capture(int row, int column, S.Cell cell)
    {
        var formula = cell.CellFormula?.Text is { Length: > 0 } text ? text : null;
        var value = cell.CellValue?.Text;
        var dataType = cell.DataType?.Value;

        if (dataType == S.CellValues.SharedString)
        {
            return new RawCell(row, column, RawKind.SharedString, value, formula);
        }

        if (dataType == S.CellValues.InlineString)
        {
            return new RawCell(row, column, RawKind.Text, ItemText(cell.InlineString), formula);
        }

        if (dataType == S.CellValues.String)
        {
            return new RawCell(row, column, RawKind.Text, value, formula);
        }

        if (dataType == S.CellValues.Boolean)
        {
            return new RawCell(row, column, RawKind.Boolean, value, formula);
        }

        if (dataType == S.CellValues.Error)
        {
            return new RawCell(row, column, RawKind.Error, value, formula);
        }

        if (dataType == S.CellValues.Date)
        {
            return new RawCell(row, column, RawKind.DateIso, value, formula);
        }

        return new RawCell(row, column, value is null ? RawKind.Blank : RawKind.Number, value, formula);
    }

    private static string? ItemText(OpenXmlElement? item) =>
        item is null
            ? null
            : string.Concat(item
                .Descendants<S.Text>()
                .Where(t => t.Ancestors<S.PhoneticRun>().FirstOrDefault() is null)
                .Select(t => t.Text));

    private static string? Attribute(OpenXmlPartReader reader, string localName)
    {
        foreach (var attribute in reader.Attributes)
        {
            if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal))
            {
                return attribute.Value;
            }
        }

        return null;
    }

    private static int? IntAttribute(OpenXmlPartReader reader, string localName) =>
        Attribute(reader, localName) is { } text &&
        int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;

    /// <summary>The 1-based column number of an A1-style ref ("C12" → 3); null when malformed.</summary>
    private static int? ColumnOfRef(string? reference)
    {
        if (string.IsNullOrEmpty(reference))
        {
            return null;
        }

        var n = 0;
        foreach (var c in reference)
        {
            if (c is >= 'A' and <= 'Z')
            {
                n = (n * 26) + (c - 'A' + 1);
            }
            else if (c is >= 'a' and <= 'z')
            {
                n = (n * 26) + (c - 'a' + 1);
            }
            else
            {
                break;
            }
        }

        return n == 0 ? null : n;
    }

    // ----- shared strings (streamed, early-stop) ------------------------------

    /// <summary>Resolves only the needed shared-string indices, stopping at the highest one.</summary>
    private static Dictionary<int, string> ResolveSharedStrings(SpreadsheetDocument document, HashSet<int> needed)
    {
        var result = new Dictionary<int, string>();
        if (needed.Count == 0 || document.WorkbookPart?.SharedStringTablePart is not { } part)
        {
            return result;
        }

        var maxIndex = needed.Max();
        using var reader = new OpenXmlPartReader(part);
        var index = -1;
        while (reader.Read())
        {
            if (reader.ElementType != typeof(S.SharedStringItem) || !reader.IsStartElement)
            {
                continue;
            }

            index++;
            if (index > maxIndex)
            {
                break;
            }

            if (!needed.Contains(index))
            {
                continue;
            }

            result[index] = ItemText(reader.LoadCurrentElement()) ?? string.Empty;
            if (result.Count == needed.Count)
            {
                break;
            }
        }

        return result;
    }

    /// <summary>The whole shared-string table as a flat list (text view needs arbitrary entries inline).</summary>
    private static List<string> AllSharedStrings(SpreadsheetDocument document)
    {
        var result = new List<string>();
        if (document.WorkbookPart?.SharedStringTablePart is not { } part)
        {
            return result;
        }

        using var reader = new OpenXmlPartReader(part);
        while (reader.Read())
        {
            if (reader.ElementType == typeof(S.SharedStringItem) && reader.IsStartElement)
            {
                result.Add(ItemText(reader.LoadCurrentElement()) ?? string.Empty);
            }
        }

        return result;
    }

    // ----- wire conversion ----------------------------------------------------

    private static object? WireValue(RawCell cell, Func<int, string> sharedString) => cell.Kind switch
    {
        RawKind.Blank => null,
        RawKind.Number when cell.Value is { } v &&
            double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) => d,
        RawKind.Number => cell.Value,
        RawKind.Boolean => cell.Value is "1" or "true",
        RawKind.SharedString => SharedStringOf(cell, sharedString),
        _ => cell.Value,
    };

    private static string SharedStringOf(RawCell cell, Func<int, string> sharedString) =>
        int.TryParse(cell.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var index)
            ? sharedString(index)
            : string.Empty;

    private static string WireType(RawCell cell) => cell.Kind switch
    {
        RawKind.Blank => "blank",
        RawKind.Number => "number",
        RawKind.Boolean => "boolean",
        RawKind.Error => "error",
        RawKind.DateIso => "dateTime",
        _ => "text",
    };

    /// <summary>Display text of a raw cell (no number formats: values appear as stored).</summary>
    private static string WireText(RawCell cell, Func<int, string> sharedString) => cell.Kind switch
    {
        RawKind.Blank => string.Empty,
        RawKind.Boolean => cell.Value is "1" or "true" ? "TRUE" : "FALSE",
        RawKind.SharedString => SharedStringOf(cell, sharedString),
        _ => cell.Value ?? string.Empty,
    };

    // ----- read --view stats --------------------------------------------------

    public static object ReadStats(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var sheets = new List<object>();
        var totalCells = 0L;
        var totalFormulas = 0L;
        foreach (var entry in Sheets(document))
        {
            var cellCount = 0L;
            var formulaCount = 0L;
            int minRow = int.MaxValue, minColumn = int.MaxValue, maxRow = 0, maxColumn = 0;
            Scan(PartOf(document, entry), 1, int.MaxValue, 1, int.MaxValue, cell =>
            {
                if (!cell.HasContent)
                {
                    return true; // style-only cells do not count as used
                }

                cellCount++;
                if (cell.Formula is not null)
                {
                    formulaCount++;
                }

                minRow = Math.Min(minRow, cell.Row);
                minColumn = Math.Min(minColumn, cell.Column);
                maxRow = Math.Max(maxRow, cell.Row);
                maxColumn = Math.Max(maxColumn, cell.Column);
                return true;
            });

            sheets.Add(new
            {
                name = entry.Name,
                position = entry.Position,
                usedRange = cellCount == 0 ? null : RangeText(minColumn, minRow, maxColumn, maxRow),
                cellCount,
                formulaCount,
            });
            totalCells += cellCount;
            totalFormulas += formulaCount;
        }

        return new
        {
            kind = "xlsx",
            streamed = true,
            sheets,
            totals = new { sheets = sheets.Count, cells = totalCells, formulas = totalFormulas },
        };
    }

    // ----- get of a cell or range ----------------------------------------------

    /// <summary>
    /// True when <paramref name="pathText"/> addresses a cell or range — the
    /// targets the streaming path can serve. Parsing failures throw
    /// <c>invalid_path</c> exactly like the DOM resolver would.
    /// </summary>
    public static bool TryParseCellOrRange(string pathText, out string sheetName, out CellRef start, out CellRef end)
    {
        sheetName = string.Empty;
        start = end = default;
        if (pathText.Contains("[@name=", StringComparison.OrdinalIgnoreCase))
        {
            return false; // pivot/name id-forms never stream
        }

        var path = DocPath.Parse(pathText); // throws invalid_path with the grammar hint
        if (path.Segments.Count != 2)
        {
            return false;
        }

        var first = path.Segments[0];
        var name = first.Kind switch
        {
            PathSegmentKind.Name => first.Name,
            PathSegmentKind.Element when first.Index is null && first.Id is null => first.Name,
            _ => null,
        };
        if (name is null)
        {
            return false;
        }

        var second = path.Segments[1];
        switch (second.Kind)
        {
            case PathSegmentKind.Cell:
                sheetName = name;
                start = end = second.Start!.Value;
                return true;
            case PathSegmentKind.Range:
                sheetName = name;
                start = second.Start!.Value;
                end = second.End!.Value;
                return true;
            default:
                return false;
        }
    }

    public static object GetCell(string file, string sheetName, CellRef reference)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var entry = SheetOrThrow(Sheets(document), sheetName);
        var column = reference.ColumnNumber;

        RawCell found = default;
        var hit = false;
        Scan(PartOf(document, entry), reference.Row, reference.Row, column, column, cell =>
        {
            found = cell;
            hit = true;
            return false; // first (only) match serves the request
        });

        var cellData = hit ? found : new RawCell(reference.Row, column, RawKind.Blank, null, null);
        var needed = new HashSet<int>();
        if (cellData.Kind == RawKind.SharedString &&
            int.TryParse(cellData.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
        {
            needed.Add(index);
        }

        var strings = ResolveSharedStrings(document, needed);
        string Lookup(int i) => strings.GetValueOrDefault(i, string.Empty);

        var value = WireValue(cellData, Lookup);
        return new
        {
            path = $"/{ExcelPaths.QuoteSheet(entry.Name)}/{reference}",
            kind = "cell",
            sheet = entry.Name,
            address = reference.ToString(),
            value,
            type = WireType(cellData),
            formula = cellData.Formula is { } f ? "=" + f : null,
            cachedValue = cellData.Formula is null ? null : value,
            text = WireText(cellData, Lookup),
            streamed = true,
        };
    }

    public sealed record RangeResult(object Data, int TotalRows, int EmittedRows);

    public static RangeResult GetRange(string file, string sheetName, CellRef start, CellRef end, int maxCells)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var entry = SheetOrThrow(Sheets(document), sheetName);

        var firstColumn = start.ColumnNumber;
        var lastColumn = end.ColumnNumber;
        var columns = lastColumn - firstColumn + 1;
        var totalRows = end.Row - start.Row + 1;
        var maxRows = Math.Max(1, maxCells / Math.Max(1, columns));
        var emitted = Math.Min(totalRows, maxRows);
        var lastRow = start.Row + emitted - 1;

        var grid = new List<List<object?>>(emitted);
        for (var r = 0; r < emitted; r++)
        {
            grid.Add([.. Enumerable.Repeat<object?>(null, columns)]);
        }

        var window = new List<RawCell>();
        var needed = new HashSet<int>();
        Scan(PartOf(document, entry), start.Row, lastRow, firstColumn, lastColumn, cell =>
        {
            window.Add(cell);
            if (cell.Kind == RawKind.SharedString &&
                int.TryParse(cell.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
            {
                needed.Add(index);
            }

            return true;
        });

        var strings = ResolveSharedStrings(document, needed);
        string Lookup(int i) => strings.GetValueOrDefault(i, string.Empty);
        foreach (var cell in window)
        {
            grid[cell.Row - start.Row][cell.Column - firstColumn] = WireValue(cell, Lookup);
        }

        var rangeText = start.Row == end.Row && firstColumn == lastColumn
            ? start.ToString()
            : $"{start}:{end}";
        return new RangeResult(
            new
            {
                path = $"/{ExcelPaths.QuoteSheet(entry.Name)}/{rangeText}",
                kind = "range",
                sheet = entry.Name,
                range = rangeText,
                rows = totalRows,
                columns,
                values = grid,
                truncated = totalRows > emitted,
                streamed = true,
            },
            totalRows,
            emitted);
    }

    // ----- read --view text -----------------------------------------------------

    public static (string Content, bool Truncated) ReadText(string file, string? rangeArg, int maxBytes)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var sheets = Sheets(document);
        if (sheets.Count == 0)
        {
            return (string.Empty, false);
        }

        // Shared strings load lazily: number-only workbooks never touch the table.
        List<string>? table = null;
        string Lookup(int i)
        {
            table ??= AllSharedStrings(document);
            return i >= 0 && i < table.Count ? table[i] : string.Empty;
        }

        var sections = ResolveSections(sheets, rangeArg);
        var sb = new StringBuilder();
        foreach (var section in sections)
        {
            if (sb.Length > 0)
            {
                sb.Append('\n');
            }

            AppendSection(document, section, sb, maxBytes, Lookup);
            if (sb.Length > maxBytes)
            {
                return (sb.ToString(0, maxBytes), true);
            }
        }

        return (sb.ToString(), false);
    }

    private sealed record TextSection(
        SheetEntry Sheet, int FirstRow, int LastRow, int FirstColumn, int LastColumn, string? WindowText);

    private static List<TextSection> ResolveSections(List<SheetEntry> sheets, string? rangeArg)
    {
        if (rangeArg is null)
        {
            return [.. sheets.Select(WholeSheet)];
        }

        if (!rangeArg.StartsWith('/'))
        {
            // A bare range targets the first sheet, mirroring the DOM path.
            rangeArg = "/" + ExcelPaths.QuoteSheet(sheets[0].Name) + "/" + rangeArg.ToUpperInvariant();
        }

        var path = DocPath.Parse(rangeArg);
        var first = path.Segments[0];
        var sheetName = first.Kind switch
        {
            PathSegmentKind.Name => first.Name,
            PathSegmentKind.Element when first.Index is null && first.Id is null => first.Name,
            _ => null,
        };
        if (sheetName is null || path.Segments.Count > 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"'{rangeArg}' does not address a sheet window.",
                "Use a range like A1:C10, a path like /Sheet1/A1:C10, or a sheet path like /Sheet1.");
        }

        var sheet = SheetOrThrow(sheets, sheetName);
        if (path.Segments.Count == 1)
        {
            return [WholeSheet(sheet)];
        }

        var second = path.Segments[1];
        return second switch
        {
            { Kind: PathSegmentKind.Cell } => [Window(sheet, second.Start!.Value, second.Start!.Value)],
            { Kind: PathSegmentKind.Range } => [Window(sheet, second.Start!.Value, second.End!.Value)],
            { Kind: PathSegmentKind.Element, Index: { } row } when
                string.Equals(second.Name, "row", StringComparison.OrdinalIgnoreCase) =>
                [new TextSection(sheet, row, row, 1, int.MaxValue, null)],
            _ => throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"'{rangeArg}' does not address a cell, range or row the text view can stream.",
                "Use a range like /Sheet1/A1:C10 or a row like /Sheet1/row[3]."),
        };

        static TextSection WholeSheet(SheetEntry sheet) => new(sheet, 1, int.MaxValue, 1, int.MaxValue, null);

        static TextSection Window(SheetEntry sheet, CellRef start, CellRef end) => new(
            sheet, start.Row, end.Row, start.ColumnNumber, end.ColumnNumber,
            start.Row == end.Row && start.ColumnNumber == end.ColumnNumber ? start.ToString() : $"{start}:{end}");
    }

    private static void AppendSection(
        SpreadsheetDocument document, TextSection section, StringBuilder sb, int maxBytes, Func<int, string> lookup)
    {
        var body = new StringBuilder();
        var currentRow = 0;
        var any = false;
        string? dimension = Scan(
            PartOf(document, section.Sheet),
            section.FirstRow, section.LastRow, section.FirstColumn, section.LastColumn,
            cell =>
            {
                if (!cell.HasContent)
                {
                    return true;
                }

                if (cell.Row != currentRow)
                {
                    if (any)
                    {
                        body.Append('\n');
                    }

                    currentRow = cell.Row;
                }
                else
                {
                    body.Append(',');
                }

                any = true;
                body.Append(ExcelValues.CsvEscape(WireText(cell, lookup)));
                return sb.Length + body.Length <= maxBytes; // early stop once over budget
            });

        if (!any)
        {
            sb.Append("# ").Append(section.Sheet.Name).Append(" (empty)\n");
            return;
        }

        var header = section.WindowText ?? dimension;
        sb.Append("# ").Append(section.Sheet.Name);
        if (header is not null)
        {
            sb.Append('!').Append(header);
        }

        sb.Append('\n').Append(body).Append('\n');
    }

    // ----- small helpers --------------------------------------------------------

    private static string RangeText(int firstColumn, int firstRow, int lastColumn, int lastRow) =>
        string.Create(
            CultureInfo.InvariantCulture,
            $"{ExcelCharts.ColumnLetters(firstColumn)}{firstRow}:{ExcelCharts.ColumnLetters(lastColumn)}{lastRow}");
}
