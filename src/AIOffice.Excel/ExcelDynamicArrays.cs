using System.Globalization;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// (1.4) Dynamic-array formula evaluation + spill. FILTER / UNIQUE / SORT /
/// SORTBY / SEQUENCE / RANDARRAY / TRANSPOSE were recognized-but-not-evaluated
/// since M0; this module evaluates them at write time, computes the rectangular
/// result ARRAY, and SPILLS it into the range anchored at the formula cell.
///
/// <para>The anchor cell carries the dynamic-array formula in the broadly
/// compatible CSE-array form (<c>&lt;f t="array" ref="A1:A3" aca="1" ca="1"&gt;</c>)
/// so Excel re-spills on open; every cell across the spill rectangle (the anchor
/// included) carries the computed CACHED value, so headless agents and viewers
/// see the result without Excel. The spill range plus its cached values are
/// authored raw in a post-save pass (ClosedXML 0.105 has no spill model), exactly
/// like charts/slicers: the ClosedXML model leaves these cells untouched.</para>
///
/// <para>If the spill rectangle would overwrite a cell that already holds content,
/// nothing is written and the edit fails with <c>spill_blocked</c> (additive 1.4
/// code) — the suggestion names clearing the target range.</para>
/// </summary>
internal static class ExcelDynamicArrays
{
    /// <summary>The dynamic-array functions this module evaluates and spills.</summary>
    public static readonly IReadOnlyList<string> Functions =
        ["FILTER", "UNIQUE", "SORT", "SORTBY", "SEQUENCE", "RANDARRAY", "TRANSPOSE"];

    /// <summary>A computed dynamic-array spill queued for the post-save raw write.</summary>
    public sealed record Pending(
        string Sheet,
        string AnchorAddress,
        string Formula,
        string SpillRef,
        int FirstRow,
        int FirstColumn,
        IReadOnlyList<IReadOnlyList<XLCellValue>> Grid);

    /// <summary>
    /// True when the (already <c>=</c>-stripped) formula text is one of the
    /// dynamic-array functions, ignoring leading whitespace and case.
    /// </summary>
    public static bool IsDynamicArrayFormula(string formula)
    {
        var name = LeadingFunctionName(formula);
        return name is not null && Functions.Contains(name, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Evaluates a dynamic-array formula anchored at <paramref name="anchor"/>,
    /// checks the spill rectangle is clear, and returns the queued spill. The
    /// <paramref name="occupied"/> predicate reports cells already taken by this
    /// batch's earlier writes (the ClosedXML model is consulted for on-disk cells).
    /// </summary>
    /// <exception cref="AiofficeException">
    /// <c>spill_blocked</c> when the rectangle is occupied; <c>invalid_args</c>
    /// when the formula's arguments cannot be parsed.
    /// </exception>
    public static Pending Evaluate(
        IXLWorksheet sheet, IXLCell anchor, string formula, int opIndex, Func<int, int, bool>? occupied = null)
    {
        var grid = ComputeGrid(sheet, formula, opIndex);
        var rows = grid.Count;
        var columns = rows == 0 ? 0 : grid[0].Count;
        if (rows == 0 || columns == 0)
        {
            // An empty result (e.g. FILTER with no matches) still occupies the
            // single anchor cell as #CALC! in Excel; we spill a 1x1 #CALC! error.
            grid = [[XLError.NoValueAvailable]];
            rows = 1;
            columns = 1;
        }

        var firstRow = anchor.Address.RowNumber;
        var firstColumn = anchor.Address.ColumnNumber;
        var lastRow = firstRow + rows - 1;
        var lastColumn = firstColumn + columns - 1;

        if (lastRow > 1_048_576 || lastColumn > 16_384)
        {
            throw new AiofficeException(
                ErrorCodes.SpillBlocked,
                $"ops[{opIndex}]: the {rows}x{columns} spill from {anchor.Address} runs past the sheet edge.",
                "Move the anchor up/left so the whole result fits on the sheet.");
        }

        // The anchor itself may carry the formula; every OTHER cell in the
        // rectangle must be empty (in the ClosedXML model AND in this batch).
        for (var r = firstRow; r <= lastRow; r++)
        {
            for (var c = firstColumn; c <= lastColumn; c++)
            {
                if (r == firstRow && c == firstColumn)
                {
                    continue;
                }

                var cell = sheet.Cell(r, c);
                var taken = !cell.IsEmpty(XLCellsUsedOptions.AllContents) || (occupied?.Invoke(r, c) ?? false);
                if (taken)
                {
                    throw new AiofficeException(
                        ErrorCodes.SpillBlocked,
                        $"ops[{opIndex}]: the dynamic array at {anchor.Address} would spill a {rows}x{columns} " +
                        $"result onto {sheet.Cell(r, c).Address}, which is not empty.",
                        "Clear the spill range first " +
                        $"({{op:remove, path:/{ExcelPaths.QuoteSheet(sheet.Name)}/" +
                        $"{anchor.Address.ColumnLetter}{anchor.Address.RowNumber}:" +
                        $"{sheet.Cell(lastRow, lastColumn).Address.ColumnLetter}{lastRow}}}), then write the formula.");
                }
            }
        }

        var spillRef = rows == 1 && columns == 1
            ? anchor.Address.ToString()!
            : string.Create(
                CultureInfo.InvariantCulture,
                $"{anchor.Address}:{sheet.Cell(lastRow, lastColumn).Address}");

        return new Pending(
            sheet.Name,
            anchor.Address.ToString()!,
            formula.StartsWith('=') ? formula[1..] : formula,
            spillRef,
            firstRow,
            firstColumn,
            grid);
    }

    // ----- evaluators --------------------------------------------------------

    private static List<List<XLCellValue>> ComputeGrid(IXLWorksheet sheet, string formula, int opIndex)
    {
        var body = formula.StartsWith('=') ? formula[1..] : formula;
        var name = LeadingFunctionName(body)!.ToUpperInvariant();
        var args = SplitArgs(InsideParens(body, name, opIndex));

        return name switch
        {
            "UNIQUE" => Unique(sheet, args, opIndex),
            "SORT" => Sort(sheet, args, opIndex),
            "SORTBY" => SortBy(sheet, args, opIndex),
            "FILTER" => Filter(sheet, args, opIndex),
            "SEQUENCE" => Sequence(args, opIndex),
            "RANDARRAY" => RandArray(args, opIndex),
            "TRANSPOSE" => Transpose(sheet, args, opIndex),
            _ => throw Unsupported(name, opIndex),
        };
    }

    /// <summary>=UNIQUE(range) — distinct rows/values, first-seen order.</summary>
    private static List<List<XLCellValue>> Unique(IXLWorksheet sheet, List<string> args, int opIndex)
    {
        RequireArgCount("UNIQUE", args, 1, 3, opIndex);
        var (data, oneColumn) = ReadRange(sheet, args[0], opIndex);
        // by_col (arg 1) and exactly_once (arg 2) are accepted; only the common
        // row-wise, all-distinct form is evaluated (the must-have). Others fall
        // back to first-seen distinct rows, which matches the default behavior.
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var result = new List<List<XLCellValue>>();
        foreach (var row in data)
        {
            var key = string.Join('\u0001', row.Select(KeyOf));
            if (seen.Add(key))
            {
                result.Add([.. row]);
            }
        }

        _ = oneColumn;
        return result;
    }

    /// <summary>=SORT(range, [sort_index], [sort_order]) — ascending by default.</summary>
    private static List<List<XLCellValue>> Sort(IXLWorksheet sheet, List<string> args, int opIndex)
    {
        RequireArgCount("SORT", args, 1, 4, opIndex);
        var (data, _) = ReadRange(sheet, args[0], opIndex);
        var sortIndex = args.Count >= 2 && args[1].Length > 0 ? (int)ConstNumber(args[1], "SORT", opIndex) : 1;
        var ascending = !(args.Count >= 3 && args[2].Length > 0 && ConstNumber(args[2], "SORT", opIndex) < 0);
        if (sortIndex < 1 || (data.Count > 0 && sortIndex > data[0].Count))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: SORT sort_index {sortIndex} is out of the range's column count.",
                "Use a 1-based column index within the sorted range, e.g. =SORT(A1:C9, 2).");
        }

        var ordered = StableOrder(data, r => r[sortIndex - 1], ascending);
        return [.. ordered.Select(r => r.ToList())];
    }

    /// <summary>=SORTBY(range, by_range, [order]) — sort range by a parallel key column.</summary>
    private static List<List<XLCellValue>> SortBy(IXLWorksheet sheet, List<string> args, int opIndex)
    {
        RequireArgCount("SORTBY", args, 2, 6, opIndex);
        var (data, _) = ReadRange(sheet, args[0], opIndex);
        var (by, _) = ReadRange(sheet, args[1], opIndex);
        if (by.Count != data.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: SORTBY by_array has {by.Count} rows but the array has {data.Count}.",
                "The by_array must be the same height as the sorted array.");
        }

        var ascending = !(args.Count >= 3 && args[2].Length > 0 && ConstNumber(args[2], "SORTBY", opIndex) < 0);
        var keyed = data.Select((row, i) => (row, key: by[i].Count > 0 ? by[i][0] : (XLCellValue)Blank.Value)).ToList();
        var ordered = StableOrder(keyed, x => x.key, ascending);
        return [.. ordered.Select(x => x.row.ToList())];
    }

    /// <summary>=FILTER(range, include[, if_empty]) — keep rows whose include flag is truthy.</summary>
    private static List<List<XLCellValue>> Filter(IXLWorksheet sheet, List<string> args, int opIndex)
    {
        RequireArgCount("FILTER", args, 2, 3, opIndex);
        var (data, _) = ReadRange(sheet, args[0], opIndex);
        var include = ReadConditionColumn(sheet, args[1], data.Count, opIndex);
        var result = new List<List<XLCellValue>>();
        for (var i = 0; i < data.Count; i++)
        {
            if (include[i])
            {
                result.Add([.. data[i]]);
            }
        }

        return result; // empty result → caller spills a 1x1 #CALC!
    }

    /// <summary>=SEQUENCE(rows, [columns], [start], [step]) — deterministic counter grid.</summary>
    private static List<List<XLCellValue>> Sequence(List<string> args, int opIndex)
    {
        RequireArgCount("SEQUENCE", args, 1, 4, opIndex);
        var rows = (int)ConstNumber(args[0], "SEQUENCE", opIndex);
        var columns = args.Count >= 2 && args[1].Length > 0 ? (int)ConstNumber(args[1], "SEQUENCE", opIndex) : 1;
        var start = args.Count >= 3 && args[2].Length > 0 ? ConstNumber(args[2], "SEQUENCE", opIndex) : 1d;
        var step = args.Count >= 4 && args[3].Length > 0 ? ConstNumber(args[3], "SEQUENCE", opIndex) : 1d;
        if (rows < 1 || columns < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: SEQUENCE needs positive rows and columns.",
                "Use =SEQUENCE(3,2) for a 3-row, 2-column counter starting at 1.");
        }

        var grid = new List<List<XLCellValue>>(rows);
        var value = start;
        for (var r = 0; r < rows; r++)
        {
            var row = new List<XLCellValue>(columns);
            for (var c = 0; c < columns; c++)
            {
                row.Add(value);
                value += step;
            }

            grid.Add(row);
        }

        return grid;
    }

    /// <summary>
    /// =RANDARRAY([rows],[cols],[min],[max],[integer]) — NON-DETERMINISTIC.
    /// Seeded from the parsed dimensions so a given formula round-trips to the
    /// SAME cached values (tests stay platform-independent); Excel reseeds on
    /// open. The non-determinism vs. Excel is documented.
    /// </summary>
    private static List<List<XLCellValue>> RandArray(List<string> args, int opIndex)
    {
        var rows = args.Count >= 1 && args[0].Length > 0 ? (int)ConstNumber(args[0], "RANDARRAY", opIndex) : 1;
        var columns = args.Count >= 2 && args[1].Length > 0 ? (int)ConstNumber(args[1], "RANDARRAY", opIndex) : 1;
        var min = args.Count >= 3 && args[2].Length > 0 ? ConstNumber(args[2], "RANDARRAY", opIndex) : 0d;
        var max = args.Count >= 4 && args[3].Length > 0 ? ConstNumber(args[3], "RANDARRAY", opIndex) : 1d;
        var integer = args.Count >= 5 && args[4].Length > 0 && ConstBool(args[4]);
        if (rows < 1 || columns < 1 || max < min)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: RANDARRAY needs positive rows/columns and max >= min.",
                "Use =RANDARRAY(2,2,1,10,TRUE) for a 2x2 grid of integers in [1,10].");
        }

        // Deterministic seed: same dimensions/bounds → same cached values, so
        // round-trip and reopen-verify assertions are stable across platforms.
        var seed = HashCode.Combine(rows, columns, min, max, integer);
        var rng = new Random(seed);
        var grid = new List<List<XLCellValue>>(rows);
        for (var r = 0; r < rows; r++)
        {
            var row = new List<XLCellValue>(columns);
            for (var c = 0; c < columns; c++)
            {
                var v = min + (rng.NextDouble() * (max - min));
                row.Add(integer ? Math.Floor(v) : v);
            }

            grid.Add(row);
        }

        return grid;
    }

    /// <summary>=TRANSPOSE(range) — swap rows and columns.</summary>
    private static List<List<XLCellValue>> Transpose(IXLWorksheet sheet, List<string> args, int opIndex)
    {
        RequireArgCount("TRANSPOSE", args, 1, 1, opIndex);
        var (data, _) = ReadRange(sheet, args[0], opIndex);
        if (data.Count == 0)
        {
            return [];
        }

        var rows = data.Count;
        var columns = data[0].Count;
        var grid = new List<List<XLCellValue>>(columns);
        for (var c = 0; c < columns; c++)
        {
            var row = new List<XLCellValue>(rows);
            for (var r = 0; r < rows; r++)
            {
                row.Add(data[r][c]);
            }

            grid.Add(row);
        }

        return grid;
    }

    // ----- range + arg reading ----------------------------------------------

    /// <summary>Reads a range reference's cached/typed values as a row-major grid.</summary>
    private static (List<List<XLCellValue>> Data, bool OneColumn) ReadRange(
        IXLWorksheet sheet, string reference, int opIndex)
    {
        var range = ResolveRange(sheet, reference, opIndex);
        var address = range.RangeAddress;
        var rows = address.LastAddress.RowNumber - address.FirstAddress.RowNumber + 1;
        var columns = address.LastAddress.ColumnNumber - address.FirstAddress.ColumnNumber + 1;
        var data = new List<List<XLCellValue>>(rows);
        for (var r = 0; r < rows; r++)
        {
            var row = new List<XLCellValue>(columns);
            for (var c = 0; c < columns; c++)
            {
                row.Add(ValueOf(range.Cell(r + 1, c + 1)));
            }

            data.Add(row);
        }

        return (data, columns == 1);
    }

    /// <summary>Reads the include condition for FILTER as a per-row bool list.</summary>
    private static List<bool> ReadConditionColumn(IXLWorksheet sheet, string reference, int rows, int opIndex)
    {
        var range = ResolveRange(sheet, reference, opIndex);
        var address = range.RangeAddress;
        var height = address.LastAddress.RowNumber - address.FirstAddress.RowNumber + 1;
        if (height != rows)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: FILTER include is {height} rows but the array is {rows}.",
                "The include array must be the same height as the filtered array (one flag per row).");
        }

        var flags = new List<bool>(rows);
        for (var r = 1; r <= rows; r++)
        {
            flags.Add(IsTruthy(ValueOf(range.Cell(r, 1))));
        }

        return flags;
    }

    private static IXLRange ResolveRange(IXLWorksheet sheet, string reference, int opIndex)
    {
        var trimmed = reference.Trim();
        var bang = trimmed.LastIndexOf('!');
        IXLWorksheet targetSheet = sheet;
        if (bang >= 0)
        {
            var sheetRef = trimmed[..bang].Trim('\'', ' ');
            trimmed = trimmed[(bang + 1)..];
            if (!sheet.Workbook.TryGetWorksheet(sheetRef, out targetSheet!))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: the dynamic-array formula references sheet '{sheetRef}', which does not exist.",
                    "Reference a sheet that exists, or drop the sheet prefix to read the current sheet.");
            }
        }

        trimmed = trimmed.Replace("$", string.Empty, StringComparison.Ordinal);
        try
        {
            return trimmed.Contains(':', StringComparison.Ordinal)
                ? targetSheet.Range(trimmed)
                : targetSheet.Range(trimmed + ":" + trimmed);
        }
        catch (Exception)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{reference}' is not a range the dynamic-array evaluator can read.",
                "Pass a cell-range reference like A1:A9 (optionally Sheet!A1:A9); whole-column refs are not supported.");
        }
    }

    /// <summary>The cell's evaluated value, preferring a clean cached value over a stale formula error.</summary>
    private static XLCellValue ValueOf(IXLCell cell)
    {
        if (!cell.HasFormula)
        {
            return cell.Value;
        }

        try
        {
            var v = cell.Value;
            return v.IsError ? cell.CachedValue : v;
        }
        catch (Exception)
        {
            return cell.CachedValue;
        }
    }

    // ----- formula text helpers ---------------------------------------------

    /// <summary>The leading function name (letters up to the first <c>(</c>), or null.</summary>
    private static string? LeadingFunctionName(string formula)
    {
        var text = formula.StartsWith('=') ? formula[1..] : formula;
        text = text.TrimStart();
        var i = 0;
        while (i < text.Length && (char.IsLetter(text[i]) || text[i] == '_'))
        {
            i++;
        }

        if (i == 0 || i >= text.Length)
        {
            return null;
        }

        return text[i] == '(' ? text[..i] : null;
    }

    private static string InsideParens(string body, string name, int opIndex)
    {
        var text = body.TrimStart();
        var open = text.IndexOf('(', StringComparison.Ordinal);
        if (open < 0 || !text.EndsWith(')'))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: {name} is missing its argument parentheses.",
                $"Write the function like ={name}(A1:A9).");
        }

        return text[(open + 1)..^1];
    }

    /// <summary>Splits an argument list on top-level commas (respecting nested parens and quotes).</summary>
    private static List<string> SplitArgs(string inside)
    {
        var args = new List<string>();
        if (inside.Trim().Length == 0)
        {
            return args;
        }

        var depth = 0;
        var inQuote = false;
        var start = 0;
        for (var i = 0; i < inside.Length; i++)
        {
            var ch = inside[i];
            if (ch == '"')
            {
                inQuote = !inQuote;
            }
            else if (!inQuote && ch is '(' or '{')
            {
                depth++;
            }
            else if (!inQuote && ch is ')' or '}')
            {
                depth--;
            }
            else if (!inQuote && depth == 0 && ch == ',')
            {
                args.Add(inside[start..i].Trim());
                start = i + 1;
            }
        }

        args.Add(inside[start..].Trim());
        return args;
    }

    private static double ConstNumber(string text, string fn, int opIndex)
    {
        var t = text.Trim();
        if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
        {
            return n;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: {fn} expected a numeric constant but got '{text}'.",
            $"Pass plain numbers to {fn} (cell-reference arguments are not evaluated for this argument).");
    }

    private static bool ConstBool(string text) =>
        text.Trim().Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
        (double.TryParse(text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var n) && n != 0);

    private static void RequireArgCount(string fn, List<string> args, int min, int max, int opIndex)
    {
        if (args.Count < min || args.Count > max)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: {fn} takes {min}-{max} argument(s) but got {args.Count}.",
                $"Check the {fn} signature; e.g. =UNIQUE(A1:A9) or =SORT(A1:C9, 2, -1).");
        }
    }

    private static AiofficeException Unsupported(string name, int opIndex) => new(
        ErrorCodes.UnsupportedFeature,
        $"ops[{opIndex}]: the dynamic-array function {name} is recognized but not evaluated.",
        "Supported dynamic arrays: " + string.Join(", ", Functions) + ".");

    private static string KeyOf(XLCellValue value) => value.Type switch
    {
        XLDataType.Number => value.GetNumber().ToString("R", CultureInfo.InvariantCulture),
        XLDataType.Boolean => value.GetBoolean() ? "b:1" : "b:0",
        XLDataType.DateTime => "d:" + value.GetDateTime().ToString("O", CultureInfo.InvariantCulture),
        XLDataType.Blank => "_:blank",
        XLDataType.Error => "e:" + value,
        _ => "t:" + value.GetText(),
    };

    private static bool IsTruthy(XLCellValue value) => value.Type switch
    {
        XLDataType.Boolean => value.GetBoolean(),
        XLDataType.Number => value.GetNumber() != 0,
        XLDataType.Text => value.GetText().Equals("TRUE", StringComparison.OrdinalIgnoreCase),
        _ => false,
    };

    /// <summary>Stable ascending/descending order over typed cell values.</summary>
    private static List<T> StableOrder<T>(IEnumerable<T> source, Func<T, XLCellValue> key, bool ascending)
    {
        var indexed = source.Select((item, i) => (item, i)).ToList();
        indexed.Sort((a, b) =>
        {
            var cmp = CompareValues(key(a.item), key(b.item));
            if (cmp == 0)
            {
                return a.i.CompareTo(b.i); // stable
            }

            return ascending ? cmp : -cmp;
        });

        return [.. indexed.Select(x => x.item)];
    }

    private static int CompareValues(XLCellValue a, XLCellValue b)
    {
        // Excel sort rank: numbers/dates < text < boolean; blanks sort last.
        var ra = Rank(a);
        var rb = Rank(b);
        if (ra != rb)
        {
            return ra.CompareTo(rb);
        }

        return a.Type switch
        {
            XLDataType.Number => a.GetNumber().CompareTo(b.GetNumber()),
            XLDataType.DateTime => a.GetDateTime().CompareTo(b.GetDateTime()),
            XLDataType.Boolean => a.GetBoolean().CompareTo(b.GetBoolean()),
            XLDataType.Text => string.Compare(a.GetText(), b.GetText(), StringComparison.OrdinalIgnoreCase),
            _ => 0,
        };
    }

    private static int Rank(XLCellValue v) => v.Type switch
    {
        XLDataType.Number or XLDataType.DateTime or XLDataType.TimeSpan => 0,
        XLDataType.Text => 1,
        XLDataType.Boolean => 2,
        XLDataType.Error => 3,
        _ => 4, // blank last
    };

    // ----- post-save raw write ----------------------------------------------

    /// <summary>
    /// Authors the queued spills raw on the saved bytes: the anchor's CSE-array
    /// formula (<c>&lt;f t="array" ref="…"&gt;</c>) plus cached values across the
    /// whole spill rectangle. ClosedXML left these cells untouched, so this is the
    /// single writer of spill metadata.
    /// </summary>
    public static void ApplyAfterSave(string file, IReadOnlyList<Pending> spills)
    {
        if (spills.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart!;
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;

        foreach (var group in spills.GroupBy(s => s.Sheet, StringComparer.OrdinalIgnoreCase))
        {
            if (ExcelFormulaParts.SheetDataFor(workbookPart, group.Key) is not { } sheetData)
            {
                continue;
            }

            foreach (var spill in group)
            {
                WriteSpill(sheetData, spill, ref sharedStrings, workbookPart);
            }

            sheetData.Ancestors<S.Worksheet>().First().Save();
        }

        // Spilled formulas reference live data; force a full recalc on open so
        // Excel re-spills authoritatively while our cached values serve headless.
        ExcelFormulaParts.SetFullRecalc(workbookPart);
        workbookPart.Workbook!.Save();
    }

    private static void WriteSpill(
        S.SheetData sheetData, Pending spill, ref S.SharedStringTable? sharedStrings, WorkbookPart workbookPart)
    {
        for (var r = 0; r < spill.Grid.Count; r++)
        {
            var rowNumber = (uint)(spill.FirstRow + r);
            var row = ExcelFormulaParts.EnsureRow(sheetData, rowNumber);
            for (var c = 0; c < spill.Grid[r].Count; c++)
            {
                var columnNumber = spill.FirstColumn + c;
                var reference = ExcelFormulaParts.CellRef(columnNumber, (int)rowNumber);
                var cell = ExcelFormulaParts.EnsureCell(row, reference, columnNumber);
                var isAnchor = r == 0 && c == 0;
                SetCellValue(cell, spill.Grid[r][c], ref sharedStrings, workbookPart);
                if (isAnchor)
                {
                    cell.CellFormula = new S.CellFormula(spill.Formula)
                    {
                        FormulaType = S.CellFormulaValues.Array,
                        Reference = spill.SpillRef,
                        AlwaysCalculateArray = true,
                        CalculateCell = true,
                    };
                }
                else
                {
                    cell.CellFormula = null; // spill cells carry the cached value only
                }
            }
        }
    }

    private static void SetCellValue(
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
}
