using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Word table formulas (v1.5.0). A cell can carry a <c>formula</c> prop —
/// <c>=SUM(ABOVE)</c>, <c>=AVERAGE(LEFT)</c>, <c>=PRODUCT(ABOVE)</c> or a
/// cell-reference arithmetic form (<c>=A1*B2</c>, <c>=(A1+A2)/2</c>) using the
/// table's A1 addressing (column letters across, 1-based rows down). It becomes a
/// Word table-formula field (<c>w:fldSimple</c> whose instruction is the
/// <c>=</c> expression) with the value computed headlessly and cached as the
/// field result, so <c>read --view text</c>, <c>get</c> and Word (before F9) all
/// show the same number. Static cell values compute exactly; when an input cell
/// is itself a field a <c>table_formula_cached</c> note flags that Word may
/// refresh to a different number.
/// </summary>
public sealed partial class WordHandler
{
    /// <summary>The directional aggregate functions Word's table formulas understand.</summary>
    private static readonly string[] TableFormulaFunctions = ["SUM", "AVERAGE", "PRODUCT", "COUNT", "MIN", "MAX"];

    [GeneratedRegex(@"^([A-Z]{1,3})([1-9][0-9]*)$")]
    private static partial Regex TableCellRef();

    [GeneratedRegex(@"^=\s*([A-Za-z]+)\s*\(\s*(ABOVE|BELOW|LEFT|RIGHT)\s*\)\s*$", RegexOptions.IgnoreCase)]
    private static partial Regex DirectionalFormula();

    /// <summary>The numeric format presets a table formula may request (Word's "\# " picture heads).</summary>
    private static readonly Dictionary<string, string> NumberFormatPictures = new(StringComparer.OrdinalIgnoreCase)
    {
        ["integer"] = "0",
        ["number"] = "#,##0.00",
        ["percent"] = "0%",
        ["currency"] = "$#,##0.00",
    };

    /// <summary>
    /// Applies a <c>formula</c> prop to a table cell: validates the expression,
    /// computes its cached numeric value against the table, and replaces the
    /// cell's content with a single <c>w:fldSimple</c> formula field. An optional
    /// <c>numberFormat</c> prop (a preset name or a raw Word "\#" picture) shapes
    /// both the cached text and the field's stored picture. Returns the warning
    /// state (a cached note when an input is a field) for the caller to surface.
    /// </summary>
    private static (string Cached, bool InputsAreFields) ApplyCellFormula(
        TableCell cell, ResolvedNode node, string formula, string? numberFormat)
    {
        if (cell.Parent is not TableRow row || row.Parent is not Table table)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                "The cell is not inside a table row, so a table formula has no grid to compute over.",
                "Re-export the file from Word.");
        }

        var expression = formula.Trim();
        if (!expression.StartsWith('='))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A table formula must start with '=', got '{formula}'.",
                "Use =SUM(ABOVE), =AVERAGE(LEFT), =PRODUCT(ABOVE) or a cell-ref form like =A1*B2.");
        }

        var picture = ResolveNumberFormatPicture(numberFormat);
        var (value, inputsAreFields) = EvaluateTableFormula(table, cell, expression);
        var cached = FormatFormulaResult(value, picture);

        // Replace the cell's content with one paragraph holding the formula field.
        cell.RemoveAllChildren<Paragraph>();
        var instruction = picture is null
            ? $" {expression.TrimStart('=').Trim().Insert(0, "=")} "
            : $" ={expression.TrimStart('=').Trim()} \\# \"{picture}\" ";
        var field = new SimpleField(new Run(NewText(cached))) { Instruction = instruction };
        cell.AppendChild(new Paragraph(field));

        return (cached, inputsAreFields);
    }

    /// <summary>A preset name maps to its Word picture; a raw "\#"-style picture passes through (quote-safe).</summary>
    private static string? ResolveNumberFormatPicture(string? numberFormat)
    {
        if (numberFormat is not { Length: > 0 })
        {
            return null;
        }

        if (NumberFormatPictures.TryGetValue(numberFormat, out var preset))
        {
            return preset;
        }

        if (numberFormat.Contains('"', StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"numberFormat must not contain double quotes, got '{numberFormat}'.",
                "Use a preset (integer, number, percent, currency) or a Word picture like \"0.00\" or \"#,##0\".");
        }

        return numberFormat;
    }

    /// <summary>Formats a computed value with the picture; absent picture trims trailing zeros to a tidy decimal.</summary>
    private static string FormatFormulaResult(double value, string? picture)
    {
        if (picture is null)
        {
            // Whole numbers show without a decimal point; otherwise a compact decimal.
            return value == Math.Truncate(value) && Math.Abs(value) < 1e15
                ? ((long)value).ToString(CultureInfo.InvariantCulture)
                : value.ToString("0.######", CultureInfo.InvariantCulture);
        }

        if (picture == "0%")
        {
            return (value * 100).ToString("0", CultureInfo.InvariantCulture) + "%";
        }

        // The supported pictures map onto standard .NET numeric formats.
        return picture switch
        {
            "0" => Math.Round(value).ToString("0", CultureInfo.InvariantCulture),
            "#,##0.00" => value.ToString("#,##0.00", CultureInfo.InvariantCulture),
            "$#,##0.00" => value.ToString("$#,##0.00", CultureInfo.InvariantCulture),
            _ => TryFormatWithPicture(value, picture),
        };
    }

    private static string TryFormatWithPicture(double value, string picture)
    {
        // A raw user picture (e.g. "0.00", "#,##0") is a .NET-compatible custom
        // numeric format in the cases Word and .NET share; fall back to a tidy
        // decimal if it is something exotic .NET cannot render.
        try
        {
            return value.ToString(picture, CultureInfo.InvariantCulture);
        }
        catch (FormatException)
        {
            return value.ToString("0.######", CultureInfo.InvariantCulture);
        }
    }

    // ----------------------------------------------------------- evaluation

    /// <summary>
    /// Computes a table formula against the live grid. Directional forms
    /// (ABOVE/BELOW/LEFT/RIGHT) aggregate the run of numeric cells in that
    /// direction from the target; cell-ref forms evaluate the arithmetic
    /// expression with each A1 reference replaced by its cell's number. Returns
    /// the value plus whether any consumed input cell is itself a field (so the
    /// caller can warn the value is cache-only).
    /// </summary>
    private static (double Value, bool InputsAreFields) EvaluateTableFormula(Table table, TableCell target, string expression)
    {
        var directional = DirectionalFormula().Match(expression);
        if (directional.Success)
        {
            var function = directional.Groups[1].Value.ToUpperInvariant();
            if (!TableFormulaFunctions.Contains(function, StringComparer.Ordinal))
            {
                throw UnknownFunction(function);
            }

            var direction = directional.Groups[2].Value.ToUpperInvariant();
            var cells = DirectionalCells(table, target, direction);
            var numbers = new List<double>();
            var inputsAreFields = false;
            foreach (var cell in cells)
            {
                inputsAreFields |= CellHasField(cell);
                if (CellNumericValue(cell) is { } n)
                {
                    numbers.Add(n);
                }
            }

            return (Aggregate(function, numbers), inputsAreFields);
        }

        // Cell-reference arithmetic: substitute every A1 reference, then evaluate.
        return EvaluateCellRefExpression(table, expression);
    }

    private static (double Value, bool InputsAreFields) EvaluateCellRefExpression(Table table, string expression)
    {
        var body = expression.TrimStart('=').Trim();
        var grid = BuildCellGrid(table);
        var inputsAreFields = false;

        var substituted = Regex.Replace(body, @"[A-Z]{1,3}[1-9][0-9]*", match =>
        {
            var refMatch = TableCellRef().Match(match.Value);
            var column = ColumnLetterToIndex(refMatch.Groups[1].Value);
            var rowIndex = int.Parse(refMatch.Groups[2].Value, CultureInfo.InvariantCulture) - 1;

            if (rowIndex < 0 || rowIndex >= grid.Count || column < 0 || column >= grid[rowIndex].Count ||
                grid[rowIndex][column] is not { } cell)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Cell reference {match.Value} is outside the table's grid.",
                    "Table A1 references use column letters across and 1-based rows down; pick a cell that exists.");
            }

            inputsAreFields |= CellHasField(cell);
            var number = CellNumericValue(cell) ?? 0.0;
            return number.ToString("R", CultureInfo.InvariantCulture);
        });

        var value = ArithmeticEvaluator.Evaluate(substituted)
            ?? throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Could not evaluate the table formula '{expression}'.",
                "Use directional forms (=SUM(ABOVE)) or simple cell-ref arithmetic (=A1*B2, =(A1+A2)/2) with + - * / and parentheses.");

        return (value, inputsAreFields);
    }

    private static AiofficeException UnknownFunction(string function) => new(
        ErrorCodes.UnsupportedFeature,
        $"Table formula function '{function}' is not supported.",
        $"Use one of: {string.Join(", ", TableFormulaFunctions)} with ABOVE/BELOW/LEFT/RIGHT, " +
        "or a cell-ref arithmetic form like =A1*B2.",
        candidates: TableFormulaFunctions);

    private static double Aggregate(string function, List<double> numbers) => function switch
    {
        "SUM" => numbers.Sum(),
        "AVERAGE" => numbers.Count > 0 ? numbers.Average() : 0.0,
        "PRODUCT" => numbers.Count > 0 ? numbers.Aggregate(1.0, (a, b) => a * b) : 0.0,
        "COUNT" => numbers.Count,
        "MIN" => numbers.Count > 0 ? numbers.Min() : 0.0,
        "MAX" => numbers.Count > 0 ? numbers.Max() : 0.0,
        _ => 0.0,
    };

    // -------------------------------------------------------------- grid

    /// <summary>
    /// The cells in one direction from a target, nearest-first: ABOVE walks up the
    /// target's column, LEFT walks left along its row, etc. Word aggregates the
    /// contiguous run; we collect every cell up to the table edge and let the
    /// numeric filter drop the non-numeric ones (header text included), matching
    /// Word's "stop at the first blank/non-number is unusual" practical behavior
    /// of summing the column.
    /// </summary>
    private static List<TableCell> DirectionalCells(Table table, TableCell target, string direction)
    {
        var grid = BuildCellGrid(table);
        var (targetRow, targetCol) = LocateCell(grid, target);
        if (targetRow < 0)
        {
            return [];
        }

        var cells = new List<TableCell>();
        switch (direction)
        {
            case "ABOVE":
                for (var r = targetRow - 1; r >= 0; r--)
                {
                    AddIfPresent(grid, r, targetCol, cells);
                }

                break;

            case "BELOW":
                for (var r = targetRow + 1; r < grid.Count; r++)
                {
                    AddIfPresent(grid, r, targetCol, cells);
                }

                break;

            case "LEFT":
                for (var c = targetCol - 1; c >= 0; c--)
                {
                    AddIfPresent(grid, targetRow, c, cells);
                }

                break;

            default: // RIGHT
                for (var c = targetCol + 1; c < grid[targetRow].Count; c++)
                {
                    AddIfPresent(grid, targetRow, c, cells);
                }

                break;
        }

        return cells;
    }

    private static void AddIfPresent(List<List<TableCell?>> grid, int row, int col, List<TableCell> into)
    {
        if (row >= 0 && row < grid.Count && col >= 0 && col < grid[row].Count && grid[row][col] is { } cell &&
            !into.Contains(cell))
        {
            into.Add(cell);
        }
    }

    /// <summary>
    /// The table as a dense row-major grid of cells. Horizontal merges
    /// (w:gridSpan) repeat the owning cell across the columns it covers; vertical
    /// merges leave the continuation slot pointing at its real cell so column
    /// walks stay aligned.
    /// </summary>
    private static List<List<TableCell?>> BuildCellGrid(Table table)
    {
        var grid = new List<List<TableCell?>>();
        foreach (var row in table.ChildElements.OfType<TableRow>())
        {
            var line = new List<TableCell?>();
            foreach (var cell in row.ChildElements.OfType<TableCell>())
            {
                var span = CellSpan(cell);
                for (var s = 0; s < span; s++)
                {
                    line.Add(cell);
                }
            }

            grid.Add(line);
        }

        return grid;
    }

    private static (int Row, int Col) LocateCell(List<List<TableCell?>> grid, TableCell target)
    {
        for (var r = 0; r < grid.Count; r++)
        {
            for (var c = 0; c < grid[r].Count; c++)
            {
                if (ReferenceEquals(grid[r][c], target))
                {
                    return (r, c);
                }
            }
        }

        return (-1, -1);
    }

    /// <summary>Column letters to a 0-based grid index (A→0, B→1, AA→26).</summary>
    private static int ColumnLetterToIndex(string letters)
    {
        var n = 0;
        foreach (var ch in letters)
        {
            n = (n * 26) + (ch - 'A' + 1);
        }

        return n - 1;
    }

    // ------------------------------------------------------------- cell value

    /// <summary>True when a cell's content includes any field (simple or complex), so its number is a cache.</summary>
    private static bool CellHasField(TableCell cell) =>
        cell.Descendants<SimpleField>().Any() ||
        cell.Descendants<FieldChar>().Any();

    /// <summary>
    /// The numeric value a cell contributes: its text parsed as a number, with
    /// currency/percent/grouping symbols stripped (Word does the same). A percent
    /// value divides by 100. Non-numeric cells return null and are skipped.
    /// </summary>
    private static double? CellNumericValue(TableCell cell)
    {
        var text = cell.InnerText.Trim();
        if (text.Length == 0)
        {
            return null;
        }

        var isPercent = text.EndsWith('%');
        var cleaned = new string([.. text.Where(c => char.IsDigit(c) || c is '.' or '-' or '+')]);
        if (cleaned.Length == 0 ||
            !double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
        {
            return null;
        }

        return isPercent ? value / 100.0 : value;
    }

    // ------------------------------------------------------------- read

    /// <summary>The formula instruction on a cell's first formula field ("=SUM(ABOVE)"), or null.</summary>
    internal static string? CellFormula(TableCell cell)
    {
        var field = cell.Descendants<SimpleField>()
            .FirstOrDefault(f => IsFormulaInstruction(f.Instruction?.Value));
        if (field?.Instruction?.Value is not { } instruction)
        {
            return null;
        }

        // Strip a trailing "\# picture" switch and surrounding whitespace.
        var trimmed = instruction.Trim();
        var switchAt = trimmed.IndexOf("\\#", StringComparison.Ordinal);
        if (switchAt >= 0)
        {
            trimmed = trimmed[..switchAt].Trim();
        }

        return trimmed;
    }

    /// <summary>The cached result text of a cell's formula field, or null when there is none.</summary>
    internal static string? CellFormulaCachedValue(TableCell cell) =>
        cell.Descendants<SimpleField>()
            .FirstOrDefault(f => IsFormulaInstruction(f.Instruction?.Value))?.InnerText;

    /// <summary>A w:fldSimple instruction whose body is an "=" expression (a table formula).</summary>
    private static bool IsFormulaInstruction(string? instruction) =>
        instruction?.Trim().StartsWith('=') ?? false;

    /// <summary>
    /// A small recursive-descent evaluator for the cell-ref arithmetic forms a
    /// Word table formula allows after every reference has been substituted with a
    /// literal number: <c>+ - * /</c>, unary minus, parentheses and decimal
    /// literals. It is deliberately tiny — table formulas are not a general
    /// spreadsheet engine — and returns null on any malformed input so the caller
    /// can raise a precise <c>invalid_args</c>.
    /// </summary>
    private static class ArithmeticEvaluator
    {
        public static double? Evaluate(string expression)
        {
            try
            {
                var parser = new Parser(expression);
                var value = parser.ParseExpression();
                return parser.AtEnd && !double.IsNaN(value) && !double.IsInfinity(value) ? value : null;
            }
            catch (FormatException)
            {
                return null;
            }
            catch (DivideByZeroException)
            {
                return null;
            }
        }

        private sealed class Parser(string text)
        {
            private int _pos;

            public bool AtEnd
            {
                get
                {
                    SkipWhitespace();
                    return _pos >= text.Length;
                }
            }

            // expression := term (('+' | '-') term)*
            public double ParseExpression()
            {
                var value = ParseTerm();
                while (true)
                {
                    SkipWhitespace();
                    if (Peek() is '+')
                    {
                        _pos++;
                        value += ParseTerm();
                    }
                    else if (Peek() is '-')
                    {
                        _pos++;
                        value -= ParseTerm();
                    }
                    else
                    {
                        return value;
                    }
                }
            }

            // term := factor (('*' | '/') factor)*
            private double ParseTerm()
            {
                var value = ParseFactor();
                while (true)
                {
                    SkipWhitespace();
                    if (Peek() is '*')
                    {
                        _pos++;
                        value *= ParseFactor();
                    }
                    else if (Peek() is '/')
                    {
                        _pos++;
                        var divisor = ParseFactor();
                        if (divisor == 0)
                        {
                            throw new DivideByZeroException();
                        }

                        value /= divisor;
                    }
                    else
                    {
                        return value;
                    }
                }
            }

            // factor := '-' factor | '(' expression ')' | number
            private double ParseFactor()
            {
                SkipWhitespace();
                var c = Peek();
                if (c is '-')
                {
                    _pos++;
                    return -ParseFactor();
                }

                if (c is '+')
                {
                    _pos++;
                    return ParseFactor();
                }

                if (c is '(')
                {
                    _pos++;
                    var value = ParseExpression();
                    SkipWhitespace();
                    if (Peek() is not ')')
                    {
                        throw new FormatException("Expected ')'.");
                    }

                    _pos++;
                    return value;
                }

                return ParseNumber();
            }

            private double ParseNumber()
            {
                SkipWhitespace();
                var start = _pos;
                while (_pos < text.Length && (char.IsDigit(text[_pos]) || text[_pos] is '.' or 'e' or 'E' or '+' or '-'))
                {
                    // Allow exponent sign only right after e/E (so it does not eat a binary operator).
                    if (text[_pos] is '+' or '-' && _pos > start && text[_pos - 1] is not ('e' or 'E'))
                    {
                        break;
                    }

                    _pos++;
                }

                if (_pos == start ||
                    !double.TryParse(text.AsSpan(start, _pos - start), NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
                {
                    throw new FormatException($"Expected a number at position {start}.");
                }

                return value;
            }

            private char? Peek() => _pos < text.Length ? text[_pos] : null;

            private void SkipWhitespace()
            {
                while (_pos < text.Length && char.IsWhiteSpace(text[_pos]))
                {
                    _pos++;
                }
            }
        }
    }
}
