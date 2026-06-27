using System.Globalization;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// (1.5) Write-time evaluation of the scalar functions ClosedXML 0.105 leaves as
/// <c>#NAME?</c>: <c>XLOOKUP</c>, <c>IFS</c>, <c>SWITCH</c>, <c>LET</c>,
/// <c>MAXIFS</c>, <c>MINIFS</c>, <c>AVERAGEIFS</c> and <c>AVERAGEIF</c>. The cached value is
/// written into the saved cell, so headless readers see a real result instead of
/// riding a <c>formula_not_evaluated</c> warning (the same discipline as the v1.4
/// financial fallback).
///
/// <para>The function-specific control flow (branch selection, sequential name
/// binding, lookup, conditional aggregation) is handled here; the arithmetic and
/// comparison <em>inside</em> a branch / binding is delegated to ClosedXML's own
/// engine through a far-off scratch cell that is always restored, so the live
/// model is left exactly as it was. A sub-expression that itself uses an
/// unevaluable function (e.g. a nested <c>XLOOKUP</c>) honestly returns
/// <c>#NAME?</c> and the cell keeps the <c>formula_not_evaluated</c> warning.</para>
///
/// <para><c>TEXTJOIN</c>, <c>CONCAT</c>, <c>CONCATENATE</c>, <c>IFERROR</c>,
/// <c>SUMIFS</c> and <c>COUNTIFS</c> are already evaluated by ClosedXML's engine
/// (verified), so they never reach this fallback — they carry a cached value
/// straight from the normal save. <c>TEXTSPLIT</c> spills, so it lives in
/// <see cref="ExcelDynamicArrays"/>.</para>
/// </summary>
internal static class ExcelScalarFunctions
{
    /// <summary>The scalar functions this module evaluates as a save-time fallback (those ClosedXML returns #NAME? for).</summary>
    public static readonly IReadOnlyList<string> Functions =
        ["XLOOKUP", "IFS", "SWITCH", "LET", "MAXIFS", "MINIFS", "AVERAGEIFS", "AVERAGEIF"];

    // A scratch cell far outside any realistic used range; the evaluator writes a
    // sub-expression here, reads ClosedXML's result, and clears it afterwards.
    private const string ScratchAddress = "XFD1048576";

    /// <summary>True when the (=-stripped) formula's leading function is one we evaluate.</summary>
    public static bool IsScalarFunction(string formula)
    {
        var name = ExcelFormulaText.LeadingFunctionName(formula);
        return name is not null && Functions.Contains(name, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Evaluates a scalar formula against the sheet. Returns the typed result, or a
    /// <c>#NAME?</c> error value when a sub-expression itself uses an unevaluable
    /// function (the caller then keeps the honest formula_not_evaluated warning) or
    /// a <c>#N/A</c>/<c>#VALUE!</c> when the inputs are degenerate. Throws nothing.
    /// </summary>
    public static XLCellValue Evaluate(IXLWorksheet sheet, string formula)
    {
        try
        {
            var body = formula.StartsWith('=') ? formula[1..] : formula;
            var name = ExcelFormulaText.LeadingFunctionName(body)!.ToUpperInvariant();
            var args = ExcelFormulaText.SplitArgs(ExcelFormulaText.InsideParens(body));
            return name switch
            {
                "XLOOKUP" => XLookup(sheet, args),
                "IFS" => Ifs(sheet, args),
                "SWITCH" => Switch(sheet, args),
                "LET" => Let(sheet, args),
                "MAXIFS" => ConditionalAggregate(sheet, args, Aggregate.Max),
                "MINIFS" => ConditionalAggregate(sheet, args, Aggregate.Min),
                "AVERAGEIFS" => ConditionalAggregate(sheet, args, Aggregate.Average),
                "AVERAGEIF" => AverageIf(sheet, args),
                _ => XLError.NameNotRecognized,
            };
        }
        catch (NameNotRecognizedException)
        {
            // A sub-expression used a function we cannot evaluate; stay honest.
            return XLError.NameNotRecognized;
        }
        catch (Exception)
        {
            return XLError.NoValueAvailable;
        }
    }

    // ----- XLOOKUP -----------------------------------------------------------

    /// <summary>
    /// =XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found],
    /// [match_mode], [search_mode]) — exact match by default; match_mode -1/1 do
    /// next-smaller/next-larger approximate match. Returns one cell from the
    /// return array; on no match returns if_not_found (or #N/A).
    /// </summary>
    private static XLCellValue XLookup(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count < 3)
        {
            return XLError.NoValueAvailable;
        }

        var lookupValue = EvalScalar(sheet, args[0]);
        var lookupArray = ReadVector(sheet, args[1]);
        var returnArray = ReadVector(sheet, args[2]);
        if (lookupArray.Count == 0 || lookupArray.Count != returnArray.Count)
        {
            return XLError.NoValueAvailable;
        }

        var matchMode = args.Count >= 5 && args[4].Trim().Length > 0
            ? (int)ToNumber(EvalScalar(sheet, args[4]))
            : 0;

        var index = matchMode switch
        {
            -1 => ApproximateMatch(lookupValue, lookupArray, preferSmaller: true),
            1 => ApproximateMatch(lookupValue, lookupArray, preferSmaller: false),
            _ => ExactMatch(lookupValue, lookupArray),
        };

        if (index >= 0)
        {
            return returnArray[index];
        }

        // if_not_found (arg 3) when supplied, else #N/A.
        return args.Count >= 4 && args[3].Trim().Length > 0
            ? EvalScalar(sheet, args[3])
            : XLError.NoValueAvailable;
    }

    private static int ExactMatch(XLCellValue needle, IReadOnlyList<XLCellValue> haystack)
    {
        for (var i = 0; i < haystack.Count; i++)
        {
            if (ValuesEqual(needle, haystack[i]))
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>
    /// Approximate numeric match: the closest value that is not larger (preferSmaller)
    /// or not smaller than the lookup value. Falls back to exact for non-numbers.
    /// </summary>
    private static int ApproximateMatch(XLCellValue needle, IReadOnlyList<XLCellValue> haystack, bool preferSmaller)
    {
        if (needle.Type != XLDataType.Number && needle.Type != XLDataType.DateTime)
        {
            return ExactMatch(needle, haystack);
        }

        var target = ToNumber(needle);
        var best = -1;
        var bestDelta = double.MaxValue;
        for (var i = 0; i < haystack.Count; i++)
        {
            if (haystack[i].Type != XLDataType.Number && haystack[i].Type != XLDataType.DateTime)
            {
                if (ValuesEqual(needle, haystack[i]))
                {
                    return i;
                }

                continue;
            }

            var candidate = ToNumber(haystack[i]);
            if (candidate == target)
            {
                return i; // exact wins immediately
            }

            var withinDirection = preferSmaller ? candidate < target : candidate > target;
            if (!withinDirection)
            {
                continue;
            }

            var delta = Math.Abs(candidate - target);
            if (delta < bestDelta)
            {
                bestDelta = delta;
                best = i;
            }
        }

        return best;
    }

    // ----- IFS / SWITCH ------------------------------------------------------

    /// <summary>=IFS(cond1, val1, cond2, val2, …) — the first truthy condition's value.</summary>
    private static XLCellValue Ifs(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count < 2 || args.Count % 2 != 0)
        {
            return XLError.NoValueAvailable;
        }

        for (var i = 0; i + 1 < args.Count; i += 2)
        {
            if (IsTruthy(EvalScalar(sheet, args[i])))
            {
                return EvalScalar(sheet, args[i + 1]);
            }
        }

        return XLError.NoValueAvailable; // no condition matched
    }

    /// <summary>=SWITCH(expr, match1, val1, …, [default]) — the value of the first matching case.</summary>
    private static XLCellValue Switch(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count < 3)
        {
            return XLError.NoValueAvailable;
        }

        var subject = EvalScalar(sheet, args[0]);
        var i = 1;
        for (; i + 1 < args.Count; i += 2)
        {
            if (ValuesEqual(subject, EvalScalar(sheet, args[i])))
            {
                return EvalScalar(sheet, args[i + 1]);
            }
        }

        // A trailing odd argument is the default value.
        return i < args.Count ? EvalScalar(sheet, args[i]) : XLError.NoValueAvailable;
    }

    // ----- LET ---------------------------------------------------------------

    /// <summary>
    /// =LET(name1, value1, [name2, value2, …], calculation) — binds each name to
    /// its value in order (later bindings may reference earlier ones), then
    /// evaluates the final calculation. Bindings are substituted as parenthesized
    /// literals so ClosedXML evaluates the calc with no awareness of LET.
    /// </summary>
    private static XLCellValue Let(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count < 3 || args.Count % 2 == 0)
        {
            // LET needs at least one name/value pair plus the calculation, so the
            // count is odd and >= 3.
            return XLError.NoValueAvailable;
        }

        var bindings = new List<(string Name, string Literal)>();
        for (var i = 0; i + 1 < args.Count - 1; i += 2)
        {
            var name = args[i].Trim();
            var valueExpr = SubstituteNames(args[i + 1], bindings);
            var value = EvalScalar(sheet, valueExpr);
            bindings.Add((name, LiteralOf(value)));
        }

        var calc = SubstituteNames(args[^1], bindings);
        return EvalScalar(sheet, calc);
    }

    /// <summary>Replaces whole-word name tokens with their bound literal (longest names first to avoid prefixes).</summary>
    private static string SubstituteNames(string expr, IReadOnlyList<(string Name, string Literal)> bindings)
    {
        foreach (var (name, literal) in bindings.OrderByDescending(b => b.Name.Length))
        {
            expr = ReplaceWholeWord(expr, name, literal);
        }

        return expr;
    }

    /// <summary>Replaces every occurrence of <paramref name="word"/> not adjacent to an identifier char or inside a string.</summary>
    private static string ReplaceWholeWord(string text, string word, string replacement)
    {
        if (word.Length == 0)
        {
            return text;
        }

        var result = new System.Text.StringBuilder(text.Length);
        var inQuote = false;
        var i = 0;
        while (i < text.Length)
        {
            var ch = text[i];
            if (ch == '"')
            {
                inQuote = !inQuote;
                result.Append(ch);
                i++;
                continue;
            }

            if (!inQuote &&
                string.CompareOrdinal(text, i, word, 0, word.Length) == 0 &&
                (i == 0 || !IsIdentifierChar(text[i - 1])) &&
                (i + word.Length >= text.Length || !IsIdentifierChar(text[i + word.Length])))
            {
                result.Append(replacement);
                i += word.Length;
                continue;
            }

            result.Append(ch);
            i++;
        }

        return result.ToString();
    }

    private static bool IsIdentifierChar(char ch) => char.IsLetterOrDigit(ch) || ch == '_' || ch == '.';

    /// <summary>The literal form of a value for substitution into a formula (strings quoted, errors propagated).</summary>
    private static string LiteralOf(XLCellValue value) => value.Type switch
    {
        XLDataType.Number => "(" + value.GetNumber().ToString("R", CultureInfo.InvariantCulture) + ")",
        XLDataType.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        XLDataType.DateTime => "(" + value.GetDateTime().ToOADate().ToString("R", CultureInfo.InvariantCulture) + ")",
        XLDataType.Text => "\"" + value.GetText().Replace("\"", "\"\"", StringComparison.Ordinal) + "\"",
        XLDataType.Error => value.ToString(CultureInfo.InvariantCulture),
        _ => "\"\"",
    };

    // ----- MAXIFS / MINIFS / AVERAGEIFS --------------------------------------

    private enum Aggregate
    {
        Max,
        Min,
        Average,
    }

    /// <summary>
    /// =MAXIFS/MINIFS(agg_range, crit_range1, crit1, …) and
    /// =AVERAGEIFS(agg_range, crit_range1, crit1, …): aggregates the values of
    /// agg_range whose row satisfies every (crit_range, criteria) pair.
    /// </summary>
    private static XLCellValue ConditionalAggregate(IXLWorksheet sheet, IReadOnlyList<string> args, Aggregate kind)
    {
        if (args.Count < 3 || (args.Count - 1) % 2 != 0)
        {
            return XLError.NoValueAvailable;
        }

        var aggValues = ReadVector(sheet, args[0]);
        var criteria = new List<(List<XLCellValue> Range, Func<XLCellValue, bool> Predicate)>();
        for (var i = 1; i + 1 < args.Count; i += 2)
        {
            var range = ReadVector(sheet, args[i]);
            if (range.Count != aggValues.Count)
            {
                return XLError.NoValueAvailable; // ranges must be the same length
            }

            criteria.Add((range, CriteriaPredicate(sheet, args[i + 1])));
        }

        var matched = new List<double>();
        for (var row = 0; row < aggValues.Count; row++)
        {
            var include = true;
            foreach (var (range, predicate) in criteria)
            {
                if (!predicate(range[row]))
                {
                    include = false;
                    break;
                }
            }

            if (include && (aggValues[row].Type == XLDataType.Number || aggValues[row].Type == XLDataType.DateTime))
            {
                matched.Add(ToNumber(aggValues[row]));
            }
        }

        if (kind == Aggregate.Average)
        {
            return matched.Count == 0 ? XLError.DivisionByZero : matched.Average();
        }

        if (matched.Count == 0)
        {
            return 0d; // MAXIFS/MINIFS with no match return 0, like Excel
        }

        return kind == Aggregate.Max ? matched.Max() : matched.Min();
    }

    /// <summary>
    /// =AVERAGEIF(range, criteria, [average_range]) — the singular twin of
    /// AVERAGEIFS, the one criteria-aggregate ClosedXML 0.105 leaves as #NAME?.
    /// Its argument order differs from AVERAGEIFS (criteria is 2nd, the optional
    /// average_range is last), so the call is reshaped onto
    /// <see cref="ConditionalAggregate"/> whose order is (agg_range, crit_range,
    /// criteria): the 2-arg form lets <c>range</c> double as the average range,
    /// the 3-arg form keeps <c>range</c> as the criteria range and uses the
    /// trailing <c>average_range</c> for aggregation. Returns #DIV/0! on no match.
    /// </summary>
    private static XLCellValue AverageIf(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count is not (2 or 3))
        {
            return XLError.NoValueAvailable;
        }

        // ConditionalAggregate wants (agg_range, crit_range, criteria).
        var reshaped = args.Count == 3
            ? new[] { args[2], args[0], args[1] }  // (average_range, range, criteria)
            : new[] { args[0], args[0], args[1] }; // (range, range, criteria) — range doubles as agg range
        return ConditionalAggregate(sheet, reshaped, Aggregate.Average);
    }

    /// <summary>
    /// Builds a predicate from a criteria argument: a quoted/cell-derived string
    /// like <c>"&gt;5"</c>, <c>"&lt;=10"</c>, <c>"&lt;&gt;0"</c>, <c>"apple"</c>, or a
    /// bare number / cell reference (equality). Operators bind numerically; plain
    /// text compares case-insensitively.
    /// </summary>
    private static Func<XLCellValue, bool> CriteriaPredicate(IXLWorksheet sheet, string arg)
    {
        var criteria = CriteriaText(sheet, arg);
        var (op, operand) = SplitCriteria(criteria);

        if (double.TryParse(operand, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            return value =>
            {
                if (value.Type != XLDataType.Number && value.Type != XLDataType.DateTime)
                {
                    return op == "<>" && !ValuesEqualText(value, operand);
                }

                var n = ToNumber(value);
                return op switch
                {
                    ">" => n > number,
                    ">=" => n >= number,
                    "<" => n < number,
                    "<=" => n <= number,
                    "<>" => n != number,
                    _ => n == number,
                };
            };
        }

        // Text criteria: only = and <> are meaningful.
        return value => op == "<>"
            ? !ValuesEqualText(value, operand)
            : ValuesEqualText(value, operand);
    }

    private static (string Op, string Operand) SplitCriteria(string criteria)
    {
        var t = criteria.Trim();
        foreach (var op in new[] { "<=", ">=", "<>", ">", "<", "=" })
        {
            if (t.StartsWith(op, StringComparison.Ordinal))
            {
                return (op, t[op.Length..].Trim());
            }
        }

        return ("=", t);
    }

    /// <summary>Resolves a criteria argument to its text: a quoted literal, or a cell reference's value.</summary>
    private static string CriteriaText(IXLWorksheet sheet, string arg)
    {
        var t = arg.Trim();
        if (t.Length >= 2 && t[0] == '"' && t[^1] == '"')
        {
            return t[1..^1].Replace("\"\"", "\"", StringComparison.Ordinal);
        }

        // A bare number stays as text; anything else may be a cell reference.
        if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
        {
            return t;
        }

        var value = EvalScalar(sheet, t);
        return value.Type switch
        {
            XLDataType.Number => value.GetNumber().ToString("R", CultureInfo.InvariantCulture),
            XLDataType.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
            XLDataType.Text => value.GetText(),
            _ => value.ToString(CultureInfo.InvariantCulture),
        };
    }

    private static bool ValuesEqualText(XLCellValue value, string operand)
    {
        if ((value.Type == XLDataType.Number || value.Type == XLDataType.DateTime) &&
            double.TryParse(operand, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
        {
            return ToNumber(value) == n;
        }

        return string.Equals(
            value.ToString(CultureInfo.InvariantCulture), operand, StringComparison.OrdinalIgnoreCase);
    }

    // ----- sub-expression evaluation (delegated to ClosedXML) ----------------

    /// <summary>
    /// Evaluates one scalar sub-expression by routing it through a scratch cell so
    /// ClosedXML's engine does the arithmetic/comparison. A literal number/string
    /// is parsed directly (no scratch needed). A result of <c>#NAME?</c> means the
    /// sub-expression used a function we cannot evaluate — surfaced as an exception
    /// so the whole formula stays honestly unevaluated.
    /// </summary>
    private static XLCellValue EvalScalar(IXLWorksheet sheet, string expr)
    {
        var t = expr.Trim();
        if (t.Length == 0)
        {
            return Blank.Value;
        }

        // Fast paths that need no engine round-trip.
        if (t.Length >= 2 && t[0] == '"' && t[^1] == '"')
        {
            return t[1..^1].Replace("\"\"", "\"", StringComparison.Ordinal);
        }

        if (double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var literal))
        {
            return literal;
        }

        if (t.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (t.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var scratch = sheet.Cell(ScratchAddress);

        // Snapshot the scratch cell's original content so a workbook that legitimately
        // uses the very last cell (XFD1048576) is left exactly as it was.
        var hadFormula = scratch.HasFormula;
        var savedFormula = hadFormula ? scratch.FormulaA1 : null;
        var savedValue = hadFormula ? default : scratch.Value;
        try
        {
            scratch.FormulaA1 = t;
            var value = scratch.Value;
            if (value.IsError && value.GetError() == XLError.NameNotRecognized)
            {
                throw new NameNotRecognizedException();
            }

            return value;
        }
        finally
        {
            if (hadFormula)
            {
                scratch.FormulaA1 = savedFormula!;
            }
            else if (savedValue.Type == XLDataType.Blank)
            {
                scratch.Clear(XLClearOptions.All);
            }
            else
            {
                scratch.Value = savedValue;
            }
        }
    }

    private static double ToNumber(XLCellValue value) => value.Type switch
    {
        XLDataType.Number => value.GetNumber(),
        XLDataType.DateTime => value.GetDateTime().ToOADate(),
        XLDataType.Boolean => value.GetBoolean() ? 1 : 0,
        _ => double.TryParse(value.ToString(CultureInfo.InvariantCulture), NumberStyles.Float,
            CultureInfo.InvariantCulture, out var n)
            ? n
            : 0,
    };

    private static bool IsTruthy(XLCellValue value) => value.Type switch
    {
        XLDataType.Boolean => value.GetBoolean(),
        XLDataType.Number => value.GetNumber() != 0,
        XLDataType.Text => value.GetText().Equals("TRUE", StringComparison.OrdinalIgnoreCase),
        _ => false,
    };

    private static bool ValuesEqual(XLCellValue a, XLCellValue b)
    {
        if ((a.Type == XLDataType.Number || a.Type == XLDataType.DateTime) &&
            (b.Type == XLDataType.Number || b.Type == XLDataType.DateTime))
        {
            return ToNumber(a) == ToNumber(b);
        }

        if (a.Type == XLDataType.Boolean && b.Type == XLDataType.Boolean)
        {
            return a.GetBoolean() == b.GetBoolean();
        }

        return string.Equals(
            a.ToString(CultureInfo.InvariantCulture),
            b.ToString(CultureInfo.InvariantCulture),
            StringComparison.OrdinalIgnoreCase);
    }

    // ----- range reading -----------------------------------------------------

    /// <summary>Reads a range reference as a flat vector of typed values (row-major).</summary>
    private static List<XLCellValue> ReadVector(IXLWorksheet sheet, string reference)
    {
        var range = ResolveRange(sheet, reference);
        var values = new List<XLCellValue>();
        foreach (var cell in range.Cells())
        {
            values.Add(ValueOf(cell));
        }

        return values;
    }

    private static IXLRange ResolveRange(IXLWorksheet sheet, string reference)
    {
        var (targetSheet, local) = SplitSheet(sheet, reference.Replace("$", string.Empty, StringComparison.Ordinal));
        return local.Contains(':', StringComparison.Ordinal)
            ? targetSheet.Range(local)
            : targetSheet.Range(local + ":" + local);
    }

    private static (IXLWorksheet Sheet, string Local) SplitSheet(IXLWorksheet sheet, string reference)
    {
        var trimmed = reference.Trim();
        var bang = trimmed.LastIndexOf('!');
        if (bang < 0)
        {
            return (sheet, trimmed);
        }

        var sheetRef = trimmed[..bang].Trim('\'', ' ');
        var local = trimmed[(bang + 1)..];
        return sheet.Workbook.TryGetWorksheet(sheetRef, out var found)
            ? (found, local)
            : (sheet, local);
    }

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

    /// <summary>Signals that a sub-expression used a function aioffice cannot evaluate.</summary>
    private sealed class NameNotRecognizedException : Exception;
}
