using System.Globalization;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// (1.11) Write-time evaluation of the statistical / ranking / lookup / reference
/// functions ClosedXML 0.105 leaves as <c>#NAME?</c>:
/// <c>SMALL</c>, <c>RANK</c>/<c>RANK.EQ</c>, <c>PERCENTILE</c>/<c>PERCENTILE.INC</c>,
/// <c>QUARTILE</c>/<c>QUARTILE.INC</c>, <c>CHOOSE</c>, <c>OFFSET</c>,
/// <c>INDIRECT</c> and <c>AGGREGATE</c>. The cached value is written into the
/// saved cell so headless readers see a real result instead of riding a
/// <c>formula_not_evaluated</c> warning (the same discipline as the v1.4 financial
/// fallback and the v1.5 scalar fallback).
///
/// <para><c>SMALL</c> is the bug-fix twin of <c>LARGE</c> (which ClosedXML already
/// evaluates natively); the rest are net-new. <c>HLOOKUP</c> is NOT here: ClosedXML
/// already evaluates it correctly (verified), so it carries a cached value straight
/// from the normal save and never reaches this fallback. <c>OFFSET</c> and
/// <c>INDIRECT</c> resolve a reference to its value(s) — the evaluator already has a
/// reference-resolution path (range reads + a scratch cell), so they compute when
/// the result is a single cell or a 1-D vector the surrounding aggregate can use;
/// a result that is itself an unevaluable reference stays honestly unevaluated.</para>
///
/// <para>Sub-expression arithmetic (an argument like <c>2*B1</c> or a cell
/// reference) is delegated to ClosedXML through a far-off scratch cell that is
/// always restored, exactly like the v1.5 scalar module; a sub-expression that
/// itself uses an unevaluable function honestly returns <c>#NAME?</c> and the cell
/// keeps its warning.</para>
/// </summary>
internal static class ExcelStatisticalFunctions
{
    /// <summary>The functions this module evaluates as a save-time fallback (those ClosedXML returns #NAME? for).</summary>
    public static readonly IReadOnlyList<string> Functions =
    [
        "SMALL", "RANK", "RANK.EQ", "PERCENTILE", "PERCENTILE.INC",
        "QUARTILE", "QUARTILE.INC", "CHOOSE", "OFFSET", "INDIRECT", "AGGREGATE",
    ];

    // A scratch cell far outside any realistic used range; the evaluator writes a
    // sub-expression here, reads ClosedXML's result, and clears it afterwards.
    // Distinct from the v1.5 scalar module's scratch cell so the two never collide
    // when one calls into the other within a single save scan.
    private const string ScratchAddress = "XFD1048575";

    /// <summary>True when the (=-stripped) formula's leading function is one we evaluate.</summary>
    public static bool IsStatisticalFunction(string formula)
    {
        var name = DottedFunctionName(formula);
        return name is not null && Functions.Contains(name, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Evaluates one of this module's formulas against the sheet. Returns the typed
    /// result, or <c>#NAME?</c> when a sub-expression itself uses an unevaluable
    /// function (the caller then keeps the honest formula_not_evaluated warning), or
    /// the right Excel error (<c>#N/A</c>/<c>#VALUE!</c>/<c>#NUM!</c>/<c>#REF!</c>)
    /// for degenerate inputs. Throws nothing.
    /// </summary>
    public static XLCellValue Evaluate(IXLWorksheet sheet, string formula)
    {
        try
        {
            var body = formula.StartsWith('=') ? formula[1..] : formula;
            var name = DottedFunctionName(body)!.ToUpperInvariant();
            var args = ExcelFormulaText.SplitArgs(ExcelFormulaText.InsideParens(
                ExcelFormulaText.StripFuturePrefix(body)));
            return name switch
            {
                "SMALL" => SmallOrLarge(sheet, args, smallest: true),
                "RANK" or "RANK.EQ" => Rank(sheet, args),
                "PERCENTILE" or "PERCENTILE.INC" => Percentile(sheet, args),
                "QUARTILE" or "QUARTILE.INC" => Quartile(sheet, args),
                "CHOOSE" => Choose(sheet, args),
                "OFFSET" => Offset(sheet, args),
                "INDIRECT" => Indirect(sheet, args),
                "AGGREGATE" => AggregateFn(sheet, args),
                _ => XLError.NameNotRecognized,
            };
        }
        catch (NameNotRecognizedException)
        {
            return XLError.NameNotRecognized;
        }
        catch (CellValueException ex)
        {
            return ex.Error;
        }
        catch (Exception)
        {
            return XLError.NoValueAvailable;
        }
    }

    /// <summary>
    /// The leading function name, allowing the dots Excel uses (<c>RANK.EQ</c>,
    /// <c>PERCENTILE.INC</c>) — <see cref="ExcelFormulaText.LeadingFunctionName"/>
    /// stops at the first dot. Strips the <c>_xlfn.</c> future prefix first.
    /// </summary>
    private static string? DottedFunctionName(string formula)
    {
        var text = ExcelFormulaText.StripFuturePrefix(
            (formula.StartsWith('=') ? formula[1..] : formula).TrimStart());
        var i = 0;
        while (i < text.Length && (char.IsLetter(text[i]) || char.IsDigit(text[i]) || text[i] == '_' || text[i] == '.'))
        {
            i++;
        }

        return i > 0 && i < text.Length && text[i] == '(' ? text[..i] : null;
    }

    // ----- SMALL (LARGE twin) ------------------------------------------------

    /// <summary>
    /// =SMALL(array,k) — the k-th SMALLEST value (1-based). Mirrors Excel's LARGE
    /// (which ClosedXML evaluates natively): numeric entries only, k outside
    /// [1, count] is <c>#NUM!</c>, an empty array is <c>#NUM!</c>.
    /// </summary>
    private static XLCellValue SmallOrLarge(IXLWorksheet sheet, IReadOnlyList<string> args, bool smallest)
    {
        if (args.Count != 2)
        {
            return XLError.IncompatibleValue;
        }

        var numbers = ReadNumbers(sheet, args[0]);
        var k = (int)Math.Floor(ToNumber(EvalScalar(sheet, args[1])));
        if (numbers.Count == 0 || k < 1 || k > numbers.Count)
        {
            return XLError.NumberInvalid; // #NUM!
        }

        numbers.Sort();
        return smallest ? numbers[k - 1] : numbers[^k];
    }

    // ----- RANK / RANK.EQ ----------------------------------------------------

    /// <summary>
    /// =RANK(number,ref,[order]) / =RANK.EQ(...) — the rank of <c>number</c> within
    /// <c>ref</c>. order omitted/0 = DESCENDING (largest is rank 1); order != 0 =
    /// ascending. Ties share the same (top) rank. <c>#N/A</c> when the number is not
    /// present in ref. (RANK.EQ is identical to RANK.)
    /// </summary>
    private static XLCellValue Rank(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count is < 2 or > 3)
        {
            return XLError.IncompatibleValue;
        }

        var number = ToNumber(EvalScalar(sheet, args[0]));
        var numbers = ReadNumbers(sheet, args[1]);
        var descending = !(args.Count == 3 && args[2].Trim().Length > 0 &&
            ToNumber(EvalScalar(sheet, args[2])) != 0);

        if (!numbers.Contains(number))
        {
            return XLError.NoValueAvailable; // #N/A: the value is not in ref
        }

        // Rank = 1 + (count of values strictly better). "Better" = larger when
        // descending, smaller when ascending. Ties therefore share the top rank.
        var better = descending
            ? numbers.Count(v => v > number)
            : numbers.Count(v => v < number);
        return (double)(better + 1);
    }

    // ----- PERCENTILE / QUARTILE (inclusive) ---------------------------------

    /// <summary>
    /// =PERCENTILE(array,k) / =PERCENTILE.INC(array,k) — the k-th percentile,
    /// k in [0,1], INCLUSIVE method: position = k*(n-1) over the sorted array,
    /// linearly interpolated between the two bracketing values. k outside [0,1] or
    /// an empty array is <c>#NUM!</c>.
    /// </summary>
    private static XLCellValue Percentile(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count != 2)
        {
            return XLError.IncompatibleValue;
        }

        var numbers = ReadNumbers(sheet, args[0]);
        var k = ToNumber(EvalScalar(sheet, args[1]));
        return PercentileInc(numbers, k);
    }

    /// <summary>
    /// =QUARTILE(array,q) / =QUARTILE.INC(array,q) — q in {0,1,2,3,4}, identical to
    /// PERCENTILE.INC(array, q/4) (0->min, 2->median, 4->max). q outside {0..4} is
    /// <c>#NUM!</c>.
    /// </summary>
    private static XLCellValue Quartile(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count != 2)
        {
            return XLError.IncompatibleValue;
        }

        var numbers = ReadNumbers(sheet, args[0]);
        var q = (int)Math.Floor(ToNumber(EvalScalar(sheet, args[1])));
        if (q is < 0 or > 4)
        {
            return XLError.NumberInvalid; // #NUM!
        }

        return PercentileInc(numbers, q / 4.0);
    }

    /// <summary>The inclusive-method percentile of a numeric list, or #NUM! on bad inputs.</summary>
    private static XLCellValue PercentileInc(List<double> numbers, double k)
    {
        if (numbers.Count == 0 || k < 0 || k > 1)
        {
            return XLError.NumberInvalid; // #NUM!
        }

        numbers.Sort();
        if (numbers.Count == 1)
        {
            return numbers[0];
        }

        var position = k * (numbers.Count - 1);
        var lower = (int)Math.Floor(position);
        var fraction = position - lower;
        if (lower >= numbers.Count - 1)
        {
            return numbers[^1];
        }

        return numbers[lower] + (fraction * (numbers[lower + 1] - numbers[lower]));
    }

    /// <summary>
    /// The exclusive-method percentile (PERCENTILE.EXC) of a numeric list. Unlike the
    /// inclusive twin this interpolates at the 1-based position k*(n+1) and is valid
    /// only for k in the OPEN interval (1/(n+1), n/(n+1)); anything outside (including
    /// k=0 and k=1) is <c>#NUM!</c>.
    /// </summary>
    private static XLCellValue PercentileExc(List<double> numbers, double k)
    {
        var n = numbers.Count;
        // Excel's exclusive percentile is #NUM! only OUTSIDE [1/(n+1), n/(n+1)] — the
        // endpoints themselves are valid (k=1/(n+1) -> the min, k=n/(n+1) -> the max).
        if (n == 0 || k < 1.0 / (n + 1) || k > (double)n / (n + 1))
        {
            return XLError.NumberInvalid; // #NUM!
        }

        numbers.Sort();
        var position = k * (n + 1); // 1-based, in [1, n]
        var lower = (int)Math.Floor(position);
        if (lower >= n)
        {
            return numbers[n - 1]; // position == n (the upper endpoint): the max, no run past the array
        }

        var fraction = position - lower;
        return numbers[lower - 1] + (fraction * (numbers[lower] - numbers[lower - 1]));
    }

    /// <summary>
    /// MODE.SNGL — the smallest value among the equally-most-frequent. When every value
    /// is unique (no repeats) or the list is empty, Excel returns <c>#N/A</c>.
    /// </summary>
    private static XLCellValue ModeSngl(List<double> numbers)
    {
        var bestValue = 0d;
        var bestCount = 0;
        foreach (var group in numbers.GroupBy(v => v))
        {
            var count = group.Count();
            if (count > bestCount || (count == bestCount && group.Key < bestValue))
            {
                bestValue = group.Key;
                bestCount = count;
            }
        }

        return bestCount < 2 ? XLError.NoValueAvailable : bestValue; // #N/A when no value repeats
    }

    // ----- CHOOSE ------------------------------------------------------------

    /// <summary>
    /// =CHOOSE(index, v1, v2, …) — the 1-based index-th value. index &lt; 1 or
    /// beyond the value count is <c>#VALUE!</c>. The chosen value is evaluated (it
    /// may be a cell reference or arithmetic).
    /// </summary>
    private static XLCellValue Choose(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count < 2)
        {
            return XLError.IncompatibleValue;
        }

        var index = (int)Math.Floor(ToNumber(EvalScalar(sheet, args[0])));
        if (index < 1 || index > args.Count - 1)
        {
            return XLError.IncompatibleValue; // #VALUE!
        }

        return EvalScalar(sheet, args[index]);
    }

    // ----- OFFSET / INDIRECT (reference functions) ---------------------------

    /// <summary>
    /// =OFFSET(reference, rows, cols, [height], [width]) — the cell (or range)
    /// shifted from <c>reference</c> by rows/cols, optionally resized to
    /// height x width. A 1x1 result returns the scalar value; a larger result
    /// surfaces its top-left value when used as a scalar (Excel implicit
    /// intersection in the common headless case). A shift off the sheet is
    /// <c>#REF!</c>.
    /// </summary>
    private static XLCellValue Offset(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count is < 3 or > 5)
        {
            return XLError.IncompatibleValue;
        }

        var (targetSheet, baseRange) = ResolveReference(sheet, args[0]);
        var rowOffset = (int)Math.Floor(ToNumber(EvalScalar(sheet, args[1])));
        var colOffset = (int)Math.Floor(ToNumber(EvalScalar(sheet, args[2])));

        var baseFirstRow = baseRange.RangeAddress.FirstAddress.RowNumber;
        var baseFirstCol = baseRange.RangeAddress.FirstAddress.ColumnNumber;
        var baseHeight = baseRange.RowCount();
        var baseWidth = baseRange.ColumnCount();

        var height = args.Count >= 4 && args[3].Trim().Length > 0
            ? (int)Math.Floor(ToNumber(EvalScalar(sheet, args[3])))
            : baseHeight;
        var width = args.Count >= 5 && args[4].Trim().Length > 0
            ? (int)Math.Floor(ToNumber(EvalScalar(sheet, args[4])))
            : baseWidth;

        var firstRow = baseFirstRow + rowOffset;
        var firstCol = baseFirstCol + colOffset;
        if (height <= 0 || width <= 0 || firstRow < 1 || firstCol < 1 ||
            firstRow + height - 1 > 1_048_576 || firstCol + width - 1 > 16_384)
        {
            throw new CellValueException(XLError.CellReference); // #REF!
        }

        // A 1x1 OFFSET resolves to a scalar; a larger result surfaces its top-left
        // cell when used as a scalar. Aggregates that wrap OFFSET are not
        // intercepted here.
        return ValueOf(targetSheet.Cell(firstRow, firstCol));
    }

    /// <summary>
    /// =INDIRECT(ref_text, [a1]) — the value(s) at the reference named by
    /// <c>ref_text</c> (e.g. "A3" or "Sheet1!B2"). a1 defaults TRUE; R1C1 form is
    /// not parsed (rare in authored files) and yields <c>#REF!</c>. A multi-cell
    /// reference used as a scalar surfaces its top-left value. An unparsable
    /// reference is <c>#REF!</c>.
    /// </summary>
    private static XLCellValue Indirect(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count is < 1 or > 2)
        {
            return XLError.IncompatibleValue;
        }

        var a1 = !(args.Count == 2 && args[1].Trim().Length > 0 &&
            ToNumber(EvalScalar(sheet, args[1])) == 0);
        if (!a1)
        {
            throw new CellValueException(XLError.CellReference); // R1C1 unsupported
        }

        var refText = AsText(EvalScalar(sheet, args[0]));
        if (refText.Length == 0)
        {
            throw new CellValueException(XLError.CellReference); // #REF!
        }

        try
        {
            var (_, range) = ResolveReference(sheet, refText);
            return ValueOf(range.FirstCell());
        }
        catch (CellValueException)
        {
            throw;
        }
        catch (Exception)
        {
            throw new CellValueException(XLError.CellReference); // #REF!
        }
    }

    // ----- AGGREGATE ---------------------------------------------------------

    /// <summary>
    /// =AGGREGATE(function_num, options, ref1, [k]) — an aggregate that can ignore
    /// errors and hidden rows. Supported function_num: 1 AVERAGE, 2 COUNT, 3 COUNTA,
    /// 4 MAX, 5 MIN, 6 PRODUCT, 7 STDEV.S, 8 STDEV.P, 9 SUM, 10 VAR.S, 11 VAR.P,
    /// 12 MEDIAN, 13 MODE.SNGL, 14 LARGE, 15 SMALL, 16 PERCENTILE.INC, 17 QUARTILE.INC,
    /// 18 PERCENTILE.EXC, 19 QUARTILE.EXC. options 0-7 are accepted; 2/3/6/7 ignore
    /// error values, 0/1/4/5 propagate them (the hidden-row distinction is not modeled
    /// headlessly). Only the array form is deferred — it falls through to the honest
    /// formula_not_evaluated warning.
    /// </summary>
    private static XLCellValue AggregateFn(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count < 3)
        {
            return XLError.IncompatibleValue;
        }

        var functionNum = (int)Math.Floor(ToNumber(EvalScalar(sheet, args[0])));
        var options = (int)Math.Floor(ToNumber(EvalScalar(sheet, args[1])));
        if (options is < 0 or > 7)
        {
            return XLError.IncompatibleValue; // #VALUE! — options must be 0..7
        }

        // Excel options 2/3/6/7 IGNORE error values; 0/1/4/5 do NOT — an error in the
        // referenced range then propagates as the result. (The hidden-row distinction
        // in 1/3/5/7 is not modeled headlessly — a documented limitation.) Blanks are
        // always skipped by the aggregate functions regardless.
        if (options is not (2 or 3 or 6 or 7))
        {
            var firstError = FirstError(sheet, args[2]);
            if (firstError is not null)
            {
                return firstError.Value;
            }
        }

        var numbers = ReadNumbers(sheet, args[2]); // numeric cells; blanks/errors dropped

        // 14-17 take a trailing k / q argument.
        double K()
        {
            if (args.Count < 4)
            {
                throw new CellValueException(XLError.IncompatibleValue);
            }

            return ToNumber(EvalScalar(sheet, args[3]));
        }

        switch (functionNum)
        {
            case 1: return numbers.Count == 0 ? XLError.DivisionByZero : numbers.Average();
            case 2: return (double)numbers.Count; // COUNT: numeric cells (blanks/errors dropped)
            case 3: return (double)CountA(sheet, args[2]); // COUNTA: all non-empty cells
            case 4: return numbers.Count == 0 ? (XLCellValue)0d : numbers.Max();
            case 5: return numbers.Count == 0 ? (XLCellValue)0d : numbers.Min();
            case 6: return numbers.Aggregate(1d, (acc, v) => acc * v);
            case 7: return SampleStdev(numbers);
            case 8: return PopulationStdev(numbers);
            case 9: return numbers.Sum();
            case 10: return SampleVariance(numbers);
            case 11: return PopulationVariance(numbers);
            case 12: return Median(numbers);
            case 14: return NthLargest(numbers, (int)Math.Floor(K()), smallest: false);
            case 15: return NthLargest(numbers, (int)Math.Floor(K()), smallest: true);
            case 16: return PercentileInc(numbers, K());
            case 17:
                var q = (int)Math.Floor(K());
                return q is < 0 or > 4 ? XLError.NumberInvalid : PercentileInc(numbers, q / 4.0);
            case 13: return ModeSngl(numbers);
            case 18: return PercentileExc(numbers, K());
            case 19:
                var qExc = (int)Math.Floor(K());
                // EXC quartile: only q in {1,2,3} is valid — q=0 and q=4 are #NUM!
                // (the EXC-specific guard, NOT the inclusive 0..4 range used by 17).
                return qExc is < 1 or > 3 ? XLError.NumberInvalid : PercentileExc(numbers, qExc / 4.0);
            default:
                // Any out-of-range func_num (e.g. 20) and the dynamic/array form
                // stay honestly unevaluated (the cell keeps formula_not_evaluated).
                throw new NameNotRecognizedException();
        }
    }

    private static XLCellValue NthLargest(List<double> numbers, int k, bool smallest)
    {
        if (numbers.Count == 0 || k < 1 || k > numbers.Count)
        {
            return XLError.NumberInvalid;
        }

        numbers.Sort();
        return smallest ? numbers[k - 1] : numbers[^k];
    }

    private static XLCellValue Median(List<double> numbers)
    {
        if (numbers.Count == 0)
        {
            return XLError.NumberInvalid;
        }

        numbers.Sort();
        var mid = numbers.Count / 2;
        return numbers.Count % 2 == 1 ? numbers[mid] : (numbers[mid - 1] + numbers[mid]) / 2.0;
    }

    private static XLCellValue SampleVariance(List<double> numbers)
    {
        if (numbers.Count < 2)
        {
            return XLError.DivisionByZero;
        }

        var mean = numbers.Average();
        return numbers.Sum(v => (v - mean) * (v - mean)) / (numbers.Count - 1);
    }

    private static XLCellValue PopulationVariance(List<double> numbers)
    {
        if (numbers.Count == 0)
        {
            return XLError.DivisionByZero;
        }

        var mean = numbers.Average();
        return numbers.Sum(v => (v - mean) * (v - mean)) / numbers.Count;
    }

    private static XLCellValue SampleStdev(List<double> numbers)
    {
        var variance = SampleVariance(numbers);
        return variance.IsError ? variance : Math.Sqrt(variance.GetNumber());
    }

    private static XLCellValue PopulationStdev(List<double> numbers)
    {
        var variance = PopulationVariance(numbers);
        return variance.IsError ? variance : Math.Sqrt(variance.GetNumber());
    }

    /// <summary>COUNTA semantics over a reference: every non-blank cell counts (text, numbers, errors).</summary>
    private static int CountA(IXLWorksheet sheet, string reference)
    {
        var range = ResolveRange(sheet, reference);
        var count = 0;
        foreach (var cell in range.Cells())
        {
            if (ValueOf(cell).Type != XLDataType.Blank)
            {
                count++;
            }
        }

        return count;
    }

    // ----- reference + range resolution --------------------------------------

    /// <summary>
    /// Resolves a reference argument (a range/cell reference, possibly sheet-qualified)
    /// to its worksheet and range. Throws <see cref="CellValueException"/> with
    /// <c>#REF!</c> when the text is not a resolvable reference.
    /// </summary>
    private static (IXLWorksheet Sheet, IXLRange Range) ResolveReference(IXLWorksheet sheet, string reference)
    {
        var (targetSheet, local) = SplitSheet(sheet, reference.Replace("$", string.Empty, StringComparison.Ordinal));
        try
        {
            var range = local.Contains(':', StringComparison.Ordinal)
                ? targetSheet.Range(local)
                : targetSheet.Range(local + ":" + local);
            return (targetSheet, range);
        }
        catch (Exception)
        {
            throw new CellValueException(XLError.CellReference);
        }
    }

    private static IXLRange ResolveRange(IXLWorksheet sheet, string reference)
    {
        var (_, range) = ResolveReference(sheet, reference);
        return range;
    }

    /// <summary>Reads a reference's numeric values (text/blank/error/boolean cells are skipped).</summary>
    private static List<double> ReadNumbers(IXLWorksheet sheet, string reference)
    {
        var trimmed = reference.Trim();
        if (LooksLikeReference(trimmed))
        {
            var range = ResolveRange(sheet, trimmed);
            var values = new List<double>();
            foreach (var cell in range.Cells())
            {
                var value = ValueOf(cell);
                if (value.Type is XLDataType.Number or XLDataType.DateTime)
                {
                    values.Add(ToNumber(value));
                }
            }

            return values;
        }

        // A brace array literal {1,2,3} (or {1;2;3}) read as a number list.
        if (trimmed.StartsWith('{') && trimmed.EndsWith('}'))
        {
            var values = new List<double>();
            foreach (var item in ExcelFormulaText.SplitArgs(trimmed[1..^1].Replace(';', ',')))
            {
                var value = EvalScalar(sheet, item);
                if (value.Type is XLDataType.Number or XLDataType.DateTime)
                {
                    values.Add(ToNumber(value));
                }
            }

            return values;
        }

        // A single scalar expression.
        var scalar = EvalScalar(sheet, trimmed);
        return scalar.Type is XLDataType.Number or XLDataType.DateTime
            ? [ToNumber(scalar)]
            : [];
    }

    /// <summary>The first error value found in a reference's cells, or null if none.
    /// Used by AGGREGATE options 0/1/4/5, which (unlike 2/3/6/7) do NOT ignore errors —
    /// an error in the range propagates as the aggregate's result.</summary>
    private static XLError? FirstError(IXLWorksheet sheet, string reference)
    {
        var trimmed = reference.Trim();
        if (!LooksLikeReference(trimmed))
        {
            // A brace literal {…} carries no cell errors; a scalar arg's error would
            // already have surfaced through EvalScalar.
            return null;
        }

        foreach (var cell in ResolveRange(sheet, trimmed).Cells())
        {
            var value = ValueOf(cell);
            if (value.IsError)
            {
                return value.GetError();
            }
        }

        return null;
    }

    /// <summary>True when the text is plausibly a cell/range reference (so we read it as a range, not arithmetic).</summary>
    private static bool LooksLikeReference(string text)
    {
        var t = text;
        var bang = t.LastIndexOf('!');
        if (bang >= 0)
        {
            t = t[(bang + 1)..];
        }

        t = t.Replace("$", string.Empty, StringComparison.Ordinal);
        if (t.Length == 0)
        {
            return false;
        }

        foreach (var part in t.Split(':'))
        {
            if (!IsA1Cell(part))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsA1Cell(string part)
    {
        var i = 0;
        while (i < part.Length && char.IsLetter(part[i]))
        {
            i++;
        }

        if (i == 0 || i >= part.Length)
        {
            return false;
        }

        for (var j = i; j < part.Length; j++)
        {
            if (!char.IsDigit(part[j]))
            {
                return false;
            }
        }

        return true;
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

    // ----- sub-expression evaluation (delegated to ClosedXML) ----------------

    /// <summary>
    /// Evaluates one scalar sub-expression by routing it through a scratch cell so
    /// ClosedXML's engine does the arithmetic. Literals and bare cell references are
    /// resolved directly. A <c>#NAME?</c> result means the sub-expression used an
    /// unevaluable function — surfaced as an exception so the whole formula stays
    /// honestly unevaluated.
    /// </summary>
    private static XLCellValue EvalScalar(IXLWorksheet sheet, string expr)
    {
        var t = expr.Trim();
        if (t.Length == 0)
        {
            return Blank.Value;
        }

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

        // A bare single-cell reference reads the cell directly (no engine round-trip).
        if (!t.Contains(':', StringComparison.Ordinal) && LooksLikeReference(t))
        {
            return ValueOf(ResolveRange(sheet, t).FirstCell());
        }

        var scratch = sheet.Cell(ScratchAddress);
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

    private static string AsText(XLCellValue value) => value.Type switch
    {
        XLDataType.Text => value.GetText(),
        XLDataType.Number => value.GetNumber().ToString("R", CultureInfo.InvariantCulture),
        XLDataType.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        _ => value.ToString(CultureInfo.InvariantCulture),
    };

    /// <summary>Signals that a sub-expression used a function aioffice cannot evaluate.</summary>
    private sealed class NameNotRecognizedException : Exception;

    /// <summary>Carries a specific Excel error value out of a nested helper.</summary>
    private sealed class CellValueException(XLError error) : Exception
    {
        public XLError Error { get; } = error;
    }
}
