using System.Globalization;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// (1.4) Write-time evaluation of the financial functions ClosedXML 0.105 leaves
/// as <c>#NAME?</c> / unevaluated. The iterative trio (<c>RATE</c>, <c>IRR</c>,
/// <c>XIRR</c>) is solved with Newton's method plus a bisection fallback; the
/// closed-form <c>PV</c>, <c>NPV</c> and <c>NPER</c> (which ClosedXML also cannot
/// do) are computed directly. The cached value is written into the saved cell, so
/// headless readers see a real number instead of a <c>formula_not_evaluated</c>
/// warning.
///
/// <para><c>PMT</c> and <c>FV</c> are already evaluated by ClosedXML's engine;
/// they appear here too as a fallback (this module only runs when ClosedXML
/// returned <c>#NAME?</c> or threw), so the full <c>PMT/PV/FV/NPV/NPER</c> set is
/// guaranteed to produce a cached value.</para>
/// </summary>
internal static class ExcelFinancialFunctions
{
    /// <summary>Every financial function this module can evaluate as a save-time fallback.</summary>
    public static readonly IReadOnlyList<string> Functions =
        ["RATE", "IRR", "XIRR", "NPV", "PV", "FV", "PMT", "NPER"];

    private const double Tolerance = 1e-7;
    private const int MaxIterations = 100;

    /// <summary>True when the (=-stripped) formula's leading function is one we evaluate.</summary>
    public static bool IsFinancial(string formula)
    {
        var name = LeadingFunctionName(formula);
        return name is not null && Functions.Contains(name, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Evaluates a financial formula against the sheet (its cell references are
    /// read from the live model). Returns the numeric result, or a <c>#NUM!</c>
    /// error value when the iteration does not converge / the inputs are
    /// degenerate. Throws nothing — the save pass writes whatever comes back.
    /// </summary>
    public static XLCellValue Evaluate(IXLWorksheet sheet, string formula)
    {
        try
        {
            var body = formula.StartsWith('=') ? formula[1..] : formula;
            var name = LeadingFunctionName(body)!.ToUpperInvariant();
            var args = SplitArgs(InsideParens(body));
            return name switch
            {
                "RATE" => Rate(sheet, args),
                "IRR" => Irr(sheet, args),
                "XIRR" => Xirr(sheet, args),
                "NPV" => Npv(sheet, args),
                "PV" => Pv(sheet, args),
                "FV" => Fv(sheet, args),
                "PMT" => Pmt(sheet, args),
                "NPER" => Nper(sheet, args),
                _ => XLError.NoValueAvailable,
            };
        }
        catch (Exception)
        {
            return XLError.NumberInvalid;
        }
    }

    // ----- RATE --------------------------------------------------------------

    /// <summary>=RATE(nper, pmt, pv, [fv], [type], [guess]) — periodic rate by Newton's method.</summary>
    private static XLCellValue Rate(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        var nper = Num(sheet, args, 0);
        var pmt = Num(sheet, args, 1);
        var pv = Num(sheet, args, 2);
        var fv = OptNum(sheet, args, 3, 0);
        var type = OptNum(sheet, args, 4, 0);
        var guess = OptNum(sheet, args, 5, 0.1);

        double TvmResidual(double rate) => rate == 0
            ? pv + (pmt * nper) + fv
            : (pv * Math.Pow(1 + rate, nper)) +
              (pmt * (1 + (rate * type)) * ((Math.Pow(1 + rate, nper) - 1) / rate)) + fv;

        var x = guess;
        for (var i = 0; i < MaxIterations; i++)
        {
            var f = TvmResidual(x);
            if (Math.Abs(f) < Tolerance)
            {
                return x;
            }

            // Numeric derivative (the closed form is unwieldy with the type flag).
            var h = 1e-6;
            var derivative = (TvmResidual(x + h) - TvmResidual(x - h)) / (2 * h);
            if (Math.Abs(derivative) < 1e-12)
            {
                break;
            }

            var next = x - (f / derivative);
            if (double.IsNaN(next) || double.IsInfinity(next))
            {
                break;
            }

            if (Math.Abs(next - x) < Tolerance)
            {
                return next;
            }

            x = next;
        }

        return Math.Abs(TvmResidual(x)) < 1e-4 ? x : XLError.NumberInvalid;
    }

    // ----- IRR ---------------------------------------------------------------

    /// <summary>=IRR(values, [guess]) — internal rate of return on evenly-spaced cash flows.</summary>
    private static XLCellValue Irr(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        var cashflows = ReadNumbers(sheet, args[0]);
        if (cashflows.Count < 2)
        {
            return XLError.NumberInvalid;
        }

        var guess = OptNum(sheet, args, 1, 0.1);

        double Npv(double rate)
        {
            var total = 0d;
            for (var t = 0; t < cashflows.Count; t++)
            {
                total += cashflows[t] / Math.Pow(1 + rate, t);
            }

            return total;
        }

        return Solve(Npv, guess);
    }

    // ----- XIRR --------------------------------------------------------------

    /// <summary>=XIRR(values, dates, [guess]) — IRR on irregularly-spaced cash flows.</summary>
    private static XLCellValue Xirr(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        var cashflows = ReadNumbers(sheet, args[0]);
        var dates = ReadDates(sheet, args[1]);
        if (cashflows.Count < 2 || cashflows.Count != dates.Count)
        {
            return XLError.NumberInvalid;
        }

        var d0 = dates[0];
        var guess = OptNum(sheet, args, 2, 0.1);

        double Xnpv(double rate)
        {
            var total = 0d;
            for (var i = 0; i < cashflows.Count; i++)
            {
                var years = (dates[i] - d0).TotalDays / 365.0;
                total += cashflows[i] / Math.Pow(1 + rate, years);
            }

            return total;
        }

        return Solve(Xnpv, guess);
    }

    // ----- closed-form (ClosedXML cannot do NPV/PV/NPER; PMT/FV are fallbacks) -

    /// <summary>=NPV(rate, value1, [value2], …) — net present value of period-end flows.</summary>
    private static XLCellValue Npv(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        if (args.Count < 2)
        {
            return XLError.NumberInvalid;
        }

        var rate = ReadScalar(sheet, args[0]);
        var total = 0d;
        var period = 1;
        for (var i = 1; i < args.Count; i++)
        {
            foreach (var value in ReadNumbers(sheet, args[i]))
            {
                total += value / Math.Pow(1 + rate, period);
                period++;
            }
        }

        return total;
    }

    /// <summary>=PV(rate, nper, pmt, [fv], [type]) — present value.</summary>
    private static XLCellValue Pv(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        var rate = Num(sheet, args, 0);
        var nper = Num(sheet, args, 1);
        var pmt = Num(sheet, args, 2);
        var fv = OptNum(sheet, args, 3, 0);
        var type = OptNum(sheet, args, 4, 0);
        if (rate == 0)
        {
            return -(fv + (pmt * nper));
        }

        var factor = Math.Pow(1 + rate, nper);
        return -(fv + (pmt * (1 + (rate * type)) * ((factor - 1) / rate))) / factor;
    }

    /// <summary>=FV(rate, nper, pmt, [pv], [type]) — future value (fallback; ClosedXML usually handles it).</summary>
    private static XLCellValue Fv(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        var rate = Num(sheet, args, 0);
        var nper = Num(sheet, args, 1);
        var pmt = Num(sheet, args, 2);
        var pv = OptNum(sheet, args, 3, 0);
        var type = OptNum(sheet, args, 4, 0);
        if (rate == 0)
        {
            return -(pv + (pmt * nper));
        }

        var factor = Math.Pow(1 + rate, nper);
        return -((pv * factor) + (pmt * (1 + (rate * type)) * ((factor - 1) / rate)));
    }

    /// <summary>=PMT(rate, nper, pv, [fv], [type]) — payment (fallback; ClosedXML usually handles it).</summary>
    private static XLCellValue Pmt(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        var rate = Num(sheet, args, 0);
        var nper = Num(sheet, args, 1);
        var pv = Num(sheet, args, 2);
        var fv = OptNum(sheet, args, 3, 0);
        var type = OptNum(sheet, args, 4, 0);
        if (nper == 0)
        {
            return XLError.DivisionByZero;
        }

        if (rate == 0)
        {
            return -(pv + fv) / nper;
        }

        var factor = Math.Pow(1 + rate, nper);
        return -(fv + (pv * factor)) * rate / ((1 + (rate * type)) * (factor - 1));
    }

    /// <summary>=NPER(rate, pmt, pv, [fv], [type]) — number of periods.</summary>
    private static XLCellValue Nper(IXLWorksheet sheet, IReadOnlyList<string> args)
    {
        var rate = Num(sheet, args, 0);
        var pmt = Num(sheet, args, 1);
        var pv = Num(sheet, args, 2);
        var fv = OptNum(sheet, args, 3, 0);
        var type = OptNum(sheet, args, 4, 0);
        if (rate == 0)
        {
            return pmt == 0 ? XLError.NumberInvalid : -(pv + fv) / pmt;
        }

        var adjusted = pmt * (1 + (rate * type));
        var numerator = adjusted - (fv * rate);
        var denominator = (pv * rate) + adjusted;
        if (numerator / denominator <= 0)
        {
            return XLError.NumberInvalid;
        }

        return Math.Log(numerator / denominator) / Math.Log(1 + rate);
    }

    /// <summary>Newton's method with a bisection fallback over [-0.999, large].</summary>
    private static XLCellValue Solve(Func<double, double> f, double guess)
    {
        var x = guess;
        for (var i = 0; i < MaxIterations; i++)
        {
            var fx = f(x);
            if (Math.Abs(fx) < Tolerance)
            {
                return x;
            }

            var h = 1e-6;
            var derivative = (f(x + h) - f(x - h)) / (2 * h);
            if (Math.Abs(derivative) < 1e-12)
            {
                break;
            }

            var next = x - (fx / derivative);
            if (double.IsNaN(next) || double.IsInfinity(next) || next <= -1)
            {
                break;
            }

            if (Math.Abs(next - x) < Tolerance)
            {
                return next;
            }

            x = next;
        }

        // Bisection fallback: scan for a sign change, then halve in.
        var low = -0.9999;
        var high = 10.0;
        var fLow = f(low);
        var fHigh = f(high);
        if (double.IsNaN(fLow) || double.IsNaN(fHigh) || fLow * fHigh > 0)
        {
            return XLError.NumberInvalid;
        }

        for (var i = 0; i < MaxIterations; i++)
        {
            var mid = (low + high) / 2;
            var fMid = f(mid);
            if (Math.Abs(fMid) < Tolerance || (high - low) / 2 < Tolerance)
            {
                return mid;
            }

            if (fLow * fMid < 0)
            {
                high = mid;
                fHigh = fMid;
            }
            else
            {
                low = mid;
                fLow = fMid;
            }
        }

        return (low + high) / 2;
    }

    // ----- arg reading -------------------------------------------------------

    private static double Num(IXLWorksheet sheet, IReadOnlyList<string> args, int i) =>
        i < args.Count ? ReadScalar(sheet, args[i]) : 0;

    private static double OptNum(IXLWorksheet sheet, IReadOnlyList<string> args, int i, double fallback) =>
        i < args.Count && args[i].Trim().Length > 0 ? ReadScalar(sheet, args[i]) : fallback;

    /// <summary>
    /// Reads one scalar argument: a number/cell reference, or a simple arithmetic
    /// expression over those (<c>+ - * /</c>, parentheses, unary minus) — enough
    /// for the common <c>0.08/12</c> rate idiom that ClosedXML would otherwise have
    /// handled for us.
    /// </summary>
    private static double ReadScalar(IXLWorksheet sheet, string arg)
    {
        var parser = new ScalarExpression(sheet, arg);
        return parser.Parse();
    }

    /// <summary>A tiny recursive-descent evaluator for scalar formula arguments.</summary>
    private sealed class ScalarExpression(IXLWorksheet sheet, string text)
    {
        private readonly string _text = text;
        private int _pos;

        public double Parse()
        {
            var value = ParseAddSub();
            return value;
        }

        private double ParseAddSub()
        {
            var value = ParseMulDiv();
            while (true)
            {
                SkipWhite();
                if (Peek() == '+')
                {
                    _pos++;
                    value += ParseMulDiv();
                }
                else if (Peek() == '-')
                {
                    _pos++;
                    value -= ParseMulDiv();
                }
                else
                {
                    return value;
                }
            }
        }

        private double ParseMulDiv()
        {
            var value = ParseUnary();
            while (true)
            {
                SkipWhite();
                if (Peek() == '*')
                {
                    _pos++;
                    value *= ParseUnary();
                }
                else if (Peek() == '/')
                {
                    _pos++;
                    value /= ParseUnary();
                }
                else
                {
                    return value;
                }
            }
        }

        private double ParseUnary()
        {
            SkipWhite();
            if (Peek() == '-')
            {
                _pos++;
                return -ParseUnary();
            }

            if (Peek() == '+')
            {
                _pos++;
                return ParseUnary();
            }

            return ParseAtom();
        }

        private double ParseAtom()
        {
            SkipWhite();
            if (Peek() == '(')
            {
                _pos++;
                var value = ParseAddSub();
                SkipWhite();
                if (Peek() == ')')
                {
                    _pos++;
                }

                return value;
            }

            var start = _pos;
            while (_pos < _text.Length && !"+-*/()".Contains(_text[_pos]))
            {
                _pos++;
            }

            var token = _text[start.._pos].Trim().Replace("$", string.Empty, StringComparison.Ordinal);
            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var literal))
            {
                return literal;
            }

            var value2 = CellValue(sheet, token);
            return value2.Type switch
            {
                XLDataType.Number => value2.GetNumber(),
                XLDataType.DateTime => value2.GetDateTime().ToOADate(),
                XLDataType.Boolean => value2.GetBoolean() ? 1 : 0,
                _ => double.TryParse(value2.ToString(CultureInfo.InvariantCulture), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var n)
                    ? n
                    : 0,
            };
        }

        private char Peek() => _pos < _text.Length ? _text[_pos] : '\0';

        private void SkipWhite()
        {
            while (_pos < _text.Length && char.IsWhiteSpace(_text[_pos]))
            {
                _pos++;
            }
        }
    }

    private static List<double> ReadNumbers(IXLWorksheet sheet, string reference)
    {
        var range = ResolveRange(sheet, reference);
        var numbers = new List<double>();
        foreach (var cell in range.Cells())
        {
            var v = ValueOf(cell);
            if (v.Type == XLDataType.Number)
            {
                numbers.Add(v.GetNumber());
            }
            else if (v.Type == XLDataType.DateTime)
            {
                numbers.Add(v.GetDateTime().ToOADate());
            }
        }

        return numbers;
    }

    private static List<DateTime> ReadDates(IXLWorksheet sheet, string reference)
    {
        var range = ResolveRange(sheet, reference);
        var dates = new List<DateTime>();
        foreach (var cell in range.Cells())
        {
            var v = ValueOf(cell);
            if (v.Type == XLDataType.DateTime)
            {
                dates.Add(v.GetDateTime());
            }
            else if (v.Type == XLDataType.Number)
            {
                dates.Add(DateTime.FromOADate(v.GetNumber()));
            }
        }

        return dates;
    }

    private static XLCellValue CellValue(IXLWorksheet sheet, string reference)
    {
        var (targetSheet, local) = SplitSheet(sheet, reference);
        return ValueOf(targetSheet.Cell(local));
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

    // ----- formula text helpers (shared shape with ExcelDynamicArrays) -------

    private static string? LeadingFunctionName(string formula)
    {
        var text = (formula.StartsWith('=') ? formula[1..] : formula).TrimStart();
        var i = 0;
        while (i < text.Length && (char.IsLetter(text[i]) || text[i] == '_'))
        {
            i++;
        }

        return i > 0 && i < text.Length && text[i] == '(' ? text[..i] : null;
    }

    private static string InsideParens(string body)
    {
        var text = body.TrimStart();
        var open = text.IndexOf('(', StringComparison.Ordinal);
        return open >= 0 && text.EndsWith(')') ? text[(open + 1)..^1] : string.Empty;
    }

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
}
