using System.Globalization;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// (1.5) Goal Seek. <c>{op:set, path:/Sheet1/B1, props:{goalSeek:{targetCell:"B5",
/// targetValue:1000}}}</c> solves for the value of the changing cell B1 that makes
/// the formula cell B5 (which depends on B1) equal targetValue, then SETs B1 to
/// the found value and recalculates. The solve is Newton's method with a bisection
/// fallback over a bracketed range, evaluating B5 through ClosedXML's engine for
/// each probe (so any dependency chain is honored).
///
/// <para>On convergence the found input is persisted to the changing cell and the
/// op result reports the input plus the recomputed target. On no convergence the
/// changing cell is left UNCHANGED and a <c>goal_seek_no_solution</c> warning is
/// returned — a soft outcome, never a hard error.</para>
/// </summary>
internal static class ExcelGoalSeek
{
    private const double Tolerance = 1e-7;
    private const int MaxIterations = 100;

    /// <summary>
    /// The outcome of a goal-seek solve: whether it converged and (when it did) the
    /// found input. The achieved target is measured by the caller after persisting
    /// the input, so it is not carried here.
    /// </summary>
    public sealed record Result(bool Converged, double Input);

    /// <summary>
    /// Solves for the changing cell's value that drives <paramref name="targetCell"/>
    /// to <paramref name="targetValue"/>. The changing cell's original value is
    /// restored on no-convergence (the caller leaves it unchanged); on success the
    /// caller persists <see cref="Result.Input"/>. The workbook's live model is
    /// always left with the changing cell at its found/original value.
    /// </summary>
    public static Result Solve(
        XLWorkbook workbook, IXLCell changingCell, IXLCell targetCell, double targetValue)
    {
        var startValue = changingCell.TryGetValue<double>(out var sv) ? sv : 0d;
        var originalValue = changingCell.Value;

        double Residual(double x)
        {
            changingCell.Value = x;
            workbook.RecalculateAllFormulas();
            var v = EvaluateTarget(targetCell);
            return v - targetValue;
        }

        var result = Search(Residual, startValue);

        // Restore the changing cell + the model regardless of outcome. On success
        // the caller persists the found input; on failure the cell stays original.
        changingCell.Value = originalValue;
        workbook.RecalculateAllFormulas();
        return result;
    }

    /// <summary>Newton's method from the start with a bisection fallback; returns the outcome.</summary>
    private static Result Search(Func<double, double> residual, double startValue)
    {
        var x = startValue;
        for (var i = 0; i < MaxIterations; i++)
        {
            var f = residual(x);
            if (Math.Abs(f) < Tolerance)
            {
                return new Result(true, x);
            }

            var h = Step(x);
            var derivative = (residual(x + h) - residual(x - h)) / (2 * h);
            if (Math.Abs(derivative) < 1e-12)
            {
                break; // flat — Newton cannot move; fall to bisection
            }

            var next = x - (f / derivative);
            if (double.IsNaN(next) || double.IsInfinity(next))
            {
                break;
            }

            if (Math.Abs(next - x) < Tolerance)
            {
                return new Result(true, next);
            }

            x = next;
        }

        return TryBisect(residual, startValue, out var root)
            ? new Result(true, root)
            : new Result(false, startValue);
    }

    /// <summary>A finite-difference step scaled to the magnitude of x (never zero).</summary>
    private static double Step(double x) => Math.Max(Math.Abs(x), 1) * 1e-6;

    /// <summary>
    /// Bisection over an expanding symmetric bracket around the start: scan outward
    /// for a sign change in the residual, then halve in. Robust where Newton stalls
    /// (flat or oscillating derivatives).
    /// </summary>
    private static bool TryBisect(Func<double, double> residual, double start, out double root)
    {
        root = double.NaN;
        var center = start;
        var span = Math.Max(Math.Abs(start), 1);

        for (var expand = 0; expand < 60; expand++, span *= 2)
        {
            var low = center - span;
            var high = center + span;
            var fLow = residual(low);
            var fHigh = residual(high);
            if (double.IsNaN(fLow) || double.IsNaN(fHigh) || fLow * fHigh > 0)
            {
                continue; // no sign change in this bracket; widen
            }

            for (var i = 0; i < MaxIterations; i++)
            {
                var mid = (low + high) / 2;
                var fMid = residual(mid);
                if (Math.Abs(fMid) < Tolerance || (high - low) / 2 < Tolerance)
                {
                    root = mid;
                    return true;
                }

                if (fLow * fMid < 0)
                {
                    high = mid;
                }
                else
                {
                    low = mid;
                    fLow = fMid;
                }
            }

            root = (low + high) / 2;
            return true;
        }

        return false;
    }

    /// <summary>Evaluates the target cell's numeric value (booleans as 0/1); NaN on a non-numeric/error result.</summary>
    private static double EvaluateTarget(IXLCell cell)
    {
        try
        {
            var v = cell.Value;
            return v.Type switch
            {
                XLDataType.Number => v.GetNumber(),
                XLDataType.DateTime => v.GetDateTime().ToOADate(),
                XLDataType.Boolean => v.GetBoolean() ? 1 : 0,
                _ => double.NaN,
            };
        }
        catch (Exception)
        {
            return double.NaN;
        }
    }

    /// <summary>Formats a found input for the op-result detail (round-trips invariant).</summary>
    public static string Format(double value) => value.ToString("R", CultureInfo.InvariantCulture);
}
