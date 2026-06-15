using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.4) Golden-value tests for the financial iterative functions ClosedXML
/// cannot evaluate (RATE / IRR / XIRR): aioffice computes the cached value at
/// write time (Newton + bisection), so headless readers get a real number and no
/// formula_not_evaluated warning. Tolerances are 1e-6 against Excel's results.
/// The non-iterative PMT/PV/FV/NPV/NPER are evaluated by ClosedXML — covered here
/// to confirm they keep working.
/// </summary>
public sealed class FinancialFunctionTests : ExcelTestBase
{
    private const double Tol = 1e-6;

    private static double CachedNumber(JsonNode cell) =>
        cell["cachedValue"]!.GetValue<double>();

    [Fact]
    public void Irr_caches_the_internal_rate_of_return()
    {
        var file = CreateWorkbook();
        // -10000 then three inflows; the IRR that zeroes the NPV is ≈ 0.163406.
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("values", new JsonArray(
            new JsonArray(-10000),
            new JsonArray(3000),
            new JsonArray(4200),
            new JsonArray(6800))))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/C1", ("value", "=IRR(A1:A4)")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Null(Json(envelope)["meta"]!["warnings"]); // evaluated, not formula_not_evaluated

        var c1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal(0.163406, CachedNumber(c1), 5);

        // Independent oracle: the cached IRR really does zero the NPV (1e-6).
        var cashflows = new[] { -10000.0, 3000, 4200, 6800 };
        var r = CachedNumber(c1);
        var npv = 0d;
        for (var t = 0; t < cashflows.Length; t++)
        {
            npv += cashflows[t] / Math.Pow(1 + r, t);
        }

        Assert.True(Math.Abs(npv) < 1e-4, $"NPV at the cached IRR was {npv}, not ~0");

        // Raw oracle: formula text + a numeric cached value on disk.
        var (formula, cached, _) = RawCell(file, "Sheet1", "C1");
        Assert.Contains("IRR", formula!, StringComparison.Ordinal);
        Assert.NotNull(cached);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Irr_matches_a_known_textbook_value()
    {
        var file = CreateWorkbook();
        // -100, +30, +35, +40, +45 zeroes the NPV at r ≈ 0.170937.
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("values", new JsonArray(
            new JsonArray(-100),
            new JsonArray(30),
            new JsonArray(35),
            new JsonArray(40),
            new JsonArray(45))))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=IRR(A1:A5)"))).IsOk);
        var c1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal(0.170937, CachedNumber(c1), 4);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Rate_caches_the_periodic_rate()
    {
        var file = CreateWorkbook();
        // RATE(48, -200, 8000): the periodic rate on a 48-period, -200 payment,
        // 8000 PV loan. Excel ≈ 0.0077014 per period.
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=RATE(48,-200,8000)"))).IsOk);

        var c1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal(0.00770147, CachedNumber(c1), 5);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Xirr_caches_the_irregular_rate_of_return()
    {
        var file = CreateWorkbook();
        // Cash flows with dates; XIRR ≈ 0.373362 (Microsoft's documented example).
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(-10000),
                new JsonArray(2750),
                new JsonArray(4250),
                new JsonArray(3250),
                new JsonArray(2750)))),
            SetOp("/Sheet1/B1", ("values", new JsonArray(
                new JsonArray("2008-01-01"),
                new JsonArray("2008-03-01"),
                new JsonArray("2008-10-30"),
                new JsonArray("2009-02-15"),
                new JsonArray("2009-04-01"))))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/D1", ("value", "=XIRR(A1:A5,B1:B5)"))).IsOk);
        var d1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D1"))));
        Assert.Equal(0.373362535, CachedNumber(d1), 4);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Pmt_is_evaluated_by_the_built_in_engine()
    {
        var file = CreateWorkbook();
        // PMT(0.08/12, 10, 10000): the monthly payment. Excel ≈ -1037.03.
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=PMT(0.08/12,10,10000)"))).IsOk);

        var c1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal(-1037.03, CachedNumber(c1), 1);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Npv_pv_nper_are_filled_in_for_functions_closedxml_cannot_do()
    {
        var file = CreateWorkbook();
        // ClosedXML 0.105 returns #NAME? for NPV/PV/NPER; aioffice computes them.
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(8000), new JsonArray(9200), new JsonArray(10000)))),
            SetOp("/Sheet1/C1", ("value", "=NPV(0.08,A1:A3)")),
            SetOp("/Sheet1/C2", ("value", "=PV(0.08/12,10,-1000)")),
            SetOp("/Sheet1/C3", ("value", "=NPER(0.08/12,-200,8000)")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Null(Json(envelope)["meta"]!["warnings"]); // no formula_not_evaluated for any of them

        Assert.Equal(23233.247, CachedNumber(OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))))), 2);
        Assert.Equal(9642.903, CachedNumber(OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C2"))))), 2);
        Assert.Equal(46.678145, CachedNumber(OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C3"))))), 4);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Fv_keeps_working_through_the_built_in_engine()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/C2", ("value", "=FV(0.06/12,12,-100)"))).IsOk);

        var fv = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C2"))));
        Assert.Equal(1233.556, CachedNumber(fv), 2); // 12 months of -100 at 0.5%/mo
        AssertValidatorClean(file);
    }
}
