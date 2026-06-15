using System.Text.Json.Nodes;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.5) Golden-value tests for the scalar functions ClosedXML 0.105 returns
/// #NAME? for (XLOOKUP / IFS / SWITCH / LET / MAXIFS / MINIFS / AVERAGEIFS):
/// aioffice computes the cached value at write time so headless readers get a
/// real result and no formula_not_evaluated warning. Functions ClosedXML already
/// evaluates (TEXTJOIN / CONCAT / IFERROR / SUMIFS / COUNTIFS) are verified to
/// keep working. LAMBDA stays honestly unevaluated.
/// </summary>
public sealed class ScalarFunctionTests : ExcelTestBase
{
    private static double CachedNumber(JsonNode cell) => cell["cachedValue"]!.GetValue<double>();

    private static string CachedText(JsonNode cell) => cell["cachedValue"]!.GetValue<string>();

    /// <summary>Seeds a small numeric column A1:A5 plus a parallel text column D1:D5 / E1:E5 lookup table.</summary>
    private string SeedLookupTable()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(3), new JsonArray(7), new JsonArray(2), new JsonArray(9), new JsonArray(5)))),
            SetOp("/Sheet1/D1", ("values", new JsonArray(
                new JsonArray("apple", 10),
                new JsonArray("banana", 20),
                new JsonArray("cherry", 30)))) ).IsOk);
        return file;
    }

    [Fact]
    public void Xlookup_exact_match_caches_the_return_value()
    {
        var file = SeedLookupTable();
        // Look up "banana" in D1:D3, return the parallel E1:E3 -> 20.
        var env = EditOps(file, SetOp("/Sheet1/G1", ("value", "=XLOOKUP(\"banana\",D1:D3,E1:E3)")));
        Assert.True(env.IsOk, env.ToJson());
        Assert.Null(Json(env)["meta"]!["warnings"]); // evaluated, no formula_not_evaluated

        var g1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/G1"))));
        Assert.Equal(20.0, CachedNumber(g1));

        var (formula, cached, _) = RawCell(file, "Sheet1", "G1");
        Assert.Contains("XLOOKUP", formula!, StringComparison.Ordinal);
        Assert.NotNull(cached);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Xlookup_not_found_returns_the_if_not_found_argument()
    {
        var file = SeedLookupTable();
        Assert.True(EditOps(
            file, SetOp("/Sheet1/G1", ("value", "=XLOOKUP(\"durian\",D1:D3,E1:E3,\"missing\")"))).IsOk);
        var g1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/G1"))));
        Assert.Equal("missing", CachedText(g1));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Xlookup_approximate_next_smaller_match()
    {
        var file = CreateWorkbook();
        // Tiered table: amounts 0/100/500 map to rates 0.0/0.1/0.2. Look up 250 with
        // match_mode -1 (next smaller) -> the 100 tier -> 0.1.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(0, 0.0), new JsonArray(100, 0.1), new JsonArray(500, 0.2)))) ).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/D1", ("value", "=XLOOKUP(250,A1:A3,B1:B3,,-1)"))).IsOk);

        var d1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D1"))));
        Assert.Equal(0.1, CachedNumber(d1), 6);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Ifs_selects_the_first_truthy_branch()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", 7))).IsOk);
        // 7 is >5 -> "big".
        Assert.True(EditOps(
            file, SetOp("/Sheet1/B1", ("value", "=IFS(A1>10,\"huge\",A1>5,\"big\",TRUE,\"small\")"))).IsOk);
        var b1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        Assert.Equal("big", CachedText(b1));

        // Change A1 to 3 in a fresh cell: 3 is neither >10 nor >5 -> TRUE -> "small".
        Assert.True(EditOps(file, SetOp("/Sheet1/A2", ("value", 3))).IsOk);
        Assert.True(EditOps(
            file, SetOp("/Sheet1/B2", ("value", "=IFS(A2>10,\"huge\",A2>5,\"big\",TRUE,\"small\")"))).IsOk);
        var b2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Equal("small", CachedText(b2));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Switch_matches_a_case_and_falls_back_to_default()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", 2))).IsOk);
        Assert.True(EditOps(
            file, SetOp("/Sheet1/B1", ("value", "=SWITCH(A1,1,\"one\",2,\"two\",3,\"three\",\"other\")"))).IsOk);
        var b1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        Assert.Equal("two", CachedText(b1));

        // 9 matches no case -> the trailing default "other".
        Assert.True(EditOps(file, SetOp("/Sheet1/A2", ("value", 9))).IsOk);
        Assert.True(EditOps(
            file, SetOp("/Sheet1/B2", ("value", "=SWITCH(A2,1,\"one\",2,\"two\",\"other\")"))).IsOk);
        var b2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Equal("other", CachedText(b2));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Let_binds_names_sequentially_then_evaluates_the_calc()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", 5))).IsOk);
        // x = A1 (5); y = x * 2 (10); result = x + y = 15.
        Assert.True(EditOps(file, SetOp("/Sheet1/B1", ("value", "=LET(x,A1,y,x*2,x+y)"))).IsOk);
        var b1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        Assert.Equal(15.0, CachedNumber(b1));

        var (formula, cached, _) = RawCell(file, "Sheet1", "B1");
        Assert.Contains("LET", formula!, StringComparison.Ordinal);
        Assert.NotNull(cached);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Maxifs_and_minifs_aggregate_conditionally()
    {
        var file = CreateWorkbook();
        // Values A, criteria column B. Find MAX/MIN of A where B = "x".
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(10, "x"), new JsonArray(40, "y"), new JsonArray(30, "x"), new JsonArray(5, "x")))) ).IsOk);
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/D1", ("value", "=MAXIFS(A1:A4,B1:B4,\"x\")")),
            SetOp("/Sheet1/D2", ("value", "=MINIFS(A1:A4,B1:B4,\"x\")"))).IsOk);

        var d1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D1"))));
        var d2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D2"))));
        Assert.Equal(30.0, CachedNumber(d1)); // max of {10,30,5}
        Assert.Equal(5.0, CachedNumber(d2));  // min of {10,30,5}
        AssertValidatorClean(file);
    }

    [Fact]
    public void Averageifs_with_a_numeric_comparison_criteria()
    {
        var file = CreateWorkbook();
        // Average A where A > 1: {3,7,2,9,5} all > 1 except none excluded -> 5.2.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(3), new JsonArray(7), new JsonArray(2), new JsonArray(9), new JsonArray(5)))) ).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=AVERAGEIFS(A1:A5,A1:A5,\">2\")"))).IsOk);
        var c1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        // > 2 keeps {3,7,9,5} -> average 6.
        Assert.Equal(6.0, CachedNumber(c1));
        AssertValidatorClean(file);
    }

    // ----- functions ClosedXML already evaluates (verify they keep working) ----

    [Fact]
    public void Textjoin_and_concat_evaluate_without_warning()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray("a"), new JsonArray("b"), new JsonArray("c")))) ).IsOk);

        var join = EditOps(file, SetOp("/Sheet1/C1", ("value", "=TEXTJOIN(\",\",TRUE,A1:A3)")));
        Assert.True(join.IsOk, join.ToJson());
        Assert.Null(Json(join)["meta"]!["warnings"]);
        var c1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal("a,b,c", c1["value"]!.GetValue<string>());

        Assert.True(EditOps(file, SetOp("/Sheet1/C2", ("value", "=CONCAT(A1:A3)"))).IsOk);
        var c2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C2"))));
        Assert.Equal("abc", c2["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Iferror_sumifs_countifs_evaluate_without_warning()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(3), new JsonArray(7), new JsonArray(2)))) ).IsOk);

        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=IFERROR(1/0,\"err\")")),
            SetOp("/Sheet1/C2", ("value", "=SUMIFS(A1:A3,A1:A3,\">2\")")),
            SetOp("/Sheet1/C3", ("value", "=COUNTIFS(A1:A3,\">2\")")));
        Assert.True(env.IsOk, env.ToJson());
        Assert.Null(Json(env)["meta"]!["warnings"]);

        var c1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        var c2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C2"))));
        var c3 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C3"))));
        Assert.Equal("err", c1["value"]!.GetValue<string>());
        Assert.Equal(10.0, c2["value"]!.GetValue<double>()); // 3 + 7
        Assert.Equal(2.0, c3["value"]!.GetValue<double>());  // {3,7}
        AssertValidatorClean(file);
    }

    [Fact]
    public void Lambda_stays_honestly_unevaluated()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(1), new JsonArray(2), new JsonArray(3)))) ).IsOk);

        // A stored LAMBDA passed to MAP cannot be evaluated by the built-in engine;
        // it must honestly keep the formula_not_evaluated warning (never a wrong
        // value) — correctness over coverage.
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=MAP(A1:A3,LAMBDA(x,x*2))")));
        Assert.True(env.IsOk, env.ToJson());
        var warnings = Json(env)["meta"]!["warnings"];
        Assert.NotNull(warnings);
        Assert.Contains(
            "formula_not_evaluated",
            warnings!.AsArray().Select(w => w!["code"]!.GetValue<string>()));

        // The cell carries the formula text but no cached value (Excel computes on open).
        var (formula, cached, _) = RawCell(file, "Sheet1", "C1");
        Assert.Contains("LAMBDA", formula!, StringComparison.Ordinal);
        Assert.Null(cached);
        AssertValidatorClean(file);
    }
}
