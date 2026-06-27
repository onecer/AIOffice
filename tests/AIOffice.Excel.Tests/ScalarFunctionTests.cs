using System.Text.Json.Nodes;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.5) Golden-value tests for the scalar functions ClosedXML 0.105 returns
/// #NAME? for (XLOOKUP / IFS / SWITCH / LET / MAXIFS / MINIFS / AVERAGEIFS / AVERAGEIF):
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

    // ----- AVERAGEIF (singular twin of AVERAGEIFS) ------------------------------

    [Fact]
    public void Averageif_three_arg_text_criteria_uses_the_average_range()
    {
        var file = CreateWorkbook();
        // Rows whose B = "X" have A = 1, 3, 10 -> average 4.6667. The other rows
        // (B = "Y") are excluded by the text criteria. average_range is the 3rd arg.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(1, "X"),
                new JsonArray(2, "Y"),
                new JsonArray(3, "X"),
                new JsonArray(99, "Y"),
                new JsonArray(10, "X")))) ).IsOk);
        var env = EditOps(file, SetOp("/Sheet1/D1", ("value", "=AVERAGEIF(B1:B5,\"X\",A1:A5)")));
        Assert.True(env.IsOk, env.ToJson());
        Assert.Null(Json(env)["meta"]!["warnings"]); // evaluated, no formula_not_evaluated

        // Read the cached value back after save+reopen (Handler.Get reopens the
        // saved file; RawCell reads the raw <v> off disk).
        var d1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D1"))));
        Assert.Equal(4.6667, CachedNumber(d1), 4); // (1 + 3 + 10) / 3

        var (formula, cached, _) = RawCell(file, "Sheet1", "D1");
        Assert.Contains("AVERAGEIF", formula!, StringComparison.Ordinal);
        Assert.NotNull(cached);
        Assert.Equal(4.6667, double.Parse(cached!, System.Globalization.CultureInfo.InvariantCulture), 4);
        AssertValidatorClean(file);
    }

    [Theory]
    [InlineData(">5", 15.0)]   // {10,20}            -> 15
    [InlineData(">=10", 15.0)] // {10,20}            -> 15
    [InlineData("<=3", 2.0)]   // {1,2,3}            -> 2
    [InlineData("<3", 1.5)]    // {1,2}              -> 1.5
    [InlineData("<>10", 6.5)]  // {1,2,3,20}         -> 6.5
    [InlineData("=2", 2.0)]    // {2}                -> 2
    public void Averageif_two_arg_operator_criteria_range_doubles_as_average_range(
        string criteria, double expected)
    {
        var file = CreateWorkbook();
        // 2-arg form: the single range is both the criteria range AND the average
        // range. Source values 1, 2, 3, 10, 20.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(1), new JsonArray(2), new JsonArray(3),
                new JsonArray(10), new JsonArray(20)))) ).IsOk);
        Assert.True(EditOps(
            file, SetOp("/Sheet1/C1", ("value", $"=AVERAGEIF(A1:A5,\"{criteria}\")"))).IsOk);
        var c1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal(expected, CachedNumber(c1), 6);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Averageif_arg_order_pins_criteria_second_and_average_range_last()
    {
        var file = CreateWorkbook();
        // Designed so the 2-arg and 3-arg forms give DIFFERENT answers — if the
        // reshape mixed up the argument order the numbers would not separate.
        //   Column A (criteria range B / values): 1, 2, 3, 4
        //   Column B (average range, distinct):   100, 200, 300, 400
        // 3-arg AVERAGEIF(A1:A4, ">2", B1:B4): A>2 keeps rows 3,4 -> avg of B = (300+400)/2 = 350.
        // 2-arg AVERAGEIF(A1:A4, ">2"):        A>2 keeps rows 3,4 -> avg of A = (3+4)/2     = 3.5.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(1, 100),
                new JsonArray(2, 200),
                new JsonArray(3, 300),
                new JsonArray(4, 400)))) ).IsOk);
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/D1", ("value", "=AVERAGEIF(A1:A4,\">2\",B1:B4)")),
            SetOp("/Sheet1/D2", ("value", "=AVERAGEIF(A1:A4,\">2\")"))).IsOk);

        var d1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D1"))));
        var d2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D2"))));
        Assert.Equal(350.0, CachedNumber(d1)); // averages the 3rd-arg range
        Assert.Equal(3.5, CachedNumber(d2));   // averages the criteria range itself
        AssertValidatorClean(file);
    }

    [Fact]
    public void Averageif_no_match_caches_div_by_zero()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(1), new JsonArray(2), new JsonArray(3)))) ).IsOk);
        // No value is > 100 -> Excel returns #DIV/0! (matching AVERAGEIFS).
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=AVERAGEIF(A1:A3,\">100\")"))).IsOk);

        var (formula, _, _) = RawCell(file, "Sheet1", "C1");
        Assert.Contains("AVERAGEIF", formula!, StringComparison.Ordinal);
        var c1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        // The cached error surfaces as the #DIV/0! token in the read-back value.
        Assert.Equal("#DIV/0!", c1["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Averageif_cell_reference_criteria_resolves_via_criteria_text()
    {
        var file = CreateWorkbook();
        // Criteria is a cell reference C1 (= the text "x"); only rows whose B = "x"
        // count. average_range A has 10, 30 for those rows -> 20.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(10, "x"),
                new JsonArray(40, "y"),
                new JsonArray(30, "x")))),
            SetOp("/Sheet1/C1", ("value", "x"))).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/D1", ("value", "=AVERAGEIF(B1:B3,C1,A1:A3)"))).IsOk);

        var d1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D1"))));
        Assert.Equal(20.0, CachedNumber(d1)); // (10 + 30) / 2
        AssertValidatorClean(file);
    }

    [Fact]
    public void Averageif_malformed_arg_count_does_not_crash()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(1), new JsonArray(2)))) ).IsOk);
        // A single-argument AVERAGEIF is malformed; it must yield a sensible error,
        // never throw / corrupt the file.
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=AVERAGEIF(A1:A2)")));
        Assert.True(env.IsOk, env.ToJson());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Averageif_dispatch_change_is_inert_for_non_averageif_workbooks()
    {
        // A workbook with no AVERAGEIF cells must save byte-identically across an
        // idempotent re-save: adding AVERAGEIF to the dispatch set must not alter
        // the save path of unrelated formulas (here a natively-evaluated SUMIFS).
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(3), new JsonArray(7), new JsonArray(2)))),
            SetOp("/Sheet1/C1", ("value", "=SUMIFS(A1:A3,A1:A3,\">2\")"))).IsOk);
        var bytesBefore = File.ReadAllBytes(file);

        // Re-apply the identical SUMIFS formula to the same cell — an idempotent
        // edit. With no AVERAGEIF anywhere the save is byte-stable.
        Assert.True(EditOps(
            file, SetOp("/Sheet1/C1", ("value", "=SUMIFS(A1:A3,A1:A3,\">2\")"))).IsOk);
        Assert.Equal(bytesBefore, File.ReadAllBytes(file));
    }

    [Fact]
    public void Averageif_does_not_shadow_native_sumif_or_countif()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(3), new JsonArray(7), new JsonArray(2)))) ).IsOk);
        // SUMIF / COUNTIF are evaluated natively by ClosedXML 0.105; this change
        // must NOT add or shadow them — they still carry a real cached value and no
        // formula_not_evaluated warning.
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=SUMIF(A1:A3,\">2\")")),
            SetOp("/Sheet1/C2", ("value", "=COUNTIF(A1:A3,\">2\")")));
        Assert.True(env.IsOk, env.ToJson());
        Assert.Null(Json(env)["meta"]!["warnings"]);

        var c1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        var c2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C2"))));
        Assert.Equal(10.0, c1["value"]!.GetValue<double>()); // 3 + 7
        Assert.Equal(2.0, c2["value"]!.GetValue<double>());  // {3,7}
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
