using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.11) Golden-value tests for the statistical / ranking / lookup / reference
/// functions ClosedXML 0.105 returns #NAME? for: SMALL / RANK[.EQ] /
/// PERCENTILE[.INC] / QUARTILE[.INC] / CHOOSE / OFFSET / INDIRECT / AGGREGATE.
/// aioffice computes the cached value at write time so headless readers get a real
/// result and no formula_not_evaluated warning. LARGE (already native) and HLOOKUP
/// (already native) are verified to keep working. Out-of-range / error cases yield
/// the right Excel error (#NUM! / #VALUE! / #N/A / #REF!).
/// </summary>
public sealed class StatisticalFunctionTests : ExcelTestBase
{
    private static double CachedNumber(JsonNode cell) => cell["cachedValue"]!.GetValue<double>();

    /// <summary>Seeds the numeric column A1:A5 = {3,7,2,9,5} used by most tests.</summary>
    private string SeedColumn()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(3), new JsonArray(7), new JsonArray(2), new JsonArray(9), new JsonArray(5)))) ).IsOk);
        return file;
    }

    private static void AssertNoWarning(Envelope env)
    {
        Assert.True(env.IsOk, env.ToJson());
        Assert.Null(Json(env)["meta"]!["warnings"]); // evaluated, no formula_not_evaluated
    }

    private double GetCached(string file, string address)
    {
        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/" + address))));
        return CachedNumber(cell);
    }

    // ----- SMALL (the bug fix; LARGE already works) --------------------------

    [Fact]
    public void Small_returns_the_kth_smallest_value()
    {
        var file = SeedColumn();
        // {3,7,2,9,5} sorted = {2,3,5,7,9}; SMALL(.,2) = 3.
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=SMALL(A1:A5,2)")));
        AssertNoWarning(env);
        Assert.Equal(3.0, GetCached(file, "C1"));

        var (formula, cached, _) = RawCell(file, "Sheet1", "C1");
        Assert.Contains("SMALL", formula!, StringComparison.Ordinal);
        Assert.NotNull(cached);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Small_literal_array_matches_excel()
    {
        var file = CreateWorkbook();
        // SMALL({3,1,2},2) = 2 (the milestone's worked example).
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=SMALL({3,1,2},2)"))).IsOk);
        Assert.Equal(2.0, GetCached(file, "C1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Large_still_works_native()
    {
        var file = SeedColumn();
        // {3,7,2,9,5}; LARGE(.,1) = 9, LARGE(.,2) = 7 (ClosedXML evaluates natively).
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=LARGE(A1:A5,1)")),
            SetOp("/Sheet1/C2", ("value", "=LARGE(A1:A5,2)")));
        AssertNoWarning(env);
        Assert.Equal(9.0, GetCached(file, "C1"));
        Assert.Equal(7.0, GetCached(file, "C2"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Small_out_of_range_k_is_num_error()
    {
        var file = SeedColumn();
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=SMALL(A1:A5,9)"))).IsOk);
        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal("#NUM!", cell["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    // ----- RANK / RANK.EQ ----------------------------------------------------

    [Fact]
    public void Rank_descending_by_default()
    {
        var file = CreateWorkbook();
        // {10,20,30,40}; RANK(30,.,0) descending -> 40 is rank1, 30 is rank2.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(10), new JsonArray(20), new JsonArray(30), new JsonArray(40)))) ).IsOk);
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=RANK(30,A1:A4,0)")));
        AssertNoWarning(env);
        Assert.Equal(2.0, GetCached(file, "C1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Rank_ascending_when_order_nonzero()
    {
        var file = CreateWorkbook();
        // {10,20,30,40}; RANK(30,.,1) ascending -> 10 is rank1, 30 is rank3.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(10), new JsonArray(20), new JsonArray(30), new JsonArray(40)))) ).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=RANK(30,A1:A4,1)"))).IsOk);
        Assert.Equal(3.0, GetCached(file, "C1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Rank_eq_equals_rank_and_ties_share_top_rank()
    {
        var file = CreateWorkbook();
        // {40,40,20,10}; both 40s rank 1 (ties share top rank); 20 ranks 3, 10 ranks 4.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(40), new JsonArray(40), new JsonArray(20), new JsonArray(10)))) ).IsOk);
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=RANK.EQ(40,A1:A4)")),
            SetOp("/Sheet1/C2", ("value", "=RANK.EQ(20,A1:A4)")),
            SetOp("/Sheet1/C3", ("value", "=RANK(20,A1:A4)")));
        AssertNoWarning(env);
        Assert.Equal(1.0, GetCached(file, "C1")); // tie -> top rank
        Assert.Equal(3.0, GetCached(file, "C2")); // RANK.EQ
        Assert.Equal(3.0, GetCached(file, "C3")); // RANK == RANK.EQ
        AssertValidatorClean(file);
    }

    [Fact]
    public void Rank_of_absent_value_is_na()
    {
        var file = SeedColumn();
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=RANK(100,A1:A5,0)"))).IsOk);
        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal("#N/A", cell["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    // ----- PERCENTILE / QUARTILE ---------------------------------------------

    [Fact]
    public void Percentile_inc_interpolates()
    {
        var file = CreateWorkbook();
        // {1,2,3,4}; PERCENTILE.INC(.,0.9): pos = 0.9*3 = 2.7 -> 3 + 0.7*(4-3) = 3.7.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(1), new JsonArray(2), new JsonArray(3), new JsonArray(4)))) ).IsOk);
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=PERCENTILE.INC(A1:A4,0.9)")),
            SetOp("/Sheet1/C2", ("value", "=PERCENTILE(A1:A4,0.9)")));
        AssertNoWarning(env);
        Assert.Equal(3.7, GetCached(file, "C1"), 6);
        Assert.Equal(3.7, GetCached(file, "C2"), 6); // PERCENTILE == PERCENTILE.INC
        AssertValidatorClean(file);
    }

    [Fact]
    public void Percentile_endpoints_are_min_and_max()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(1), new JsonArray(2), new JsonArray(3), new JsonArray(4)))) ).IsOk);
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=PERCENTILE.INC(A1:A4,0)")),
            SetOp("/Sheet1/C2", ("value", "=PERCENTILE.INC(A1:A4,1)")),
            SetOp("/Sheet1/C3", ("value", "=PERCENTILE.INC(A1:A4,0.5)"))).IsOk);
        Assert.Equal(1.0, GetCached(file, "C1"));
        Assert.Equal(4.0, GetCached(file, "C2"));
        Assert.Equal(2.5, GetCached(file, "C3")); // median
        AssertValidatorClean(file);
    }

    [Fact]
    public void Quartile_inc_matches_percentile_quarters()
    {
        var file = CreateWorkbook();
        // {1,2,4,8}; QUARTILE.INC(.,1) = PERCENTILE.INC(.,0.25): pos = 0.25*3 = 0.75
        // -> 1 + 0.75*(2-1) = 1.75. q=0 -> min 1, q=4 -> max 8.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(1), new JsonArray(2), new JsonArray(4), new JsonArray(8)))) ).IsOk);
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=QUARTILE.INC(A1:A4,1)")),
            SetOp("/Sheet1/C2", ("value", "=QUARTILE(A1:A4,0)")),
            SetOp("/Sheet1/C3", ("value", "=QUARTILE.INC(A1:A4,4)")),
            SetOp("/Sheet1/C4", ("value", "=QUARTILE.INC(A1:A4,2)")));
        AssertNoWarning(env);
        Assert.Equal(1.75, GetCached(file, "C1"), 6);
        Assert.Equal(1.0, GetCached(file, "C2"));
        Assert.Equal(8.0, GetCached(file, "C3"));
        Assert.Equal(3.0, GetCached(file, "C4"), 6); // median of {1,2,4,8} = 3
        AssertValidatorClean(file);
    }

    [Fact]
    public void Quartile_out_of_range_is_num()
    {
        var file = SeedColumn();
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=QUARTILE.INC(A1:A5,5)"))).IsOk);
        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal("#NUM!", cell["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    // ----- CHOOSE ------------------------------------------------------------

    [Fact]
    public void Choose_picks_the_indexth_value()
    {
        var file = CreateWorkbook();
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=CHOOSE(2,10,20,30)")));
        AssertNoWarning(env);
        Assert.Equal(20.0, GetCached(file, "C1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Choose_resolves_cell_reference_arguments()
    {
        var file = SeedColumn();
        // CHOOSE(3, A1, A2, A3) -> A3 = 2.
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=CHOOSE(3,A1,A2,A3)"))).IsOk);
        Assert.Equal(2.0, GetCached(file, "C1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Choose_out_of_range_index_is_value_error()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=CHOOSE(5,10,20,30)"))).IsOk);
        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal("#VALUE!", cell["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    // ----- OFFSET / INDIRECT (reference functions) ---------------------------

    [Fact]
    public void Offset_returns_the_shifted_cell()
    {
        var file = SeedColumn();
        // A1 = 3; OFFSET(A1,2,0) -> A3 = 2.
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=OFFSET(A1,2,0)")));
        AssertNoWarning(env);
        Assert.Equal(2.0, GetCached(file, "C1"));

        var (formula, cached, _) = RawCell(file, "Sheet1", "C1");
        Assert.Contains("OFFSET", formula!, StringComparison.Ordinal);
        Assert.NotNull(cached);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Offset_with_column_shift()
    {
        var file = CreateWorkbook();
        // B2 = 42; OFFSET(A1,1,1) -> B2 = 42.
        Assert.True(EditOps(file, SetOp("/Sheet1/B2", ("value", 42))).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/D1", ("value", "=OFFSET(A1,1,1)"))).IsOk);
        Assert.Equal(42.0, GetCached(file, "D1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Offset_off_sheet_is_ref_error()
    {
        var file = SeedColumn();
        // Shifting A1 up by one row runs off the sheet -> #REF!.
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=OFFSET(A1,-1,0)"))).IsOk);
        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal("#REF!", cell["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Indirect_resolves_text_reference()
    {
        var file = SeedColumn();
        // A3 = 2; INDIRECT("A3") -> 2.
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=INDIRECT(\"A3\")")));
        AssertNoWarning(env);
        Assert.Equal(2.0, GetCached(file, "C1"));

        var (formula, cached, _) = RawCell(file, "Sheet1", "C1");
        Assert.Contains("INDIRECT", formula!, StringComparison.Ordinal);
        Assert.NotNull(cached);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Indirect_from_built_reference_in_a_cell()
    {
        var file = SeedColumn();
        // E1 holds the text "A4"; INDIRECT(E1) -> A4 = 9.
        Assert.True(EditOps(file, SetOp("/Sheet1/E1", ("value", "A4"))).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=INDIRECT(E1)"))).IsOk);
        Assert.Equal(9.0, GetCached(file, "C1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Indirect_unparsable_reference_is_ref_error()
    {
        var file = SeedColumn();
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=INDIRECT(\"not a ref\")"))).IsOk);
        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C1"))));
        Assert.Equal("#REF!", cell["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    // ----- AGGREGATE ---------------------------------------------------------

    [Fact]
    public void Aggregate_sum_and_average_ignoring_errors()
    {
        var file = SeedColumn();
        // {3,7,2,9,5}: SUM=26, AVERAGE=5.2. options 6 ignores errors (none here).
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=AGGREGATE(9,6,A1:A5)")),
            SetOp("/Sheet1/C2", ("value", "=AGGREGATE(1,6,A1:A5)")));
        AssertNoWarning(env);
        Assert.Equal(26.0, GetCached(file, "C1"));
        Assert.Equal(5.2, GetCached(file, "C2"), 6);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_skips_error_cells_in_the_range()
    {
        var file = SeedColumn();
        // Put a divide-by-zero error into A6; AGGREGATE(9,6,A1:A6) ignores it -> 26.
        Assert.True(EditOps(file, SetOp("/Sheet1/A6", ("value", "=1/0"))).IsOk);
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=AGGREGATE(9,6,A1:A6)")));
        AssertNoWarning(env);
        Assert.Equal(26.0, GetCached(file, "C1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_max_min_count_large_small_median()
    {
        var file = SeedColumn();
        // {3,7,2,9,5}: MAX=9, MIN=2, COUNT=5, LARGE k=2 -> 7, SMALL k=2 -> 3, MEDIAN=5.
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=AGGREGATE(4,6,A1:A5)")),
            SetOp("/Sheet1/C2", ("value", "=AGGREGATE(5,6,A1:A5)")),
            SetOp("/Sheet1/C3", ("value", "=AGGREGATE(2,6,A1:A5)")),
            SetOp("/Sheet1/C4", ("value", "=AGGREGATE(14,6,A1:A5,2)")),
            SetOp("/Sheet1/C5", ("value", "=AGGREGATE(15,6,A1:A5,2)")),
            SetOp("/Sheet1/C6", ("value", "=AGGREGATE(12,6,A1:A5)")));
        AssertNoWarning(env);
        Assert.Equal(9.0, GetCached(file, "C1"));
        Assert.Equal(2.0, GetCached(file, "C2"));
        Assert.Equal(5.0, GetCached(file, "C3"));
        Assert.Equal(7.0, GetCached(file, "C4")); // LARGE k=2
        Assert.Equal(3.0, GetCached(file, "C5")); // SMALL k=2
        Assert.Equal(5.0, GetCached(file, "C6")); // MEDIAN
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_percentile_and_quartile_inc()
    {
        var file = CreateWorkbook();
        // {1,2,3,4}; AGGREGATE(16,6,.,0.9) = 3.7; AGGREGATE(17,6,.,1) = 1.75.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(1), new JsonArray(2), new JsonArray(3), new JsonArray(4)))) ).IsOk);
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=AGGREGATE(16,6,A1:A4,0.9)")),
            SetOp("/Sheet1/C2", ("value", "=AGGREGATE(17,6,A1:A4,1)")));
        AssertNoWarning(env);
        Assert.Equal(3.7, GetCached(file, "C1"), 6);
        Assert.Equal(1.75, GetCached(file, "C2"), 6);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_exc_variant_evaluates()
    {
        var file = SeedColumn();
        // func_num 18 (PERCENTILE.EXC) now evaluates at write time. {3,7,2,9,5}
        // sorted = {2,3,5,7,9}, n=5; EXC interpolates at the 1-based position
        // k*(n+1): k=0.5 -> 3.0 -> sorted[3] = 5.0.
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=AGGREGATE(18,6,A1:A5,0.5)")));
        AssertNoWarning(env);
        Assert.Equal(5.0, GetCached(file, "C1"), 6);

        var (formula, cached, _) = RawCell(file, "Sheet1", "C1");
        Assert.Contains("AGGREGATE", formula!, StringComparison.Ordinal);
        Assert.NotNull(cached); // a real value is cached now
        AssertValidatorClean(file);
    }

    // ----- AGGREGATE 13 / 18 / 19 (1.16: MODE.SNGL / PERCENTILE.EXC / QUARTILE.EXC) -----

    /// <summary>Reads one cell back as its surfaced value (the error string for error cells).</summary>
    private string CellValueString(string file, string address) =>
        OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/" + address))))["value"]!.GetValue<string>();

    /// <summary>Seeds A1:A6 = {3,7,2,9,5,7} — has a repeated value (7) for MODE.SNGL.</summary>
    private string SeedRepeatColumn()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(3), new JsonArray(7), new JsonArray(2),
                new JsonArray(9), new JsonArray(5), new JsonArray(7)))) ).IsOk);
        return file;
    }

    [Fact]
    public void Aggregate_mode_sngl_returns_the_smallest_most_frequent()
    {
        var file = SeedRepeatColumn();
        // {3,7,2,9,5,7}; 7 is the only repeated value -> MODE.SNGL = 7.
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=AGGREGATE(13,6,A1:A6)")));
        AssertNoWarning(env);
        Assert.Equal(7.0, GetCached(file, "C1"), 6);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_mode_sngl_picks_smallest_among_equally_frequent()
    {
        var file = CreateWorkbook();
        // {4,4,2,2,9}: 4 and 2 both occur twice -> smallest (2) wins.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(4), new JsonArray(4), new JsonArray(2),
                new JsonArray(2), new JsonArray(9)))) ).IsOk);
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=AGGREGATE(13,6,A1:A5)")));
        AssertNoWarning(env);
        Assert.Equal(2.0, GetCached(file, "C1"), 6);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_mode_sngl_all_unique_is_na()
    {
        var file = SeedColumn(); // {3,7,2,9,5} — all distinct.
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=AGGREGATE(13,6,A1:A5)"))).IsOk);
        Assert.Equal("#N/A", CellValueString(file, "C1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_mode_sngl_empty_is_na()
    {
        var file = CreateWorkbook(); // B1:B5 has no numeric cells.
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=AGGREGATE(13,6,B1:B5)"))).IsOk);
        Assert.Equal("#N/A", CellValueString(file, "C1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_percentile_exc_interior_matches_excel()
    {
        var file = SeedColumn();
        // {3,7,2,9,5} sorted = {2,3,5,7,9}, n=5. EXC position = k*(n+1):
        //   k=0.25 -> 1.5 -> 2 + 0.5*(3-2) = 2.5
        //   k=0.75 -> 4.5 -> 7 + 0.5*(9-7) = 8.0
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=AGGREGATE(18,6,A1:A5,0.25)")),
            SetOp("/Sheet1/C2", ("value", "=AGGREGATE(18,6,A1:A5,0.75)")));
        AssertNoWarning(env);
        Assert.Equal(2.5, GetCached(file, "C1"), 6);
        Assert.Equal(8.0, GetCached(file, "C2"), 6);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_percentile_exc_out_of_open_interval_is_num_error()
    {
        var file = SeedColumn();
        // n=5; valid open interval is (1/6, 5/6) ~ (0.16667, 0.83333).
        // k=0.0, k=1.0, and just-outside both ends -> #NUM!.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=AGGREGATE(18,6,A1:A5,0)")),
            SetOp("/Sheet1/C2", ("value", "=AGGREGATE(18,6,A1:A5,1)")),
            SetOp("/Sheet1/C3", ("value", "=AGGREGATE(18,6,A1:A5,0.16)")),   // just below 1/6
            SetOp("/Sheet1/C4", ("value", "=AGGREGATE(18,6,A1:A5,0.84)"))).IsOk); // just above 5/6
        Assert.Equal("#NUM!", CellValueString(file, "C1"));
        Assert.Equal("#NUM!", CellValueString(file, "C2"));
        Assert.Equal("#NUM!", CellValueString(file, "C3"));
        Assert.Equal("#NUM!", CellValueString(file, "C4"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_percentile_exc_endpoints_are_inclusive_like_excel()
    {
        // The valid range endpoints 1/(n+1) and n/(n+1) are VALID in Excel (not #NUM!):
        // they resolve to the min and the max. n=3 makes them exactly representable
        // (1/4 and 3/4); sorted {2,5,9} -> min 2, max 9.
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(new JsonArray(5), new JsonArray(2), new JsonArray(9)))),
            SetOp("/Sheet1/C1", ("value", "=AGGREGATE(18,6,A1:A3,0.25)")),  // k = 1/(n+1) -> min
            SetOp("/Sheet1/C2", ("value", "=AGGREGATE(18,6,A1:A3,0.75)"))).IsOk); // k = n/(n+1) -> max
        Assert.Equal(2.0, GetCached(file, "C1"), 6);
        Assert.Equal(9.0, GetCached(file, "C2"), 6);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_percentile_exc_missing_k_is_value_error()
    {
        // PERCENTILE.EXC needs the 4th (k) arg; omitting it is #VALUE! (matches Excel).
        var file = SeedColumn();
        Assert.True(EditOps(file, SetOp("/Sheet1/C1", ("value", "=AGGREGATE(18,6,A1:A5)"))).IsOk);
        Assert.Equal("#VALUE!", CellValueString(file, "C1"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_quartile_exc_matches_percentile_exc()
    {
        var file = SeedColumn();
        // q=1/2/3 == PERCENTILE.EXC at .25/.5/.75 -> 2.5 / 5.0 / 8.0.
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=AGGREGATE(19,6,A1:A5,1)")),
            SetOp("/Sheet1/C2", ("value", "=AGGREGATE(19,6,A1:A5,2)")),
            SetOp("/Sheet1/C3", ("value", "=AGGREGATE(19,6,A1:A5,3)")));
        AssertNoWarning(env);
        Assert.Equal(2.5, GetCached(file, "C1"), 6);
        Assert.Equal(5.0, GetCached(file, "C2"), 6);
        Assert.Equal(8.0, GetCached(file, "C3"), 6);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_quartile_exc_boundary_quarters_are_num_error()
    {
        var file = SeedColumn();
        // EXC-specific guard: q=0 and q=4 are #NUM! (NOT valid like inclusive 17);
        // q=5 is out of {1,2,3} -> #NUM!.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=AGGREGATE(19,6,A1:A5,0)")),
            SetOp("/Sheet1/C2", ("value", "=AGGREGATE(19,6,A1:A5,4)")),
            SetOp("/Sheet1/C3", ("value", "=AGGREGATE(19,6,A1:A5,5)"))).IsOk);
        Assert.Equal("#NUM!", CellValueString(file, "C1"));
        Assert.Equal("#NUM!", CellValueString(file, "C2"));
        Assert.Equal("#NUM!", CellValueString(file, "C3"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_exc_options_propagate_or_skip_errors()
    {
        var file = SeedColumn();
        Assert.True(EditOps(file, SetOp("/Sheet1/A6", ("value", "=1/0"))).IsOk);
        // Options 0/1/4/5 do NOT ignore errors -> the #DIV/0! propagates for 13/18/19.
        // Options 2/3/6/7 skip it -> the func evaluates over {3,7,2,9,5}.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=AGGREGATE(13,4,A1:A6)")),     // MODE.SNGL, propagate
            SetOp("/Sheet1/C2", ("value", "=AGGREGATE(18,0,A1:A6,0.5)")), // PERCENTILE.EXC, propagate
            SetOp("/Sheet1/C3", ("value", "=AGGREGATE(19,5,A1:A6,2)")),   // QUARTILE.EXC, propagate
            SetOp("/Sheet1/C4", ("value", "=AGGREGATE(18,6,A1:A6,0.5)")), // skip -> 5.0
            SetOp("/Sheet1/C5", ("value", "=AGGREGATE(19,2,A1:A6,2)")),   // skip -> 5.0
            SetOp("/Sheet1/C6", ("value", "=AGGREGATE(13,7,A1:A6)"))).IsOk); // skip; all unique -> #N/A
        Assert.Equal("#DIV/0!", CellValueString(file, "C1"));
        Assert.Equal("#DIV/0!", CellValueString(file, "C2"));
        Assert.Equal("#DIV/0!", CellValueString(file, "C3"));
        Assert.Equal(5.0, GetCached(file, "C4"), 6);
        Assert.Equal(5.0, GetCached(file, "C5"), 6);
        Assert.Equal("#N/A", CellValueString(file, "C6")); // {3,7,2,9,5} after skip -> no repeats
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_func_num_20_stays_unevaluated()
    {
        var file = SeedColumn();
        // 20 is out of range -> the untouched default throws -> formula_not_evaluated.
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=AGGREGATE(20,6,A1:A5)")));
        Assert.True(env.IsOk, env.ToJson());
        var warnings = Json(env)["meta"]!["warnings"];
        Assert.NotNull(warnings);
        Assert.Contains(
            "formula_not_evaluated",
            warnings!.AsArray().Select(w => w!["code"]!.GetValue<string>()));

        var (formula, cached, _) = RawCell(file, "Sheet1", "C1");
        Assert.Contains("AGGREGATE", formula!, StringComparison.Ordinal);
        Assert.Null(cached); // no wrong value cached
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_options_that_do_not_ignore_errors_propagate_them()
    {
        var file = SeedColumn();
        // A6 = #DIV/0!. Options 0/1/4/5 do NOT ignore errors -> the error propagates;
        // option 6 still ignores it (-> 26). This is the AGGREGATE options fix.
        Assert.True(EditOps(file, SetOp("/Sheet1/A6", ("value", "=1/0"))).IsOk);
        var env = EditOps(
            file,
            SetOp("/Sheet1/C1", ("value", "=AGGREGATE(9,4,A1:A6)")),   // SUM, opt4: ignore nothing
            SetOp("/Sheet1/C2", ("value", "=AGGREGATE(1,0,A1:A6)")),   // AVERAGE, opt0
            SetOp("/Sheet1/C3", ("value", "=AGGREGATE(9,6,A1:A6)")));  // opt6: ignores -> 26
        Assert.True(env.IsOk, env.ToJson());
        var (_, c1, _) = RawCell(file, "Sheet1", "C1");
        var (_, c2, _) = RawCell(file, "Sheet1", "C2");
        Assert.Equal("#DIV/0!", c1);
        Assert.Equal("#DIV/0!", c2);
        Assert.Equal(26.0, GetCached(file, "C3"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Aggregate_invalid_options_is_value_error()
    {
        var file = SeedColumn();
        // options must be 0..7; 8 -> #VALUE!.
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=AGGREGATE(9,8,A1:A5)")));
        Assert.True(env.IsOk, env.ToJson());
        var (_, c1, _) = RawCell(file, "Sheet1", "C1");
        Assert.Equal("#VALUE!", c1);
        AssertValidatorClean(file);
    }

    // ----- HLOOKUP (already native; verify it keeps working) -----------------

    [Fact]
    public void Hlookup_exact_and_approximate_match_natively()
    {
        var file = CreateWorkbook();
        // First row keys 10/20/30/40 ; second row a/b/c/d ; third row 100/200/300/400.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(10, 20, 30, 40),
                new JsonArray("a", "b", "c", "d"),
                new JsonArray(100, 200, 300, 400)))) ).IsOk);

        var env = EditOps(
            file,
            // exact: lookup 30 in row1, return row2 -> "c".
            SetOp("/Sheet1/F1", ("value", "=HLOOKUP(30,A1:D3,2,FALSE)")),
            // exact: lookup 30 in row1, return row3 -> 300.
            SetOp("/Sheet1/F2", ("value", "=HLOOKUP(30,A1:D3,3,FALSE)")),
            // approximate: lookup 25 -> largest key <= 25 is 20 -> row3 = 200.
            SetOp("/Sheet1/F3", ("value", "=HLOOKUP(25,A1:D3,3,TRUE)")));
        AssertNoWarning(env);

        var f1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/F1"))));
        Assert.Equal("c", f1["cachedValue"]!.GetValue<string>());
        Assert.Equal(300.0, GetCached(file, "F2"));
        Assert.Equal(200.0, GetCached(file, "F3"));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Hlookup_exact_miss_is_na()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("values", new JsonArray(
                new JsonArray(10, 20, 30, 40),
                new JsonArray("a", "b", "c", "d")))) ).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/F1", ("value", "=HLOOKUP(25,A1:D2,2,FALSE)"))).IsOk);
        var f1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/F1"))));
        Assert.Equal("#N/A", f1["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    // ----- nested unevaluable function stays honest --------------------------

    [Fact]
    public void Choose_with_nested_unevaluable_stays_unevaluated()
    {
        var file = SeedColumn();
        // The chosen branch uses an unevaluable LAMBDA-bearing call; the cell must
        // honestly keep the formula_not_evaluated warning, never a wrong value.
        var env = EditOps(file, SetOp("/Sheet1/C1", ("value", "=CHOOSE(1,MAP(A1:A3,LAMBDA(x,x)),20)")));
        Assert.True(env.IsOk, env.ToJson());
        var warnings = Json(env)["meta"]!["warnings"];
        Assert.NotNull(warnings);
        Assert.Contains(
            "formula_not_evaluated",
            warnings!.AsArray().Select(w => w!["code"]!.GetValue<string>()));
        var (_, cached, _) = RawCell(file, "Sheet1", "C1");
        Assert.Null(cached);
        AssertValidatorClean(file);
    }
}
