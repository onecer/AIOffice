using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.5) Goal Seek: {op:set, path:/Sheet1/B1, props:{goalSeek:{targetCell:"B5",
/// targetValue:1000}}} solves for the changing cell's value that makes the formula
/// cell equal targetValue (Newton + bisection), SETs the changing cell and recalcs.
/// No convergence leaves the cell unchanged with a goal_seek_no_solution warning.
/// </summary>
public sealed class GoalSeekTests : ExcelTestBase
{
    [Fact]
    public void Solves_a_linear_target_b5_equals_b1_times_2()
    {
        var file = CreateWorkbook();
        // B5 = B1 * 2; seek B1 so that B5 = 100 -> B1 = 50.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B1", ("value", 1)),
            SetOp("/Sheet1/B5", ("value", "=B1*2"))).IsOk);

        var env = EditOps(
            file, SetOp("/Sheet1/B1", ("goalSeek", GoalSeek("B5", 100))));
        Assert.True(env.IsOk, env.ToJson());
        var op = OkData(env)["ops"]![0]!;
        Assert.True(op["goalSeek"]!["converged"]!.GetValue<bool>());
        Assert.Equal(50.0, op["goalSeek"]!["input"]!.GetValue<double>(), 4);

        // Persisted: B1 is now 50 and B5 recomputes to 100 (real cached value).
        var b1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        var b5 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B5"))));
        Assert.Equal(50.0, b1["value"]!.GetValue<double>(), 4);
        Assert.Equal(100.0, b5["cachedValue"]!.GetValue<double>(), 4);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Solves_a_nonlinear_target_b5_equals_b1_squared()
    {
        var file = CreateWorkbook();
        // B5 = B1^2; seek B1 so B5 = 144 -> B1 = 12 (positive root from a positive start).
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B1", ("value", 1)),
            SetOp("/Sheet1/B5", ("value", "=B1^2"))).IsOk);

        var env = EditOps(file, SetOp("/Sheet1/B1", ("goalSeek", GoalSeek("B5", 144))));
        Assert.True(env.IsOk, env.ToJson());
        var found = OkData(env)["ops"]![0]!["goalSeek"]!["input"]!.GetValue<double>();
        Assert.Equal(12.0, Math.Abs(found), 3);

        var b5 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B5"))));
        Assert.Equal(144.0, b5["cachedValue"]!.GetValue<double>(), 2);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Solves_through_a_multi_cell_dependency_chain()
    {
        var file = CreateWorkbook();
        // B1 -> B2 = B1 + 10 -> B5 = B2 * 3. Seek B1 so B5 = 90 -> B2 = 30 -> B1 = 20.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B1", ("value", 0)),
            SetOp("/Sheet1/B2", ("value", "=B1+10")),
            SetOp("/Sheet1/B5", ("value", "=B2*3"))).IsOk);

        var env = EditOps(file, SetOp("/Sheet1/B1", ("goalSeek", GoalSeek("B5", 90))));
        Assert.True(env.IsOk, env.ToJson());
        var b1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        Assert.Equal(20.0, b1["value"]!.GetValue<double>(), 3);
        AssertValidatorClean(file);
    }

    [Fact]
    public void No_solution_warns_and_leaves_the_cell_unchanged()
    {
        var file = CreateWorkbook();
        // B5 = B1*0 + 5 is constant 5; no B1 makes it equal 1000 -> no solution.
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B1", ("value", 7)),
            SetOp("/Sheet1/B5", ("value", "=B1*0+5"))).IsOk);

        var env = EditOps(file, SetOp("/Sheet1/B1", ("goalSeek", GoalSeek("B5", 1000))));
        Assert.True(env.IsOk, env.ToJson()); // soft outcome, not a hard error
        var warning = Assert.Single(env.Meta.Warnings!);
        Assert.Equal(WarningCodes.GoalSeekNoSolution, warning.Code);

        // B1 is left unchanged at its original value.
        var b1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        Assert.Equal(7.0, b1["value"]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Goal_seek_on_a_non_formula_target_is_rejected()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B1", ("value", 1)),
            SetOp("/Sheet1/B5", ("value", 42))).IsOk); // B5 is a constant, not a formula

        var env = EditOps(file, SetOp("/Sheet1/B1", ("goalSeek", GoalSeek("B5", 100))));
        Assert.False(env.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, env.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(env.Error.Suggestion));
    }

    private static System.Text.Json.Nodes.JsonObject GoalSeek(string targetCell, double targetValue) =>
        new() { ["targetCell"] = targetCell, ["targetValue"] = targetValue };
}
