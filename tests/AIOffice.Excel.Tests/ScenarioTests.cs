using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.5) Scenario Manager: {op:add, type:scenario, path:/Sheet1, props:{name,
/// cells, comment?}} saves a named scenario (changing cells + values) in the
/// worksheet's scenarios part; get/remove address /Sheet1/scenario[@name=…];
/// read --view structure lists scenarios; {op:set, props:{applyScenario:name}}
/// writes the stored values into the cells and recalculates. Validator-clean.
/// </summary>
public sealed class ScenarioTests : ExcelTestBase
{
    /// <summary>Seeds B1/B2 inputs and a B3=B1*B2 formula that a scenario will drive.</summary>
    private string SeedModel()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B1", ("value", 100)),
            SetOp("/Sheet1/B2", ("value", 0.10)),
            SetOp("/Sheet1/B3", ("value", "=B1*B2"))).IsOk);
        return file;
    }

    [Fact]
    public void Add_scenario_saves_the_changing_cells_and_values()
    {
        var file = SeedModel();
        var env = EditOps(
            file,
            AddOp("/Sheet1", "scenario",
                ("name", "Best Case"),
                ("cells", new JsonObject { ["B1"] = 120, ["B2"] = 0.15 }),
                ("comment", "optimistic")));
        Assert.True(env.IsOk, env.ToJson());
        Assert.Equal("scenario", OkData(env)["ops"]![0]!["type"]!.GetValue<string>());

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/scenario[@name=Best Case]"))));
        Assert.Equal("scenario", data["kind"]!.GetValue<string>());
        Assert.Equal("Best Case", data["name"]!.GetValue<string>());
        Assert.Equal("optimistic", data["comment"]!.GetValue<string>());
        Assert.Equal(120.0, data["cells"]!["B1"]!.GetValue<double>());
        Assert.Equal(0.15, data["cells"]!["B2"]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Structure_view_lists_scenarios()
    {
        var file = SeedModel();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "scenario", ("name", "Best Case"),
                ("cells", new JsonObject { ["B1"] = 120 })),
            AddOp("/Sheet1", "scenario", ("name", "Worst Case"),
                ("cells", new JsonObject { ["B1"] = 80 }))).IsOk);

        var structure = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        var scenarios = structure["sheets"]![0]!["scenarios"]!.AsArray();
        Assert.Equal(2, scenarios.Count);
        Assert.Contains(scenarios, s => s!["name"]!.GetValue<string>() == "Best Case");
        Assert.Contains(scenarios, s => s!["name"]!.GetValue<string>() == "Worst Case");
        var best = scenarios.First(s => s!["name"]!.GetValue<string>() == "Best Case");
        Assert.Contains("B1", best!["changingCells"]!.AsArray().Select(c => c!.GetValue<string>()));
    }

    [Fact]
    public void Apply_scenario_writes_the_values_and_recalculates()
    {
        var file = SeedModel();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "scenario", ("name", "Best Case"),
                ("cells", new JsonObject { ["B1"] = 200, ["B2"] = 0.20 }))).IsOk);

        // Apply it: B1 -> 200, B2 -> 0.20, and the dependent B3 = B1*B2 recomputes
        // to 40 with a real cached value (headless readers see the effect).
        var apply = EditOps(file, SetOp("/Sheet1", ("applyScenario", "Best Case")));
        Assert.True(apply.IsOk, apply.ToJson());

        var b1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B1"))));
        var b2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        var b3 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B3"))));
        Assert.Equal(200.0, b1["value"]!.GetValue<double>());
        Assert.Equal(0.20, b2["value"]!.GetValue<double>());
        Assert.Equal(40.0, b3["cachedValue"]!.GetValue<double>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Add_and_apply_scenario_in_one_batch()
    {
        var file = SeedModel();
        // The scenario is added and applied in a single edit (not yet on disk when
        // applied) — the apply reads the pending definition.
        var env = EditOps(
            file,
            AddOp("/Sheet1", "scenario", ("name", "Doubled"),
                ("cells", new JsonObject { ["B1"] = 100, ["B2"] = 0.50 })),
            SetOp("/Sheet1", ("applyScenario", "Doubled")));
        Assert.True(env.IsOk, env.ToJson());

        var b3 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B3"))));
        Assert.Equal(50.0, b3["cachedValue"]!.GetValue<double>()); // 100 * 0.50
        AssertValidatorClean(file);
    }

    [Fact]
    public void Remove_scenario_drops_it()
    {
        var file = SeedModel();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "scenario", ("name", "Best Case"),
                ("cells", new JsonObject { ["B1"] = 120 }))).IsOk);

        var removeEnv = EditOps(file, RemoveOp("/Sheet1/scenario[@name=Best Case]"));
        Assert.True(removeEnv.IsOk, removeEnv.ToJson());

        var get = Handler.Get(Ctx(file, ("path", "/Sheet1/scenario[@name=Best Case]")));
        Assert.False(get.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, get.Error!.Code);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Apply_unknown_scenario_is_rejected_with_a_suggestion()
    {
        var file = SeedModel();
        var env = EditOps(file, SetOp("/Sheet1", ("applyScenario", "Nonexistent")));
        Assert.False(env.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, env.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(env.Error.Suggestion));
    }

    [Fact]
    public void Add_scenario_with_a_formula_value_is_rejected()
    {
        var file = SeedModel();
        var env = EditOps(
            file,
            AddOp("/Sheet1", "scenario", ("name", "Bad"),
                ("cells", new JsonObject { ["B1"] = "=A1+1" })));
        Assert.False(env.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, env.Error!.Code);
    }

    [Fact]
    public void Add_scenario_without_a_name_is_rejected()
    {
        var file = SeedModel();
        var env = EditOps(file, AddOp("/Sheet1", "scenario", ("cells", new JsonObject { ["B1"] = 1 })));
        Assert.False(env.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, env.Error!.Code);
    }
}
