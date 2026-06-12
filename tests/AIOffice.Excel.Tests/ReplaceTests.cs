using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M4 shared find/replace contract on xlsx: text-cell matching, scopes,
/// flags (regex/matchCase/wholeWord/inFormulas), per-op replacements +
/// locations, find_no_match warning, formula re-evaluation.
/// </summary>
public sealed class ReplaceTests : ExcelTestBase
{
    private static EditOp ReplaceOp(string path, params (string Key, JsonNode? Value)[] props)
    {
        var obj = new JsonObject();
        foreach (var (key, value) in props)
        {
            obj[key] = value;
        }

        return new EditOp { Op = "replace", Path = path, Props = obj };
    }

    [Fact]
    public void Literal_replace_counts_every_occurrence_and_lists_locations()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "beta beta")),
            SetOp("/Sheet1/B2", ("value", "the beta release")),
            SetOp("/Sheet1/C3", ("value", "BETA")), // matchCase defaults to false
            SetOp("/Sheet1/D4", ("value", "untouched"))).IsOk);

        var envelope = EditOps(file, ReplaceOp("/Sheet1", ("find", "beta"), ("replace", "gamma")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Null(envelope.Meta.Warnings); // matches found: no find_no_match
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal(4, detail["replacements"]!.GetValue<int>()); // A1 twice, B2, C3
        var locations = detail["locations"]!.AsArray().Select(n => n!.GetValue<string>()).ToList();
        Assert.Equal(["/Sheet1/A1", "/Sheet1/B2", "/Sheet1/C3"], locations);

        Assert.Equal("gamma gamma", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<string>());
        Assert.Equal("the gamma release", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))))["value"]!.GetValue<string>());
        Assert.Equal("gamma", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/C3"))))["value"]!.GetValue<string>());
        Assert.Equal("untouched", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D4"))))["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Match_case_and_whole_word_narrow_the_matches()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "Cat scattered cat"))).IsOk);

        var envelope = EditOps(file, ReplaceOp(
            "/Sheet1",
            ("find", "cat"), ("replace", "dog"), ("matchCase", true), ("wholeWord", true)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal(1, Json(envelope)["data"]!["ops"]![0]!["replacements"]!.GetValue<int>());
        Assert.Equal(
            "Cat scattered dog",
            OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Range_scope_leaves_cells_outside_it_alone()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "x")),
            SetOp("/Sheet1/B5", ("value", "x"))).IsOk);

        var envelope = EditOps(file, ReplaceOp("/Sheet1/A1:A4", ("find", "x"), ("replace", "y")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal(1, Json(envelope)["data"]!["ops"]![0]!["replacements"]!.GetValue<int>());
        Assert.Equal("y", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<string>());
        Assert.Equal("x", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B5"))))["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Regex_replace_supports_group_substitution()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "Order-123"))).IsOk);

        var envelope = EditOps(file, ReplaceOp(
            "/Sheet1", ("find", @"Order-(\d+)"), ("replace", "#$1"), ("regex", true)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("#123", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Literal_mode_treats_dollar_in_replacement_as_text()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "price"))).IsOk);

        Assert.True(EditOps(file, ReplaceOp("/Sheet1", ("find", "price"), ("replace", "$1 only"))).IsOk);

        Assert.Equal(
            "$1 only",
            OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<string>());
    }

    [Fact]
    public void Invalid_regex_is_invalid_args_with_suggestion()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "x"))).IsOk);

        var envelope = EditOps(file, ReplaceOp("/Sheet1", ("find", "("), ("regex", true)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("regular expression", envelope.Error.Message, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion));
    }

    [Fact]
    public void Zero_matches_is_ok_with_find_no_match_warning()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "hello"))).IsOk);

        var envelope = EditOps(file, ReplaceOp("/Sheet1", ("find", "nope"), ("replace", "x")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal(0, detail["replacements"]!.GetValue<int>());
        Assert.Empty(detail["locations"]!.AsArray());
        var warning = Assert.Single(envelope.Meta.Warnings!);
        Assert.Equal(ErrorCodes.FindNoMatch, warning.Code);
        Assert.Contains("nope", warning.Message, StringComparison.Ordinal);
        Assert.Equal("hello", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<string>());
    }

    [Fact]
    public void Formulas_are_skipped_unless_inFormulas_is_set()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1)),
            SetOp("/Sheet1/A2", ("value", 2)),
            SetOp("/Sheet1/A3", ("value", "=SUM(A1:A2)"))).IsOk);

        // Default: formula text is invisible to replace.
        var skipped = EditOps(file, ReplaceOp("/Sheet1", ("find", "SUM"), ("replace", "AVERAGE")));
        Assert.True(skipped.IsOk, skipped.ToJson());
        Assert.Equal(0, Json(skipped)["data"]!["ops"]![0]!["replacements"]!.GetValue<int>());
        Assert.Equal(ErrorCodes.FindNoMatch, Assert.Single(skipped.Meta.Warnings!).Code);
        Assert.Equal("SUM(A1:A2)", RawCell(file, "Sheet1", "A3").Formula);

        // inFormulas:true rewrites the formula and the replaced cell re-evaluates.
        var replaced = EditOps(file, ReplaceOp(
            "/Sheet1", ("find", "SUM"), ("replace", "AVERAGE"), ("inFormulas", true)));
        Assert.True(replaced.IsOk, replaced.ToJson());
        Assert.Equal(1, Json(replaced)["data"]!["ops"]![0]!["replacements"]!.GetValue<int>());

        var raw = RawCell(file, "Sheet1", "A3");
        Assert.Equal("AVERAGE(A1:A2)", raw.Formula);
        Assert.Equal("1.5", raw.CachedValue); // re-evaluated, not the stale SUM result
        AssertValidatorClean(file);
    }

    [Fact]
    public void Replacing_away_the_equals_sign_turns_the_cell_into_text()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1)),
            SetOp("/Sheet1/A3", ("value", "=SUM(A1:A1)"))).IsOk);

        var envelope = EditOps(file, ReplaceOp(
            "/Sheet1", ("find", "=SUM(A1:A1)"), ("replace", "gone"), ("inFormulas", true)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var a3 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A3"))));
        Assert.Equal("text", a3["type"]!.GetValue<string>());
        Assert.Equal("gone", a3["value"]!.GetValue<string>());
        Assert.Null(a3["formula"]);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Numbers_booleans_and_dates_never_match()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 42)),
            SetOp("/Sheet1/A2", ("value", true)),
            SetOp("/Sheet1/A3", ("value", "2024-05-01"))).IsOk);

        var envelope = EditOps(file, ReplaceOp("/Sheet1", ("find", "4"), ("replace", "9")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal(0, Json(envelope)["data"]!["ops"]![0]!["replacements"]!.GetValue<int>());
        Assert.Equal(ErrorCodes.FindNoMatch, Assert.Single(envelope.Meta.Warnings!).Code);
        Assert.Equal(42.0, OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<double>());
    }

    [Fact]
    public void Replacement_results_stay_literal_text()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "n/a"))).IsOk);

        Assert.True(EditOps(file, ReplaceOp("/Sheet1", ("find", "n/a"), ("replace", "123"))).IsOk);

        var a1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal("text", a1["type"]!.GetValue<string>()); // no re-auto-typing
        Assert.Equal("123", a1["value"]!.GetValue<string>());
    }

    [Fact]
    public void Missing_find_and_unknown_props_are_invalid_args()
    {
        var file = CreateWorkbook();

        var missing = EditOps(file, ReplaceOp("/Sheet1", ("replace", "x")));
        Assert.False(missing.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, missing.Error!.Code);
        Assert.Contains("find", missing.Error.Message, StringComparison.Ordinal);

        var unknown = EditOps(file, ReplaceOp("/Sheet1", ("find", "x"), ("caseMatch", true)));
        Assert.False(unknown.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, unknown.Error!.Code);
        Assert.Contains("matchCase", unknown.Error.Candidates!);
    }

    [Fact]
    public void Row_and_other_scopes_are_rejected_with_a_workaround()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "x"))).IsOk);

        var envelope = EditOps(file, ReplaceOp("/Sheet1/row[1]", ("find", "x"), ("replace", "y")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("range", envelope.Error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Single_cell_scope_works()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "alpha")),
            SetOp("/Sheet1/A2", ("value", "alpha"))).IsOk);

        var envelope = EditOps(file, ReplaceOp("/Sheet1/A2", ("find", "alpha"), ("replace", "beta")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal(1, Json(envelope)["data"]!["ops"]![0]!["replacements"]!.GetValue<int>());
        Assert.Equal("alpha", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))))["value"]!.GetValue<string>());
        Assert.Equal("beta", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A2"))))["value"]!.GetValue<string>());
    }
}
