using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M8 workbook diff (the flagship <see cref="IDiffer"/> slice). The current file
/// is <c>ctx.File</c>; the baseline is the OLD workbook. Every assertion pins the
/// semantic change list (per-cell value/formula, added/removed cells + sheets,
/// defined-name changes, table add/remove) and the deterministic sorted order
/// the contract requires.
/// </summary>
public sealed class DiffTests : ExcelTestBase
{
    /// <summary>Diffs <paramref name="currentFile"/> against <paramref name="baselineFile"/> via the IDiffer surface.</summary>
    private DiffResult Diff(string currentFile, string baselineFile) =>
        ((IDiffer)Handler).Diff(Ctx(currentFile), baselineFile);

    /// <summary>Seeds a workbook with a 3x3 grid of values + a formula in C3.</summary>
    private string SeedBaseline(string name)
    {
        var file = CreateWorkbook(name);
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:C3", ("values", new JsonArray(
            new JsonArray("Region", "Q1", "Q2"),
            new JsonArray("North", 100, 200),
            new JsonArray("South", 150, "=B3*2"))))).IsOk);
        return file;
    }

    [Fact]
    public void Identical_workbooks_diff_empty()
    {
        var baseline = SeedBaseline("old.xlsx");
        var current = SeedBaseline("new.xlsx");

        var result = Diff(current, baseline);
        Assert.Empty(result.Changes);
        Assert.Equal(0, result.Summary.Added);
        Assert.Equal(0, result.Summary.Removed);
        Assert.Equal(0, result.Summary.Modified);
        Assert.Equal(0, result.Summary.Moved);
        Assert.Null(result.Warnings);
    }

    [Fact]
    public void Modified_cell_value_reports_before_and_after()
    {
        var baseline = SeedBaseline("old.xlsx");
        var current = SeedBaseline("new.xlsx");
        Assert.True(EditOps(current, SetOp("/Sheet1/B2", ("value", 999))).IsOk);

        var result = Diff(current, baseline);
        var change = Assert.Single(result.Changes);
        Assert.Equal("modified", change.Kind);
        Assert.Equal("/Sheet1/B2", change.Path);
        Assert.Equal("100", change.Before);
        Assert.Equal("999", change.After);
        Assert.Equal(1, result.Summary.Modified);
    }

    [Fact]
    public void Modified_formula_reports_formula_text()
    {
        var baseline = SeedBaseline("old.xlsx");
        var current = SeedBaseline("new.xlsx");
        Assert.True(EditOps(current, SetOp("/Sheet1/C3", ("value", "=B3*3"))).IsOk);

        var result = Diff(current, baseline);
        var change = Assert.Single(result.Changes);
        Assert.Equal("modified", change.Kind);
        Assert.Equal("/Sheet1/C3", change.Path);
        Assert.Equal("=B3*2", change.Before);
        Assert.Equal("=B3*3", change.After);
    }

    [Fact]
    public void Added_cell_reports_only_after()
    {
        var baseline = SeedBaseline("old.xlsx");
        var current = SeedBaseline("new.xlsx");
        Assert.True(EditOps(current, SetOp("/Sheet1/A4", ("value", "East"))).IsOk);

        var result = Diff(current, baseline);
        var change = Assert.Single(result.Changes);
        Assert.Equal("added", change.Kind);
        Assert.Equal("/Sheet1/A4", change.Path);
        Assert.Equal("East", change.After);
        Assert.Null(change.Before);
        Assert.Equal(1, result.Summary.Added);
    }

    [Fact]
    public void Removed_cell_reports_only_before()
    {
        var baseline = SeedBaseline("old.xlsx");
        var current = SeedBaseline("new.xlsx");
        // Clear B2 in the current file (set blank).
        Assert.True(EditOps(current, RemoveOp("/Sheet1/B2")).IsOk);

        var result = Diff(current, baseline);
        var change = Assert.Single(result.Changes);
        Assert.Equal("removed", change.Kind);
        Assert.Equal("/Sheet1/B2", change.Path);
        Assert.Equal("100", change.Before);
        Assert.Null(change.After);
        Assert.Equal(1, result.Summary.Removed);
    }

    [Fact]
    public void Added_sheet_is_one_sheet_change()
    {
        var baseline = SeedBaseline("old.xlsx");
        var current = SeedBaseline("new.xlsx");
        Assert.True(EditOps(current, AddOp("/Summary", "sheet")).IsOk);

        var result = Diff(current, baseline);
        var change = Assert.Single(result.Changes);
        Assert.Equal("added", change.Kind);
        Assert.Equal("/Summary", change.Path);
        Assert.Equal("sheet", change.Detail);
    }

    [Fact]
    public void Removed_sheet_is_one_sheet_change()
    {
        var baseline = SeedBaseline("old.xlsx");
        Assert.True(EditOps(baseline, AddOp("/Extra", "sheet")).IsOk);
        var current = SeedBaseline("new.xlsx");

        var result = Diff(current, baseline);
        var change = Assert.Single(result.Changes);
        Assert.Equal("removed", change.Kind);
        Assert.Equal("/Extra", change.Path);
        Assert.Equal("sheet", change.Detail);
    }

    [Fact]
    public void Number_format_only_change_is_modified_with_format_detail()
    {
        var baseline = SeedBaseline("old.xlsx");
        var current = SeedBaseline("new.xlsx");
        Assert.True(EditOps(current, SetOp("/Sheet1/B2", ("numberFormat", "$#,##0.00"))).IsOk);

        var result = Diff(current, baseline);
        var change = Assert.Single(result.Changes);
        Assert.Equal("modified", change.Kind);
        Assert.Equal("/Sheet1/B2", change.Path);
        Assert.Equal("format", change.Detail);
        Assert.Equal("$#,##0.00", change.After);
    }

    [Fact]
    public void Defined_name_change_is_reported()
    {
        // The name's path IS its target range; baseline points at A1, current at B1.
        var baseline = SeedBaseline("old.xlsx");
        Assert.True(EditOps(baseline, AddOp("/Sheet1/A1", "name", ("name", "TaxRate"))).IsOk);
        var current = SeedBaseline("new.xlsx");
        Assert.True(EditOps(current, AddOp("/Sheet1/B1", "name", ("name", "TaxRate"))).IsOk);

        var result = Diff(current, baseline);
        var nameChange = Assert.Single(result.Changes, c => c.Detail == "definedName");
        Assert.Equal("modified", nameChange.Kind);
        Assert.Contains("TaxRate", nameChange.Path);
        Assert.NotNull(nameChange.Before);
        Assert.NotNull(nameChange.After);
    }

    [Fact]
    public void Builtin_print_area_name_is_not_diffed_as_a_defined_name()
    {
        var baseline = SeedBaseline("old.xlsx");
        var current = SeedBaseline("new.xlsx");
        // A print area is a builtin (_xlnm.Print_Area) surfaced via sheet props,
        // not a user defined name: it must NOT appear as a definedName change.
        Assert.True(EditOps(current, SetOp("/Sheet1", ("printArea", "A1:C3"))).IsOk);

        var result = Diff(current, baseline);
        Assert.DoesNotContain(result.Changes, c => c.Detail == "definedName");
    }

    [Fact]
    public void Added_table_is_reported()
    {
        var baseline = SeedBaseline("old.xlsx");
        var current = SeedBaseline("new.xlsx");
        Assert.True(EditOps(current, AddOp("/Sheet1/A1:C3", "table", ("name", "Sales"))).IsOk);

        var result = Diff(current, baseline);
        var tableChange = Assert.Single(result.Changes, c => c.Detail == "table");
        Assert.Equal("added", tableChange.Kind);
        Assert.Equal("/Sheet1/table[@name=Sales]", tableChange.Path);
    }

    [Fact]
    public void Changes_are_returned_in_deterministic_sorted_order()
    {
        var baseline = SeedBaseline("old.xlsx");
        var current = SeedBaseline("new.xlsx");
        // Mutate several cells out of path order so the sort is observable.
        Assert.True(EditOps(current,
            SetOp("/Sheet1/C2", ("value", 555)),
            SetOp("/Sheet1/A2", ("value", "Northwest")),
            SetOp("/Sheet1/B2", ("value", 111))).IsOk);

        var result = Diff(current, baseline);
        var paths = result.Changes.Select(c => c.Path).ToList();
        var sorted = paths.OrderBy(p => p, StringComparer.Ordinal).ToList();
        Assert.Equal(sorted, paths);

        // Re-running the same diff yields the identical ordering (platform-stable).
        var again = Diff(current, baseline).Changes.Select(c => c.Path).ToList();
        Assert.Equal(paths, again);
    }

    [Fact]
    public void Large_diff_is_truncated_with_a_warning()
    {
        var baseline = CreateWorkbook("old.xlsx");
        var current = CreateWorkbook("new.xlsx");

        // Seed a 600-row column in the current file only -> 600 added cells,
        // past the 500 cap.
        var rows = new JsonArray();
        for (var i = 0; i < 600; i++)
        {
            rows.Add(new JsonArray("v" + i.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }

        Assert.True(EditOps(current, SetOp("/Sheet1/A1:A600", ("values", rows))).IsOk);

        var result = Diff(current, baseline);
        Assert.Equal(ExcelDiff.MaxCellChanges, result.Changes.Count(c => c.Kind == "added"));
        Assert.NotNull(result.Warnings);
        Assert.Contains(result.Warnings!, w => w.Code == "diff_truncated");
    }

    [Fact]
    public void Diff_against_a_non_xlsx_baseline_is_invalid_args()
    {
        var current = SeedBaseline("new.xlsx");
        var docxBaseline = Path.Combine(Dir, "old.docx");
        File.WriteAllText(docxBaseline, "not really a docx");

        var exception = Assert.Throws<AiofficeException>(() => Diff(current, docxBaseline));
        Assert.Equal("invalid_args", exception.Code);
    }

    [Fact]
    public void Diff_against_a_missing_baseline_is_file_not_found()
    {
        var current = SeedBaseline("new.xlsx");
        var missing = Path.Combine(Dir, "ghost.xlsx");

        var exception = Assert.Throws<AiofficeException>(() => Diff(current, missing));
        Assert.Equal("file_not_found", exception.Code);
    }
}
