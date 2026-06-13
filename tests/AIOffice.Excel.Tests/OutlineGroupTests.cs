using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M6 outline grouping: group/ungroup row and column spans, collapse on create,
/// reflect outline levels in get + structure, and survive reopen — all
/// validator-clean.
/// </summary>
public sealed class OutlineGroupTests : ExcelTestBase
{
    private string SeedGrid(string name = "book.xlsx")
    {
        var file = CreateWorkbook(name);
        // 8 rows x 5 columns of data so grouped spans have content.
        var rows = new JsonArray();
        for (var r = 1; r <= 8; r++)
        {
            rows.Add(new JsonArray(r, r * 2, r * 3, r * 4, r * 5));
        }

        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("values", rows))).IsOk);
        return file;
    }

    // Regression: the CLI/MCP surface routes every op through EditOp.ParseBatch,
    // which validates each path with DocPath.Parse. The row/col span form must
    // survive that gate (it did not before M6 added the ElementSpan path kind),
    // otherwise grouping is reachable only by the tests that build EditOps
    // directly — never by a real user.
    [Theory]
    [InlineData("/Sheet1/row[2]:row[6]", "row")]
    [InlineData("/Sheet1/col[B]:col[E]", "col")]
    public void Group_span_paths_survive_the_ParseBatch_gate(string path, string axis)
    {
        var file = SeedGrid();
        var json = $"[{{\"op\":\"add\",\"path\":\"{path}\",\"type\":\"group\"}}]";

        var ops = EditOp.ParseBatch(json); // the exact gate the CLI/MCP use
        var envelope = Handler.Edit(Ctx(file), ops);

        Assert.True(envelope.IsOk, envelope.ToJson());
        var data = OkData(envelope)["ops"]![0]!;
        Assert.Equal("group", data["type"]!.GetValue<string>());
        Assert.Equal(axis, data["axis"]!.GetValue<string>());
        Assert.Equal(path, data["path"]!.GetValue<string>());
    }

    [Fact]
    public void Group_rows_sets_an_outline_level()
    {
        var file = SeedGrid();

        var envelope = EditOps(file, AddOp("/Sheet1/row[2]:row[6]", "group"));
        Assert.True(envelope.IsOk, envelope.ToJson());
        var data = OkData(envelope)["ops"]![0]!;
        Assert.Equal("group", data["type"]!.GetValue<string>());
        Assert.Equal("row", data["axis"]!.GetValue<string>());
        Assert.Equal("/Sheet1/row[2]:row[6]", data["path"]!.GetValue<string>());
        Assert.Equal(1, data["outlineLevel"]!.GetValue<int>());

        // The grouped rows carry outline level 1 in the package.
        using var workbook = new XLWorkbook(file);
        var sheet = workbook.Worksheet("Sheet1");
        for (var r = 2; r <= 6; r++)
        {
            Assert.Equal(1, sheet.Row(r).OutlineLevel);
        }

        Assert.Equal(0, sheet.Row(1).OutlineLevel); // outside the span
        Assert.Equal(0, sheet.Row(7).OutlineLevel);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Group_columns_sets_an_outline_level()
    {
        var file = SeedGrid();

        var envelope = EditOps(file, AddOp("/Sheet1/col[B]:col[E]", "group"));
        Assert.True(envelope.IsOk, envelope.ToJson());
        var data = OkData(envelope)["ops"]![0]!;
        Assert.Equal("col", data["axis"]!.GetValue<string>());
        Assert.Equal("/Sheet1/col[B]:col[E]", data["path"]!.GetValue<string>());

        using var workbook = new XLWorkbook(file);
        var sheet = workbook.Worksheet("Sheet1");
        for (var c = 2; c <= 5; c++)
        {
            Assert.Equal(1, sheet.Column(c).OutlineLevel);
        }

        Assert.Equal(0, sheet.Column(1).OutlineLevel);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Collapsed_group_hides_its_rows_and_survives_reopen()
    {
        var file = SeedGrid();

        var envelope = EditOps(file, AddOp("/Sheet1/row[2]:row[6]", "group", ("collapsed", true)));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.True(OkData(envelope)["ops"]![0]!["collapsed"]!.GetValue<bool>());

        // Reopen and verify: collapsed groups hide their inner rows in OOXML.
        using var workbook = new XLWorkbook(file);
        var sheet = workbook.Worksheet("Sheet1");
        for (var r = 2; r <= 6; r++)
        {
            Assert.True(sheet.Row(r).IsHidden, $"row {r} should be hidden in a collapsed group");
            Assert.Equal(1, sheet.Row(r).OutlineLevel);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Nested_groups_increase_the_outline_level()
    {
        var file = SeedGrid();

        Assert.True(EditOps(file, AddOp("/Sheet1/row[2]:row[7]", "group")).IsOk);
        Assert.True(EditOps(file, AddOp("/Sheet1/row[3]:row[5]", "group")).IsOk);

        using var workbook = new XLWorkbook(file);
        var sheet = workbook.Worksheet("Sheet1");
        Assert.Equal(1, sheet.Row(2).OutlineLevel);
        Assert.Equal(2, sheet.Row(3).OutlineLevel); // inner group
        Assert.Equal(2, sheet.Row(5).OutlineLevel);
        Assert.Equal(1, sheet.Row(7).OutlineLevel);
        AssertValidatorClean(file);
    }

    [Fact]
    public void Remove_ungroups_the_span()
    {
        var file = SeedGrid();
        Assert.True(EditOps(file, AddOp("/Sheet1/row[2]:row[6]", "group")).IsOk);

        var envelope = EditOps(file, RemoveOp("/Sheet1/row[2]:row[6]"));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("group", OkData(envelope)["ops"]![0]!["removed"]!.GetValue<string>());

        using var workbook = new XLWorkbook(file);
        var sheet = workbook.Worksheet("Sheet1");
        for (var r = 2; r <= 6; r++)
        {
            Assert.Equal(0, sheet.Row(r).OutlineLevel);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Get_row_reports_its_outline_level()
    {
        var file = SeedGrid();
        Assert.True(EditOps(file, AddOp("/Sheet1/row[2]:row[6]", "group")).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/row[3]"))));
        Assert.Equal(1, data["outlineLevel"]!.GetValue<int>());

        // A row outside the group has no outlineLevel field.
        var outside = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/row[1]"))));
        Assert.Null(outside["outlineLevel"]);
    }

    [Fact]
    public void Get_column_reports_its_outline_level()
    {
        var file = SeedGrid();
        Assert.True(EditOps(file, AddOp("/Sheet1/col[B]:col[E]", "group")).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/col[C]"))));
        Assert.Equal(1, data["outlineLevel"]!.GetValue<int>());
    }

    [Fact]
    public void Structure_view_reports_outline_groups()
    {
        var file = SeedGrid();
        Assert.True(EditOps(file, AddOp("/Sheet1/row[2]:row[6]", "group")).IsOk);
        Assert.True(EditOps(file, AddOp("/Sheet1/col[B]:col[E]", "group")).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        var outline = data["sheets"]![0]!["outline"]!;
        Assert.NotNull(outline);

        var rowGroups = outline["rowGroups"]!.AsArray();
        Assert.Single(rowGroups);
        Assert.Equal("/Sheet1/row[2]:row[6]", rowGroups[0]!["path"]!.GetValue<string>());
        Assert.Equal(1, rowGroups[0]!["outlineLevel"]!.GetValue<int>());

        var columnGroups = outline["columnGroups"]!.AsArray();
        Assert.Single(columnGroups);
        Assert.Equal("/Sheet1/col[B]:col[E]", columnGroups[0]!["path"]!.GetValue<string>());
    }

    [Fact]
    public void Structure_view_has_no_outline_block_without_grouping()
    {
        var file = SeedGrid();

        var data = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        Assert.Null(data["sheets"]![0]!["outline"]);
    }

    [Fact]
    public void Add_group_with_a_non_span_path_is_rejected()
    {
        var file = SeedGrid();

        var envelope = EditOps(file, AddOp("/Sheet1/A1", "group"));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("span", envelope.Error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Reversed_span_endpoints_normalize()
    {
        var file = SeedGrid();

        // row[6]:row[2] is accepted and normalized to row[2]:row[6].
        var envelope = EditOps(file, AddOp("/Sheet1/row[6]:row[2]", "group"));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("/Sheet1/row[2]:row[6]", OkData(envelope)["ops"]![0]!["path"]!.GetValue<string>());

        using var workbook = new XLWorkbook(file);
        Assert.Equal(1, workbook.Worksheet("Sheet1").Row(4).OutlineLevel);
    }

    [Fact]
    public void Group_on_a_quoted_sheet_name_resolves()
    {
        var file = CreateWorkbook("book.xlsx", "Q3 Data");
        Assert.True(EditOps(file, SetOp("/'Q3 Data'/A1", ("values", new JsonArray(
            new JsonArray(1), new JsonArray(2), new JsonArray(3), new JsonArray(4))))).IsOk);

        var envelope = EditOps(file, AddOp("/'Q3 Data'/row[2]:row[3]", "group"));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("/'Q3 Data'/row[2]:row[3]", OkData(envelope)["ops"]![0]!["path"]!.GetValue<string>());
        AssertValidatorClean(file);
    }
}
