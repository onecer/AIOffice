using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M7 named cell styles slice: define a reusable style by name
/// (<c>/styles</c>), apply it to a cell/range by name, list/get/remove it, and
/// the payoff — the number format takes effect on the cached display, the named
/// style surfaces as a real OpenXml <c>cellStyle</c>, and everything survives a
/// reopen.
/// </summary>
public sealed class CellStyleTests : ExcelTestBase
{
    private string CreateStyledWorkbook()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B2", ("value", 1234.5))).IsOk);
        return file;
    }

    private static EditOp AddStyle(string name, params (string Key, JsonNode? Value)[] props)
    {
        var obj = new JsonObject { ["name"] = name };
        foreach (var (key, value) in props)
        {
            obj[key] = value;
        }

        return new EditOp { Op = "add", Path = "/styles", Type = "cellStyle", Props = obj };
    }

    [Fact]
    public void Create_named_style_then_get_returns_its_facets()
    {
        var file = CreateStyledWorkbook();

        var envelope = EditOps(file, AddStyle(
            "Currency-Red",
            ("numberFormat", "$#,##0.00;[Red]-$#,##0.00"),
            ("bold", true),
            ("color", "#CC0000")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal("/style[@name=Currency-Red]", detail["path"]!.GetValue<string>());
        Assert.Equal("cellStyle", detail["type"]!.GetValue<string>());
        AssertValidatorClean(file);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/style[@name=Currency-Red]"))));
        Assert.Equal("cellStyle", data["kind"]!.GetValue<string>());
        Assert.Equal("Currency-Red", data["name"]!.GetValue<string>());
        Assert.Equal("$#,##0.00;[Red]-$#,##0.00", data["numberFormat"]!.GetValue<string>());
        Assert.True(data["bold"]!.GetValue<bool>());
        Assert.Equal("#CC0000", data["color"]!.GetValue<string>());
    }

    [Fact]
    public void Apply_named_style_sets_the_number_format_on_the_cached_display()
    {
        var file = CreateStyledWorkbook();
        Assert.True(EditOps(file, AddStyle(
            "Currency-Red", ("numberFormat", "$#,##0.00;[Red]-$#,##0.00"), ("bold", true))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/B2", ("cellStyle", "Currency-Red")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Contains("cellStyle", envelope.ToJson(), StringComparison.Ordinal);
        AssertValidatorClean(file);

        // The number format took effect: get reports the formatted display text
        // and the bold facet rode along.
        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Equal("$#,##0.00;[Red]-$#,##0.00", cell["numberFormat"]!.GetValue<string>());
        Assert.True(cell["bold"]!.GetValue<bool>());
        Assert.Contains("1,234.50", cell["text"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_named_style_over_a_range_styles_every_cell()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/B2:B4", ("values", new JsonArray(
            new JsonArray(10), new JsonArray(20), new JsonArray(30))))).IsOk);
        Assert.True(EditOps(file, AddStyle("Pct", ("numberFormat", "0.0%"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/B2:B4", ("cellStyle", "Pct"))).IsOk);
        AssertValidatorClean(file);

        foreach (var address in new[] { "/Sheet1/B2", "/Sheet1/B3", "/Sheet1/B4" })
        {
            var cell = OkData(Handler.Get(Ctx(file, ("path", address))));
            Assert.Equal("0.0%", cell["numberFormat"]!.GetValue<string>());
        }
    }

    [Fact]
    public void Named_style_surfaces_as_a_real_cellStyle_and_survives_reopen()
    {
        var file = CreateStyledWorkbook();
        Assert.True(EditOps(file, AddStyle("Currency-Red", ("numberFormat", "$#,##0.00"))).IsOk);

        // Raw oracle: a cellStyle named exactly that exists in the styles part.
        using (var document = SpreadsheetDocument.Open(file, isEditable: false))
        {
            var styles = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            Assert.Contains(
                styles.CellStyles!.Elements<S.CellStyle>(),
                s => s.Name?.Value == "Currency-Red");
        }

        // The style survives a plain ClosedXML reopen/resave (registry property).
        var data = OkData(Handler.Read(Ctx(file, ("view", "styles"))));
        var styleList = data["styles"]!.AsArray();
        Assert.Single(styleList);
        Assert.Equal("Currency-Red", styleList[0]!["name"]!.GetValue<string>());
    }

    [Fact]
    public void List_styles_via_read_view_styles()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, AddStyle("Alpha", ("bold", true))).IsOk);
        Assert.True(EditOps(file, AddStyle("Beta", ("italic", true))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "styles"))));
        var names = data["styles"]!.AsArray().Select(s => s!["name"]!.GetValue<string>()).ToList();
        Assert.Equal(["Alpha", "Beta"], names);
    }

    [Fact]
    public void Remove_named_style_drops_it_from_the_workbook()
    {
        var file = CreateStyledWorkbook();
        Assert.True(EditOps(file, AddStyle("Doomed", ("bold", true))).IsOk);

        var envelope = EditOps(file, RemoveOp("/style[@name=Doomed]"));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("cellStyle", Json(envelope)["data"]!["ops"]![0]!["removed"]!.GetValue<string>());
        AssertValidatorClean(file);

        var data = OkData(Handler.Read(Ctx(file, ("view", "styles"))));
        Assert.Empty(data["styles"]!.AsArray());

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var styles = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
        Assert.DoesNotContain(
            styles.CellStyles?.Elements<S.CellStyle>() ?? [],
            s => s.Name?.Value == "Doomed");
    }

    [Fact]
    public void Duplicate_style_name_is_invalid_args()
    {
        var file = CreateStyledWorkbook();
        Assert.True(EditOps(file, AddStyle("Dup", ("bold", true))).IsOk);

        var envelope = EditOps(file, AddStyle("Dup", ("italic", true)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("already exists", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_unknown_style_is_invalid_path_with_candidates()
    {
        var file = CreateStyledWorkbook();
        Assert.True(EditOps(file, AddStyle("Currency-Red", ("numberFormat", "$#,##0.00"))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/B2", ("cellStyle", "Currency-Rd")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/style[@name=Currency-Red]", envelope.Error.Candidates!);
    }

    [Fact]
    public void Get_unknown_style_is_invalid_path()
    {
        var file = CreateStyledWorkbook();

        var envelope = Handler.Get(Ctx(file, ("path", "/style[@name=Nope]")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
    }

    [Fact]
    public void Bad_fill_color_in_a_style_is_invalid_args()
    {
        var file = CreateStyledWorkbook();

        var envelope = EditOps(file, AddStyle("Bad", ("fill", "not-a-color")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    [Fact]
    public void Set_on_a_style_path_is_rejected_with_a_recipe()
    {
        var file = CreateStyledWorkbook();
        Assert.True(EditOps(file, AddStyle("S", ("bold", true))).IsOk);

        var envelope = EditOps(file, SetOp("/style[@name=S]", ("bold", false)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("cellStyle", envelope.Error.Suggestion, StringComparison.Ordinal);
    }
}
