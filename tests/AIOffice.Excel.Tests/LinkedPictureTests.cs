using System.Linq;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// (1.7) The linked picture (Excel camera tool): a static snapshot of a cell range
/// embedded as a real PNG picture, anchored where asked, with the source/anchor in
/// a registry. Validator-clean, opens as a picture, get/remove round-trip, and the
/// add carries the honest <c>linked_picture_static</c> warning.
/// </summary>
public sealed class LinkedPictureTests : ExcelTestBase
{
    private string CreateDataWorkbook()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:D3", ("values", new JsonArray(
                new JsonArray("Region", "Q1", "Q2", "Q3"),
                new JsonArray("North", 10, 20, 30),
                new JsonArray("South", 15, 25, 35))))).IsOk);
        return file;
    }

    [Fact]
    public void Add_embeds_a_picture_anchored_at_the_target_and_warns_static()
    {
        var file = CreateDataWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "linkedPicture",
            ("sourceRange", "A1:D3"), ("anchor", "G2")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        var json = Json(envelope);
        var warning = Assert.Single(
            json["meta"]!["warnings"]!.AsArray(),
            w => w!["code"]!.GetValue<string>() == "linked_picture_static");
        Assert.Contains("snapshot", warning!["message"]!.GetValue<string>(), StringComparison.Ordinal);

        var op = json["data"]!["ops"]![0]!;
        Assert.Equal("linkedPicture", op["type"]!.GetValue<string>());
        Assert.Equal("/Sheet1/linkedPicture[1]", op["path"]!.GetValue<string>());
        Assert.Equal("A1:D3", op["sourceRange"]!.GetValue<string>());
        Assert.Equal("G2", op["anchor"]!.GetValue<string>());

        // A real PNG media part lands in the drawing.
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var hasPng = document.WorkbookPart!.WorksheetParts
            .Any(p => p.DrawingsPart?.ImageParts.Any(i => i.ContentType == "image/png") == true);
        Assert.True(hasPng);
    }

    [Fact]
    public void Get_reflects_source_range_and_anchor()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "linkedPicture",
            ("sourceRange", "A1:D3"), ("anchor", "G2"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/linkedPicture[1]"))));
        Assert.Equal("linkedPicture", data["kind"]!.GetValue<string>());
        Assert.Equal("A1:D3", data["sourceRange"]!.GetValue<string>());
        Assert.Equal("G2", data["anchor"]!.GetValue<string>());
        Assert.Contains("snapshot", data["note"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void Sheet_get_reports_the_linked_picture_count()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "linkedPicture",
            ("sourceRange", "A1:D3"), ("anchor", "G2"))).IsOk);

        var sheet = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Equal(1, sheet["linkedPictures"]!.GetValue<int>());
    }

    [Fact]
    public void Cross_sheet_source_range_is_qualified()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Summary", "sheet")).IsOk);

        var envelope = EditOps(file, AddOp("/Summary", "linkedPicture",
            ("sourceRange", "A1:D3"), ("anchor", "B2"), ("sheet", "Sheet1")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Summary/linkedPicture[1]"))));
        Assert.Equal("Sheet1!A1:D3", data["sourceRange"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_drops_the_linked_picture_and_its_part()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "linkedPicture",
            ("sourceRange", "A1:D3"), ("anchor", "G2"))).IsOk);

        Assert.True(EditOps(file, RemoveOp("/Sheet1/linkedPicture[1]")).IsOk);

        AssertValidatorClean(file);
        var sheet = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Null(sheet["linkedPictures"]);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var hasPng = document.WorkbookPart!.WorksheetParts
            .Any(p => p.DrawingsPart?.ImageParts.Any() == true);
        Assert.False(hasPng);
    }

    [Fact]
    public void Two_linked_pictures_index_one_based_per_sheet()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file,
            AddOp("/Sheet1", "linkedPicture", ("sourceRange", "A1:B2"), ("anchor", "F2")),
            AddOp("/Sheet1", "linkedPicture", ("sourceRange", "C1:D3"), ("anchor", "F10"))).IsOk);

        AssertValidatorClean(file);
        Assert.Equal(2, OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))))["linkedPictures"]!.GetValue<int>());
        Assert.True(Handler.Get(Ctx(file, ("path", "/Sheet1/linkedPicture[1]"))).IsOk);
        Assert.True(Handler.Get(Ctx(file, ("path", "/Sheet1/linkedPicture[2]"))).IsOk);
    }

    [Fact]
    public void Single_cell_source_range_is_accepted()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1", "linkedPicture",
            ("sourceRange", "A1"), ("anchor", "G2")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Survives_an_unrelated_edit_and_reopen()
    {
        var file = CreateDataWorkbook();
        Assert.True(EditOps(file, AddOp("/Sheet1", "linkedPicture",
            ("sourceRange", "A1:D3"), ("anchor", "G2"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/A10", ("value", "later"))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/linkedPicture[1]"))));
        Assert.Equal("A1:D3", data["sourceRange"]!.GetValue<string>());
    }

    // ----- validation ---------------------------------------------------------

    [Fact]
    public void Missing_source_range_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1", "linkedPicture", ("anchor", "G2")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("sourceRange", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Missing_anchor_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1", "linkedPicture", ("sourceRange", "A1:D3")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("anchor", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Bad_source_range_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1", "linkedPicture",
            ("sourceRange", "Sheet1!A1:D3"), ("anchor", "G2")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    [Fact]
    public void Unknown_prop_is_invalid_args_with_candidates()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1", "linkedPicture",
            ("sourceRange", "A1:D3"), ("anchor", "G2"), ("widthPx", 100)));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("sourceRange", envelope.Error.Candidates!);
    }

    [Fact]
    public void Unknown_source_sheet_is_invalid_path()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1", "linkedPicture",
            ("sourceRange", "A1:D3"), ("anchor", "G2"), ("sheet", "Nope")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
    }

    [Fact]
    public void Get_missing_linked_picture_is_invalid_path_with_candidates()
    {
        var file = CreateDataWorkbook();
        var envelope = Handler.Get(Ctx(file, ("path", "/Sheet1/linkedPicture[1]")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
    }

    [Fact]
    public void Add_on_a_cell_path_is_invalid_args()
    {
        var file = CreateDataWorkbook();
        var envelope = EditOps(file, AddOp("/Sheet1/A1", "linkedPicture",
            ("sourceRange", "A1:D3"), ("anchor", "G2")));
        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }
}
