using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M2 image slice: PNG + JPEG embedding via ClosedXML, header-sniffed,
/// with the src path forced through the Core Workspace sandbox.
/// </summary>
public sealed class ImageTests : ExcelTestBase
{
    /// <summary>A 4x2 red PNG (so aspect-ratio math is observable).</summary>
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAEElEQVR4nGP4z8AARwzIHABvqgf5gNwAKAAAAABJRU5ErkJggg==");

    /// <summary>A 4x2 JPEG.</summary>
    private static readonly byte[] JpegBytes = Convert.FromBase64String(
        "/9j/4AAQSkZJRgABAQAASABIAAD/4QBMRXhpZgAATU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKAC" +
        "AAQAAAABAAAABKADAAQAAAABAAAAAgAAAAD/7QA4UGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAAAAAABDUHYzZ" +
        "jwCyBOmACZjs+EJ+/8AAEQgAAgAEAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIB" +
        "AwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNE" +
        "RUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfI" +
        "ycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIB" +
        "AgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpD" +
        "REVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXG" +
        "x8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/bAEMAAgICAgICAwICAwUDAwMFBgUFBQUGCAYGBgYGCAoICAgICAgK" +
        "CgoKCgoKCgwMDAwMDA4ODg4ODw8PDw8PDw8PD//bAEMBAgICBAQEBwQEBxALCQsQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQ" +
        "EBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEP/dAAQAAf/aAAwDAQACEQMRAD8A+L6KKK/lM/38P//Z");

    private string WritePng(string name = "dot.png")
    {
        var path = Path.Combine(Dir, name);
        File.WriteAllBytes(path, PngBytes);
        return name; // relative to the workspace root — the sandbox resolves it
    }

    /// <summary>The first sheet's drawings part and its picture anchors.</summary>
    private static (WorksheetPart Sheet, List<Xdr.Picture> Pictures) RawPictures(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>().First();
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var pictures = worksheetPart.DrawingsPart?.WorksheetDrawing?
            .Descendants<Xdr.Picture>().ToList() ?? [];
        return (worksheetPart, pictures);
    }

    [Fact]
    public void Add_png_embeds_part_at_anchor_validator_clean()
    {
        var file = CreateWorkbook();
        var src = WritePng();

        var envelope = EditOps(file, AddOp("/Sheet1", "image",
            ("src", src), ("anchor", "E2"), ("name", "Logo")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal("/Sheet1/image[1]", detail["path"]!.GetValue<string>());
        Assert.Equal("png", detail["format"]!.GetValue<string>());
        Assert.Equal("E2", detail["anchor"]!.GetValue<string>());

        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (worksheetPart, pictures) = RawPictures(document);
        var picture = Assert.Single(pictures);

        // The image part holds exactly the source bytes (png content type).
        var imagePart = Assert.Single(worksheetPart.DrawingsPart!.ImageParts);
        Assert.Equal("image/png", imagePart.ContentType);
        using var stream = imagePart.GetStream();
        var stored = new byte[stream.Length];
        stream.ReadExactly(stored);
        Assert.Equal(PngBytes, stored);

        // Anchored at E2 (0-based col 4, row 1 in the drawing markup).
        var from = picture.Parent switch
        {
            Xdr.OneCellAnchor one => one.FromMarker,
            Xdr.TwoCellAnchor two => two.FromMarker,
            Xdr.AbsoluteAnchor => null,
            _ => null,
        };
        Assert.NotNull(from);
        Assert.Equal("4", from!.ColumnId!.Text);
        Assert.Equal("1", from.RowId!.Text);
    }

    [Fact]
    public void Jpeg_is_sniffed_and_embedded()
    {
        var file = CreateWorkbook();
        File.WriteAllBytes(Path.Combine(Dir, "photo.bin"), JpegBytes); // wrong extension on purpose

        var envelope = EditOps(file, AddOp("/Sheet1", "image",
            ("src", "photo.bin"), ("anchor", "B2")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("jpeg", Json(envelope)["data"]!["ops"]![0]!["format"]!.GetValue<string>());
        AssertValidatorClean(file);

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (worksheetPart, _) = RawPictures(document);
        var imagePart = Assert.Single(worksheetPart.DrawingsPart!.ImageParts);
        Assert.Equal("image/jpeg", imagePart.ContentType);
    }

    [Fact]
    public void Non_png_jpeg_bytes_are_unsupported_naming_the_formats()
    {
        var file = CreateWorkbook();
        File.WriteAllBytes(Path.Combine(Dir, "anim.gif"), "GIF89a trailer"u8.ToArray());

        var envelope = EditOps(file, AddOp("/Sheet1", "image",
            ("src", "anim.gif"), ("anchor", "B2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Equal(["png", "jpeg"], envelope.Error.Candidates!);
        Assert.Contains("convert", envelope.Error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Src_escaping_the_sandbox_is_denied_without_reading_the_file()
    {
        var file = CreateWorkbook();
        var outside = Path.Combine(Path.GetTempPath(), "aioffice-outside-" + Guid.NewGuid().ToString("N") + ".png");
        File.WriteAllBytes(outside, PngBytes);
        try
        {
            foreach (var src in new[] { outside, "../" + Path.GetFileName(outside) })
            {
                var envelope = EditOps(file, AddOp("/Sheet1", "image",
                    ("src", src), ("anchor", "B2")));

                Assert.False(envelope.IsOk);
                Assert.Equal(ErrorCodes.SandboxDenied, envelope.Error!.Code);
            }
        }
        finally
        {
            File.Delete(outside);
        }
    }

    [Fact]
    public void Missing_src_file_is_file_not_found()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "image",
            ("src", "nope.png"), ("anchor", "B2")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.FileNotFound, envelope.Error!.Code);
    }

    [Fact]
    public void Width_only_keeps_the_aspect_ratio()
    {
        var file = CreateWorkbook();
        var src = WritePng(); // 4x2 source

        var envelope = EditOps(file, AddOp("/Sheet1", "image",
            ("src", src), ("anchor", "B2"), ("widthPx", 100)));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal(100, detail["widthPx"]!.GetValue<int>());
        Assert.Equal(50, detail["heightPx"]!.GetValue<int>()); // 4:2 kept
    }

    [Fact]
    public void Explicit_width_and_height_win()
    {
        var file = CreateWorkbook();
        var src = WritePng();

        Assert.True(EditOps(file, AddOp("/Sheet1", "image",
            ("src", src), ("anchor", "B2"), ("widthPx", 64), ("heightPx", 48))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/image[1]"))));
        Assert.Equal(64, data["widthPx"]!.GetValue<int>());
        Assert.Equal(48, data["heightPx"]!.GetValue<int>());
    }

    [Fact]
    public void Get_returns_name_format_anchor_and_size()
    {
        var file = CreateWorkbook();
        var src = WritePng();
        Assert.True(EditOps(file, AddOp("/Sheet1", "image",
            ("src", src), ("anchor", "E2"), ("name", "Logo"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/image[1]"))));

        Assert.Equal("/Sheet1/image[1]", data["path"]!.GetValue<string>());
        Assert.Equal("image", data["kind"]!.GetValue<string>());
        Assert.Equal("Logo", data["name"]!.GetValue<string>());
        Assert.Equal("png", data["format"]!.GetValue<string>());
        Assert.Equal("E2", data["anchor"]!.GetValue<string>());
        Assert.Equal(4, data["widthPx"]!.GetValue<int>());
        Assert.Equal(2, data["heightPx"]!.GetValue<int>());
    }

    [Fact]
    public void Get_missing_image_is_invalid_path_with_candidates()
    {
        var file = CreateWorkbook();
        var src = WritePng();
        Assert.True(EditOps(file, AddOp("/Sheet1", "image", ("src", src), ("anchor", "B2"))).IsOk);

        var envelope = Handler.Get(Ctx(file, ("path", "/Sheet1/image[2]")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1/image[1]", envelope.Error.Candidates!);
    }

    [Fact]
    public void Read_structure_lists_images()
    {
        var file = CreateWorkbook();
        var src = WritePng();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "image", ("src", src), ("anchor", "B2"), ("name", "One")),
            AddOp("/Sheet1", "image", ("src", src), ("anchor", "E8"), ("name", "Two"))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "structure"))));

        var images = data["sheets"]![0]!["images"]!.AsArray();
        Assert.Equal(2, images.Count);
        Assert.Equal("/Sheet1/image[1]", images[0]!["path"]!.GetValue<string>());
        Assert.Equal("One", images[0]!["name"]!.GetValue<string>());
        Assert.Equal("E8", images[1]!["anchor"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_image_shifts_indices_and_stays_clean()
    {
        var file = CreateWorkbook();
        var src = WritePng();
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "image", ("src", src), ("anchor", "B2"), ("name", "One")),
            AddOp("/Sheet1", "image", ("src", src), ("anchor", "E8"), ("name", "Two"))).IsOk);

        var remove = EditOps(file, RemoveOp("/Sheet1/image[1]"));
        Assert.True(remove.IsOk, remove.ToJson());
        AssertValidatorClean(file);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/image[1]"))));
        Assert.Equal("Two", data["name"]!.GetValue<string>()); // shifted into [1]

        Assert.True(EditOps(file, RemoveOp("/Sheet1/image[1]")).IsOk);
        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var (_, pictures) = RawPictures(document);
        Assert.Empty(pictures);
    }

    [Fact]
    public void Remove_missing_image_is_invalid_path_with_candidates()
    {
        var file = CreateWorkbook();
        var src = WritePng();
        Assert.True(EditOps(file, AddOp("/Sheet1", "image", ("src", src), ("anchor", "B2"))).IsOk);

        var envelope = EditOps(file, RemoveOp("/Sheet1/image[9]"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1/image[1]", envelope.Error.Candidates!);
    }

    [Fact]
    public void Image_survives_a_later_unrelated_edit()
    {
        var file = CreateWorkbook();
        var src = WritePng();
        Assert.True(EditOps(file, AddOp("/Sheet1", "image",
            ("src", src), ("anchor", "E2"), ("name", "Logo"))).IsOk);

        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/image[1]"))));
        Assert.Equal("Logo", data["name"]!.GetValue<string>());
        Assert.Equal("E2", data["anchor"]!.GetValue<string>());
    }

    [Fact]
    public void Image_and_chart_coexist_on_one_sheet()
    {
        var file = CreateWorkbook();
        var src = WritePng();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B3", ("values", new System.Text.Json.Nodes.JsonArray(
            new System.Text.Json.Nodes.JsonArray("Item", "Qty"),
            new System.Text.Json.Nodes.JsonArray("ant", 5),
            new System.Text.Json.Nodes.JsonArray("bee", 7))))).IsOk);

        var envelope = EditOps(
            file,
            AddOp("/Sheet1", "image", ("src", src), ("anchor", "E2")),
            AddOp("/Sheet1", "chart", ("kind", "bar"), ("dataRange", "A1:B3"), ("anchor", "E20")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);
        Assert.Equal("png", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/image[1]"))))["format"]!.GetValue<string>());
        Assert.Equal("bar", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/chart[1]"))))["chartKind"]!.GetValue<string>());
    }

    [Fact]
    public void Set_on_an_image_is_a_typed_unsupported_feature()
    {
        var file = CreateWorkbook();
        var src = WritePng();
        Assert.True(EditOps(file, AddOp("/Sheet1", "image", ("src", src), ("anchor", "B2"))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/image[1]", ("value", 1)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("remove", envelope.Error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }
}
