using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class ImageTests : WordTestBase
{
    /// <summary>A real, decodable 1x1 PNG (for data-URI rendering).</summary>
    private const string OnePixelPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

    /// <summary>A PNG header claiming the given size — enough for the sniffer, which never decodes.</summary>
    private static byte[] MakePng(int width, int height)
    {
        var bytes = new byte[33];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }.CopyTo(bytes, 0);
        bytes[11] = 13; // IHDR length
        "IHDR"u8.ToArray().CopyTo(bytes, 12);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(16, 4), width);
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(20, 4), height);
        bytes[24] = 8; // bit depth
        bytes[25] = 6; // RGBA
        return bytes;
    }

    /// <summary>A minimal JPEG byte stream: SOI + SOF0 carrying the dimensions + EOI.</summary>
    private static byte[] MakeJpeg(int width, int height) =>
    [
        0xFF, 0xD8, // SOI
        0xFF, 0xC0, 0x00, 0x0B, 0x08, // SOF0, length 11, precision 8
        (byte)(height >> 8), (byte)height,
        (byte)(width >> 8), (byte)width,
        0x01, 0x01, 0x11, 0x00, // one component
        0xFF, 0xD9, // EOI
    ];

    private string WriteImage(string name, byte[] bytes)
    {
        var path = Path.Combine(Dir, name);
        File.WriteAllBytes(path, bytes);
        return name;
    }

    [Fact]
    public void Add_png_with_width_scales_height_from_the_aspect_ratio()
    {
        var file = CreateDoc(title: "Pictures");
        WriteImage("wide.png", MakePng(200, 100));

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"image","props":{"src":"wide.png","width":"10cm"}}]""");

        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("/body/p[3]", summary["path"]!.GetValue<string>());
        Assert.Equal(10, summary["widthCm"]!.GetValue<double>());
        Assert.Equal(5, summary["heightCm"]!.GetValue<double>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var part = Assert.Single(doc.MainDocumentPart!.ImageParts);
            Assert.Equal("image/png", part.ContentType);
        }

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[3]" })))["properties"]!;
        Assert.Equal("image", got["kind"]!.GetValue<string>());
        Assert.Equal(10, got["widthCm"]!.GetValue<double>());
        Assert.Equal(5, got["heightCm"]!.GetValue<double>());
        Assert.Contains("not stored", got["note"]!.GetValue<string>(), StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Both_width_and_height_override_the_aspect_ratio()
    {
        var file = CreateDoc(title: "Sized");
        WriteImage("square.png", MakePng(50, 50));

        var envelope = Edit(file, """
            [{"op":"add","path":"/body","type":"image","props":{"src":"square.png","width":"4cm","height":"3cm"}}]
            """);

        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal(4, summary["widthCm"]!.GetValue<double>());
        Assert.Equal(3, summary["heightCm"]!.GetValue<double>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Jpeg_is_sniffed_by_header_even_with_a_png_extension()
    {
        var file = CreateDoc(title: "Liar file");
        WriteImage("really-a.png", MakeJpeg(2, 1)); // jpeg bytes behind a png name

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"image","props":{"src":"really-a.png","width":"2cm"}}]""");

        Assert.Equal("jpeg", Data(envelope)["ops"]!.AsArray()[0]!["format"]!.GetValue<string>());
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var part = Assert.Single(doc.MainDocumentPart!.ImageParts);
        Assert.Equal("image/jpeg", part.ContentType);
    }

    [Fact]
    public void Escaping_src_is_sandbox_denied_and_changes_nothing()
    {
        var file = CreateDoc(title: "Locked");
        var before = File.ReadAllBytes(file);

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"image","props":{"src":"../../etc/escape.png","width":"5cm"}}]"""));

        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
        Assert.Equal(before, File.ReadAllBytes(file));
    }

    [Fact]
    public void Non_image_bytes_are_unsupported_feature()
    {
        var file = CreateDoc(title: "Text trap");
        WriteImage("not-an-image.png", "just some text"u8.ToArray());

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"image","props":{"src":"not-an-image.png"}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("PNG", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Missing_src_is_invalid_args()
    {
        var file = CreateDoc(title: "No src");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"image","props":{"width":"5cm"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("src", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_can_land_before_an_existing_paragraph()
    {
        var file = CreateDoc(title: "Order");
        WriteImage("pic.png", MakePng(10, 10));

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"image","position":"before","props":{"src":"pic.png","width":"1cm"}}]
            """);

        Assert.Equal("/body/p[1]", Data(envelope)["ops"]!.AsArray()[0]!["path"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Render_html_embeds_a_data_uri_img_with_canonical_path()
    {
        var file = CreateDoc(title: "Gallery");
        WriteImage("dot.png", Convert.FromBase64String(OnePixelPngBase64));
        Edit(file, """[{"op":"add","path":"/body","type":"image","props":{"src":"dot.png","width":"2cm"}}]""");

        var html = Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();

        Assert.Contains("<img data-aio-path=\"/body/p[3]/run[1]\"", html, StringComparison.Ordinal);
        Assert.Contains("src=\"data:image/png;base64,", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Tracked_image_add_is_a_rejectable_insert_revision()
    {
        var file = CreateDoc(title: "Tracked pics");
        WriteImage("pic.png", MakePng(10, 10));

        Edit(
            file,
            """[{"op":"add","path":"/body","type":"image","props":{"src":"pic.png","width":"1cm"}}]""",
            new JsonObject { ["track"] = true });

        var revisions = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "revisions" })))["revisions"]!.AsArray();
        Assert.Contains(revisions, r => r!["kind"]!.GetValue<string>() == "insert");
        AssertValidatesClean(file);

        Edit(file, """[{"op":"reject","path":"/body"}]""");
        AssertValidatesClean(file);
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
    }
}
