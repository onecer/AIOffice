using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>M2 picture shapes: header-sniffed PNG/JPEG embedding with aspect-aware sizing.</summary>
public sealed class ImageTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() =>
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Cover"))));

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private string WriteImage(string name, byte[] bytes)
    {
        File.WriteAllBytes(_ws.PathOf(name), bytes);
        return name;
    }

    [Fact]
    public void AddPng_CreatesImagePartAndDerivesHeightFromAspect()
    {
        Create();
        WriteImage("logo.png", TestImages.Png(200, 100)); // 2:1

        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(
            ("src", "logo.png"),
            ("x", "2cm"),
            ("y", "3cm"),
            ("w", "10cm"))));
        var canonical = data["results"]![0]!["target"]!.GetValue<string>();
        Assert.Matches(@"^/slide\[1\]/shape\[@id=[0-9]+\]$", canonical);

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var imagePart = Assert.Single(slidePart.ImageParts);
            Assert.Equal("image/png", imagePart.ContentType);
            using var stream = imagePart.GetStream();
            Assert.Equal(TestImages.Png(200, 100).Length, stream.Length); // bytes embedded verbatim

            var picture = Assert.Single(slidePart.Slide!.Descendants<P.Picture>());
            Assert.Equal(
                slidePart.GetIdOfPart(imagePart),
                picture.BlipFill!.Blip!.Embed!.Value); // blipFill points at the part
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal("picture", detail["kind"]!.GetValue<string>());
        Assert.Equal(2.0, detail["x"]!.GetValue<double>());
        Assert.Equal(3.0, detail["y"]!.GetValue<double>());
        Assert.Equal(10.0, detail["w"]!.GetValue<double>());
        Assert.Equal(5.0, detail["h"]!.GetValue<double>()); // aspect honored: 10cm * (100/200)
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddPng_WidthFromHeight_AndExplicitBothWin()
    {
        Create();
        WriteImage("tall.png", TestImages.Png(100, 400)); // 1:4

        Edit(
            TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(
                ("src", "tall.png"), ("h", "8cm"))),
            TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(
                ("src", "tall.png"), ("w", "3cm"), ("h", "3cm"))));

        var slide = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        var pictures = slide["shapes"]!.AsArray()
            .Where(s => s!["kind"]!.GetValue<string>() == "picture")
            .Select(s => TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", s!["path"]!.GetValue<string>())))))
            .ToList();

        Assert.Equal(2.0, pictures[0]["w"]!.GetValue<double>()); // 8cm * (100/400)
        Assert.Equal(8.0, pictures[0]["h"]!.GetValue<double>());
        Assert.Equal(3.0, pictures[1]["w"]!.GetValue<double>()); // both given: no aspect correction
        Assert.Equal(3.0, pictures[1]["h"]!.GetValue<double>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddPng_WithoutSize_UsesNaturalPixelSizeAt96Dpi()
    {
        Create();
        WriteImage("nat.png", TestImages.Png(96, 48));

        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", "nat.png"))));
        var canonical = data["results"]![0]!["target"]!.GetValue<string>();

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal(2.54, detail["w"]!.GetValue<double>()); // 96px at 96dpi = 1in
        Assert.Equal(1.27, detail["h"]!.GetValue<double>());
    }

    [Fact]
    public void AddJpeg_IsSniffedByHeaderNotExtension()
    {
        Create();
        WriteImage("photo.bin", TestImages.Jpeg(640, 480)); // extension lies; header decides

        Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(
            ("src", "photo.bin"), ("w", "4cm"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var imagePart = Assert.Single(doc.PresentationPart!.SlideParts.Single().ImageParts);
        Assert.Equal("image/jpeg", imagePart.ContentType);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddImage_TypePictureAlias_Works()
    {
        Create();
        WriteImage("logo.png", TestImages.Png(10, 10));

        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "picture", props: TestEnv.Props(("src", "logo.png"))));

        Assert.Matches(@"^/slide\[1\]/shape\[@id=[0-9]+\]$", data["results"]![0]!["target"]!.GetValue<string>());
    }

    [Fact]
    public void RemovePicture_ByCanonicalPath()
    {
        Create();
        WriteImage("logo.png", TestImages.Png(10, 10));
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", "logo.png"))));
        var canonical = data["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("remove", canonical));

        var query = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape[kind=picture]"))));
        Assert.Equal(0, query["count"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RenderSvg_DrawsLabeledPlaceholderWithDataAioPath()
    {
        Create();
        WriteImage("hero.png", TestImages.Png(20, 10));
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(
            ("src", "hero.png"), ("name", "Hero Shot"), ("w", "10cm"))));
        var canonical = data["results"]![0]!["target"]!.GetValue<string>();

        var rendered = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = rendered["slides"]![0]!["svg"]!.GetValue<string>();

        Assert.Contains($"<g data-aio-path=\"{canonical}\"", svg, StringComparison.Ordinal);
        Assert.Contains("[image] Hero Shot", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("<image", svg, StringComparison.Ordinal); // never rasterized into the svg
    }

    [Fact]
    public void SrcOutsideWorkspace_IsSandboxDenied()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(
                ("src", "../escapee.png"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.SandboxDenied);
        Assert.Equal(ExitCodes.SandboxDenied, envelope.ExitCode);
    }

    [Fact]
    public void MissingSrc_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("w", "4cm"))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("src", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MissingFile_IsFileNotFound()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", "ghost.png"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.FileNotFound);
    }

    [Fact]
    public void NonPngJpegBytes_IsTypedUnsupportedFeature()
    {
        Create();
        WriteImage("anim.gif", [.. "GIF89a"u8, 0x01, 0x00, 0x01, 0x00, 0x00, 0x00, 0x3B]);

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", "anim.gif"))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("Convert", error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void FailedImageOp_IsAtomic_NoWrite()
    {
        Create();
        WriteImage("logo.png", TestImages.Png(10, 10));
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", "logo.png"))),
            TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", "../outside.png"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.SandboxDenied);
        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }
}
