using System.Text;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>v1.1 embedded media: video/audio on a slide as a MediaDataPart + p:pic with timing wiring.</summary>
public sealed class MediaTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() => TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    /// <summary>A tiny but header-valid mp4 (ftyp isom box).</summary>
    private string WriteMp4(string name = "clip.mp4")
    {
        var bytes = new byte[64];
        bytes[3] = 0x18;
        Encoding.ASCII.GetBytes("ftyp").CopyTo(bytes, 4);
        Encoding.ASCII.GetBytes("isom").CopyTo(bytes, 8);
        return WriteFile(name, bytes);
    }

    /// <summary>A tiny but header-valid mp3 (ID3 tag).</summary>
    private string WriteMp3(string name = "sound.mp3")
    {
        var bytes = new byte[64];
        Encoding.ASCII.GetBytes("ID3").CopyTo(bytes, 0);
        bytes[3] = 0x03;
        return WriteFile(name, bytes);
    }

    /// <summary>A tiny but header-valid wav (RIFF/WAVE).</summary>
    private string WriteWav(string name = "tone.wav")
    {
        var bytes = new byte[64];
        Encoding.ASCII.GetBytes("RIFF").CopyTo(bytes, 0);
        Encoding.ASCII.GetBytes("WAVE").CopyTo(bytes, 8);
        return WriteFile(name, bytes);
    }

    private string WriteFile(string name, byte[] bytes)
    {
        var path = Path.Combine(_ws.Dir, name);
        File.WriteAllBytes(path, bytes);
        return name;
    }

    [Fact]
    public void AddVideo_EmbedsMediaPartAndPicture()
    {
        Create();
        WriteMp4();
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(
            ("src", "clip.mp4"), ("x", "2cm"), ("y", "3cm"), ("w", "16cm"), ("h", "9cm"))));
        Assert.StartsWith("/slide[1]/media[@id=", data["results"]![0]!["target"]!.GetValue<string>(), StringComparison.Ordinal);

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Single(doc.DataParts);
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var pic = slidePart.Slide!.Descendants<P.Picture>().Single();
            var nvPr = pic.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!;
            Assert.NotNull(nvPr.GetFirstChild<A.VideoFromFile>());
            // A video reference (r:link) + a media reference (r:embed) both resolve to the data part.
            Assert.Equal(2, slidePart.DataPartReferenceRelationships.Count());
            // The click action turns the picture into a media trigger.
            Assert.Equal("ppaction://media", nvPr.Parent!.Descendants<A.HyperlinkOnClick>().Single().Action!.Value);
            // Timing tree carries a video node targeting the picture.
            Assert.Single(slidePart.Slide.Descendants<P.Video>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("sound.mp3")]
    [InlineData("tone.wav")]
    public void AddAudio_EmbedsAudioPartAndPicture(string name)
    {
        Create();
        if (name.EndsWith(".mp3", StringComparison.Ordinal))
        {
            WriteMp3(name);
        }
        else
        {
            WriteWav(name);
        }

        Edit(TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(("src", name))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var pic = slidePart.Slide!.Descendants<P.Picture>().Single();
            var nvPr = pic.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!;
            Assert.NotNull(nvPr.GetFirstChild<A.AudioFromFile>());
            Assert.Null(nvPr.GetFirstChild<A.VideoFromFile>());
            Assert.Single(slidePart.Slide.Descendants<P.Audio>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddVideo_WithPoster_EmbedsTheImage()
    {
        Create();
        WriteMp4();
        File.WriteAllBytes(Path.Combine(_ws.Dir, "thumb.png"), TestImages.Png(64, 48));
        Edit(TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(
            ("src", "clip.mp4"), ("poster", "thumb.png"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var pic = slidePart.Slide!.Descendants<P.Picture>().Single();
            // The poster is a real image part referenced by the blip.
            Assert.NotNull(pic.BlipFill!.Blip!.Embed?.Value);
            Assert.Single(slidePart.ImageParts);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddVideo_Autoplay_StartsOnLoad()
    {
        Create();
        WriteMp4();
        Edit(TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(("src", "clip.mp4"), ("autoplay", true))));

        var media = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var slideMedia = media["slides"]!.AsArray()[0]!["media"]!.AsArray();
        Assert.True(slideMedia[0]!["autoplay"]!.GetValue<bool>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_MediaPath_ReportsKindAndGeometry()
    {
        Create();
        WriteMp4();
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(
            ("src", "clip.mp4"), ("x", "2cm"), ("y", "3cm"), ("w", "16cm"), ("h", "9cm"))));
        var mediaPath = data["results"]![0]!["target"]!.GetValue<string>();

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", mediaPath))));
        Assert.Equal("video", detail["kind"]!.GetValue<string>());
        Assert.Equal("clip.mp4", detail["src"]!.GetValue<string>());
        Assert.Equal(2, detail["x"]!.GetValue<double>());
        Assert.Equal(16, detail["w"]!.GetValue<double>());
        Assert.StartsWith("/slide[1]/shape[@id=", detail["shapePath"]!.GetValue<string>(), StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Remove_MediaPath_DropsPictureAndDataPart()
    {
        Create();
        WriteMp4();
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(("src", "clip.mp4"))));
        var mediaPath = data["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("remove", mediaPath));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            Assert.Empty(slidePart.Slide!.Descendants<P.Picture>());
            Assert.Empty(slidePart.DataPartReferenceRelationships);
            Assert.Empty(doc.DataParts);
            Assert.Empty(slidePart.Slide.Descendants<P.Video>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Remove_ByShapePath_AlsoDropsDataPart()
    {
        Create();
        WriteMp4();
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(("src", "clip.mp4"))));
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", data["results"]![0]!["target"]!.GetValue<string>()))));
        var shapePath = detail["shapePath"]!.GetValue<string>();

        Edit(TestEnv.Op("remove", shapePath));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Empty(doc.DataParts);
            Assert.Empty(doc.PresentationPart!.SlideParts.Single().DataPartReferenceRelationships);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddMedia_SrcOutsideWorkspace_IsSandboxDenied()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(("src", "../escape.mp4")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.SandboxDenied);
    }

    [Fact]
    public void AddMedia_PosterOutsideWorkspace_IsSandboxDenied()
    {
        Create();
        WriteMp4();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(
                ("src", "clip.mp4"), ("poster", "../escape.png")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.SandboxDenied);
    }

    [Fact]
    public void AddMedia_UnknownFormat_IsTypedUnsupported()
    {
        Create();
        File.WriteAllBytes(Path.Combine(_ws.Dir, "data.bin"), [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07]);
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(("src", "data.bin")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
    }

    [Fact]
    public void Structure_ListsSlideMedia()
    {
        Create();
        WriteMp4();
        Edit(TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(("src", "clip.mp4"))));

        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var media = structure["slides"]!.AsArray()[0]!["media"]!.AsArray();
        Assert.Single(media);
        Assert.Equal("video", media[0]!["kind"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Svg_Media_DrawsPlaceholderWithPlayGlyph()
    {
        Create();
        WriteMp4();
        Edit(TestEnv.Op("add", "/slide[1]", type: "media", props: TestEnv.Props(("src", "clip.mp4"))));

        var rendered = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = rendered["slides"]![0]!["svg"]!.GetValue<string>();
        Assert.Contains("aio-media-play", svg, StringComparison.Ordinal);
        Assert.Contains("[video]", svg, StringComparison.Ordinal);
        Assert.Contains("data-aio-path=\"/slide[1]/shape[@id=", svg, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
