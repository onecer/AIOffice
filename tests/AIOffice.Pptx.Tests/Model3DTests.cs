using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// Embedded 3D models (v1.3.0): a glb/gltf model rides in a model/* data part
/// behind a poster picture fallback, addressed as /slide[i]/model3d[@id=N]. The
/// embed is validator-clean, src/poster are sandbox-resolved, and remove drops the
/// orphaned part.
/// </summary>
public sealed class Model3DTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public Model3DTests()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
    }

    public void Dispose() => _ws.Dispose();

    private void WriteFile(string name, byte[] bytes) => File.WriteAllBytes(_ws.PathOf(name), bytes);

    private JsonObject AddModel(JsonObject props) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "model3d", props: props)]));

    private static int ModelDataPartCount(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.Single().DataPartReferenceRelationships
            .Select(r => r.DataPart)
            .OfType<MediaDataPart>()
            .Count(p => p.ContentType.StartsWith("model/", StringComparison.Ordinal));

    [Fact]
    public void AddGlb_EmbedsModelPartBehindPosterAndValidates()
    {
        WriteFile("chair.glb", TestModels.Glb());
        WriteFile("thumb.png", TestImages.Png(8, 8));

        var data = AddModel(new JsonObject
        {
            ["src"] = "chair.glb",
            ["poster"] = "thumb.png",
            ["x"] = "2cm",
            ["y"] = "2cm",
            ["w"] = "12cm",
            ["h"] = "10cm",
        });

        Assert.StartsWith("/slide[1]/model3d[@id=", data["results"]![0]!["target"]!.GetValue<string>(), StringComparison.Ordinal);

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            Assert.Equal(1, ModelDataPartCount(doc));

            // The host picture carries a p14:media extension and a poster blip-fill.
            var picture = slidePart.Slide!.Descendants<P.Picture>().Single();
            Assert.NotNull(picture.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().SingleOrDefault());
            Assert.NotNull(picture.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().SingleOrDefault()?.Embed?.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddGltf_EmbedsJsonModelAndValidates()
    {
        WriteFile("lamp.gltf", TestModels.Gltf());
        AddModel(new JsonObject { ["src"] = "lamp.gltf" });

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var part = doc.PresentationPart!.SlideParts.Single().DataPartReferenceRelationships
                .Select(r => r.DataPart).OfType<MediaDataPart>()
                .Single(p => p.ContentType.StartsWith("model/", StringComparison.Ordinal));
            Assert.Equal("model/gltf+json", part.ContentType);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Add_EmitsModel3DAsMediaWarning()
    {
        WriteFile("chair.glb", TestModels.Glb());
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "model3d", props: new JsonObject { ["src"] = "chair.glb" })]);

        Assert.True(envelope.IsOk);
        Assert.NotNull(envelope.Meta?.Warnings);
        Assert.Contains(envelope.Meta!.Warnings!, w => w.Code == WarningCodes.Model3DAsMedia);
    }

    [Fact]
    public void Get_ReportsKindGeometryAndSandboxedGeometry()
    {
        WriteFile("chair.glb", TestModels.Glb());
        AddModel(new JsonObject { ["src"] = "chair.glb", ["x"] = "2cm", ["y"] = "3cm", ["w"] = "12cm", ["h"] = "10cm" });

        var data = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/model3d[1]"))));
        Assert.Equal("model3d", data["kind"]!.GetValue<string>());
        Assert.Equal("glb", data["geometry"]!.GetValue<string>());
        Assert.Equal("chair.glb", data["src"]!.GetValue<string>());
        Assert.Equal(12.0, data["w"]!.GetValue<double>());
        Assert.Equal(10.0, data["h"]!.GetValue<double>());
    }

    [Fact]
    public void Get_ByStableId_ResolvesSameModel()
    {
        WriteFile("chair.glb", TestModels.Glb());
        var path = AddModel(new JsonObject { ["src"] = "chair.glb" })["results"]![0]!["target"]!.GetValue<string>();

        // The canonical model path is the @id form; getting it back resolves.
        var data = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("model3d", data["kind"]!.GetValue<string>());
    }

    [Fact]
    public void SandboxDenied_OnSrcOutsideWorkspace()
    {
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("add", "/slide[1]", type: "model3d", props: new JsonObject { ["src"] = "../../etc/passwd" })]),
            ErrorCodes.SandboxDenied);
    }

    [Fact]
    public void SandboxDenied_OnPosterOutsideWorkspace()
    {
        WriteFile("chair.glb", TestModels.Glb());
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("add", "/slide[1]", type: "model3d", props: new JsonObject
                {
                    ["src"] = "chair.glb",
                    ["poster"] = "../../etc/hosts",
                })]),
            ErrorCodes.SandboxDenied);
    }

    [Fact]
    public void UnknownModelFormat_IsTypedUnsupported()
    {
        WriteFile("x.obj", System.Text.Encoding.ASCII.GetBytes("v 0 0 0\n"));
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("add", "/slide[1]", type: "model3d", props: new JsonObject { ["src"] = "x.obj" })]),
            ErrorCodes.UnsupportedFeature);
    }

    [Fact]
    public void Structure_ListsModelsPerSlide()
    {
        WriteFile("chair.glb", TestModels.Glb());
        AddModel(new JsonObject { ["src"] = "chair.glb" });

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var models = data["slides"]![0]!["models"]!.AsArray();
        Assert.Single(models);
        Assert.Equal("model3d", models[0]!["kind"]!.GetValue<string>());
    }

    [Fact]
    public void SvgRender_ShowsModelPlaceholderWithDataPath()
    {
        WriteFile("chair.glb", TestModels.Glb());
        var path = AddModel(new JsonObject { ["src"] = "chair.glb" })["results"]![0]!["target"]!.GetValue<string>();
        var shapeId = path[(path.IndexOf("=", StringComparison.Ordinal) + 1)..].TrimEnd(']');

        var rendered = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = rendered.ToJsonString();
        Assert.Contains("[3d model]", svg, StringComparison.Ordinal);
        Assert.Contains("aio-model3d", svg, StringComparison.Ordinal);
        Assert.Contains($"/slide[1]/shape[@id={shapeId}]", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void Remove_DropsModelPartAndKeepsValid()
    {
        WriteFile("chair.glb", TestModels.Glb());
        AddModel(new JsonObject { ["src"] = "chair.glb" });

        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("remove", "/slide[1]/model3d[1]")]));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Equal(0, ModelDataPartCount(doc));
            Assert.Empty(doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Picture>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Add_RoundTripsModelBytesExactly()
    {
        var glb = TestModels.Glb();
        WriteFile("chair.glb", glb);
        AddModel(new JsonObject { ["src"] = "chair.glb" });

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var part = doc.PresentationPart!.SlideParts.Single().DataPartReferenceRelationships
            .Select(r => r.DataPart).OfType<MediaDataPart>()
            .Single(p => p.ContentType.StartsWith("model/", StringComparison.Ordinal));
        using var stream = part.GetStream();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        Assert.Equal(glb, ms.ToArray());
    }

    [Fact]
    public void MissingSrc_IsInvalidArgs()
    {
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("add", "/slide[1]", type: "model3d", props: new JsonObject())]),
            ErrorCodes.InvalidArgs);
    }
}
