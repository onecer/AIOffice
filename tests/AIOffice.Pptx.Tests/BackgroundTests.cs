using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>M2 real slide backgrounds: a proper p:bg solid fill, not a full-bleed rectangle.</summary>
public sealed class BackgroundTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() =>
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Cover"))));

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    [Fact]
    public void SetBackground_WritesProperBgSolidFill()
    {
        Create();
        var data = Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", "0F172A"))));
        Assert.Equal("/slide[1]", data["results"]![0]!["target"]!.GetValue<string>());

        // Reopen and verify the raw OOXML: p:cSld/p:bg/p:bgPr/a:solidFill/a:srgbClr.
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slideData = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!;
            var background = slideData.Background!.BackgroundProperties!;
            Assert.Equal("0F172A", background.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.True(
                slideData.Background.IsBefore(slideData.ShapeTree!),
                "p:bg must precede p:spTree inside p:cSld");
            Assert.DoesNotContain(
                slideData.ShapeTree!.Elements<P.Shape>(),
                s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Background");
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        Assert.Equal("0F172A", detail["background"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetBackground_SolidPath_IsByteStableVs1_13()
    {
        // Regression: a solid hex must serialize to the EXACT p:bg XML 1.13.0 produced —
        // the legacy branch stays first and untouched by the v1.14 object widening.
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", "0F172A"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var background = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.Background!;
        Assert.Equal(
            "<p:bg xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">" +
            "<p:bgPr><a:solidFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
            "<a:srgbClr val=\"0F172A\" /></a:solidFill>" +
            "<a:effectLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" /></p:bgPr></p:bg>",
            background.OuterXml);
    }

    [Fact]
    public void SetBackground_ReplacesThePreviousOne()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", "#112233"))));
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", "445566"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var slideData = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!;
        Assert.Single(slideData.Elements<P.Background>());
        Assert.Equal(
            "445566",
            slideData.Background!.BackgroundProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
    }

    [Fact]
    public void GetSlide_WithoutBackground_OmitsTheField()
    {
        Create();
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));

        Assert.Null(detail["background"]);
    }

    [Fact]
    public void AddSlide_AcceptsBackgroundProp()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(
            ("title", "Dark"),
            ("background", "1E293B"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[2]"))));
        Assert.Equal("1E293B", detail["background"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Render_UsesTheBackgroundAsSlideFill()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", "0F172A"))));

        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();
        Assert.Contains("fill=\"#0f172a\"", svg, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("linear-gradient(#000,#fff)")]
    [InlineData("image:hero.png")]
    [InlineData("hero.jpg")]
    public void GradientOrImageBackground_IsTypedUnsupportedFeature(string value)
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", value)))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Equal(ExitCodes.UnsupportedFeature, envelope.ExitCode);
        Assert.Contains("solid color", error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    // ---------------------------------------------------------- v1.14: object background fills

    /// <summary>A test PNG written into the sandbox, returned by its workspace-relative name.</summary>
    private string WritePng(string name = "bg.png", int w = 8, int h = 8)
    {
        File.WriteAllBytes(_ws.PathOf(name), TestImages.Png(w, h));
        return name;
    }

    private static JsonObject GradientSpec() => new()
    {
        ["gradient"] = new JsonObject
        {
            ["type"] = "linear",
            ["angle"] = 90,
            ["stops"] = new JsonArray(
                new JsonObject { ["color"] = "0EA5E9", ["at"] = 0 },
                new JsonObject { ["color"] = "6366F1", ["at"] = 100 }),
        },
    };

    [Fact]
    public void SetBackground_GradientObject_WritesGradFillWithStops()
    {
        Create();
        var data = Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", GradientSpec()))));
        Assert.Equal("/slide[1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slideData = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!;
            var bgPr = slideData.Background!.BackgroundProperties!;
            Assert.Single(slideData.Elements<P.Background>());

            // p:bg/p:bgPr/a:gradFill with the two stops, read straight off the part.
            var grad = bgPr.GetFirstChild<A.GradientFill>();
            Assert.NotNull(grad);
            Assert.Null(bgPr.GetFirstChild<A.SolidFill>());
            var stops = grad!.GetFirstChild<A.GradientStopList>()!.Elements<A.GradientStop>().ToList();
            Assert.Equal(2, stops.Count);
            Assert.Equal("0EA5E9", stops[0].RgbColorModelHex!.Val!.Value);
            Assert.Equal(0, stops[0].Position!.Value);
            Assert.Equal("6366F1", stops[1].RgbColorModelHex!.Val!.Value);
            Assert.Equal(100000, stops[1].Position!.Value);
            // angle 90° → 90 * 60000 1/60000ths.
            Assert.Equal(90 * 60000, grad.GetFirstChild<A.LinearGradientFill>()!.Angle!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetBackground_ImageObject_EmbedsBlipFillOnANewPart()
    {
        Create();
        var src = WritePng();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(
            ("background", new JsonObject { ["image"] = new JsonObject { ["src"] = src } }))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var bgPr = slidePart.Slide!.CommonSlideData!.Background!.BackgroundProperties!;

            // p:bg/p:bgPr/a:blipFill referencing a NEW image part on the slide part.
            var blip = bgPr.GetFirstChild<A.BlipFill>()!.Blip!;
            var imagePart = Assert.Single(slidePart.ImageParts);
            Assert.Equal(slidePart.GetIdOfPart(imagePart), blip.Embed!.Value); // rel points at it, no orphan

            using var stream = imagePart.GetStream();
            Assert.Equal(TestImages.Png(8, 8).Length, stream.Length);
        }

        // Package opens repair-free (validator clean), no orphaned media part.
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetBackground_GradientThenSolid_LeavesExactlyOneBgAndNoOrphanMedia()
    {
        Create();
        var src = WritePng();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(
            ("background", new JsonObject { ["image"] = new JsonObject { ["src"] = src } }))));
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", GradientSpec()))));
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", "445566"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var slideData = slidePart.Slide!.CommonSlideData!;

            // Exactly one p:bg, and it is the final solid fill.
            Assert.Single(slideData.Elements<P.Background>());
            var bgPr = slideData.Background!.BackgroundProperties!;
            Assert.Equal("445566", bgPr.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Null(bgPr.GetFirstChild<A.GradientFill>());
            Assert.Null(bgPr.GetFirstChild<A.BlipFill>());

            // The earlier image background's media part was pruned on replace — no orphan left.
            Assert.Empty(slidePart.ImageParts);
        }

        // The package must still validate clean (repair-free) after the replacements.
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetBackground_JsonArray_StaysUnsupportedFeatureLike1_13()
    {
        // A JsonArray background was rejected with UnsupportedFeature in 1.13.0 (only a JsonObject
        // is the new gradient/image path); its error contract must not drift to InvalidArgs.
        Create();
        var props = new JsonObject { ["background"] = new JsonArray("0F172A", "112233") };
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", "/slide[1]", props: props)]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Equal(ExitCodes.UnsupportedFeature, envelope.ExitCode);
        Assert.Contains("solid color", error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetSlide_ReportsBackgroundKind()
    {
        Create();
        var src = WritePng();

        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", "0F172A"))));
        Assert.Equal("solid", Kind("/slide[1]"));

        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", GradientSpec()))));
        Assert.Equal("gradient", Kind("/slide[1]"));

        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(
            ("background", new JsonObject { ["image"] = new JsonObject { ["src"] = src } }))));
        Assert.Equal("image", Kind("/slide[1]"));
    }

    private string? Kind(string path)
    {
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        return detail["backgroundKind"]?.GetValue<string>();
    }

    [Fact]
    public void SetBackground_GradientWithEmptyStops_IsInvalidArgs()
    {
        Create();
        var props = new JsonObject
        {
            ["background"] = new JsonObject
            {
                ["gradient"] = new JsonObject { ["stops"] = new JsonArray() },
            },
        };
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", "/slide[1]", props: props)]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void SetBackground_ImageWithoutSrc_IsInvalidArgs()
    {
        Create();
        var props = new JsonObject
        {
            ["background"] = new JsonObject { ["image"] = new JsonObject { ["mode"] = "tile" } },
        };
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", "/slide[1]", props: props)]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void SetBackground_ObjectWithoutGradientOrImage_IsInvalidArgs()
    {
        Create();
        var props = new JsonObject
        {
            ["background"] = new JsonObject { ["solid"] = "0F172A" },
        };
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", "/slide[1]", props: props)]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Equal(new[] { "gradient", "image" }, error.Candidates!);
    }

    [Fact]
    public void InvalidBackgroundColor_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("background", "bluish")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void UnknownSlideProp_IsInvalidArgsWithCandidates()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("text", "nope")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Equal(
            new[] { "background", "gradient", "image", "transition", "transitionDuration", "footer", "slideNumber", "date" },
            error.Candidates!);
    }
}
