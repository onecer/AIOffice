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

    [Fact]
    public void GradientObjectBackground_IsTypedUnsupportedFeature()
    {
        Create();
        var props = new JsonObject
        {
            ["background"] = new JsonObject { ["gradient"] = new JsonArray("000000", "FFFFFF") },
        };
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", "/slide[1]", props: props)]);

        TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
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
        Assert.Equal(new[] { "background" }, error.Candidates!);
    }
}
