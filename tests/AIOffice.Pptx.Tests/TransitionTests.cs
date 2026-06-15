using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>M3 slide transitions: p:transition set/get/outline round-trips.</summary>
public sealed class TransitionTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Cover"))));
    }

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    [Fact]
    public void SetFadeWithDuration_PersistsAndReflects()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(
            ("transition", "fade"), ("transitionDuration", "0.5s"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var transition = doc.PresentationPart!.SlideParts.Single().Slide!.Transition!;
            Assert.IsType<P.FadeTransition>(transition.ChildElements.Single());
            Assert.Equal("500", transition.Duration!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        Assert.Equal("fade", detail["transition"]!.GetValue<string>());
        Assert.Equal("0.5s", detail["transitionDuration"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("push", typeof(P.PushTransition))]
    [InlineData("wipe", typeof(P.WipeTransition))]
    public void SetKind_ReopenVerify(string kind, Type effectType)
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("transition", kind))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var transition = doc.PresentationPart!.SlideParts.Single().Slide!.Transition!;
            Assert.IsType(effectType, transition.ChildElements.Single());
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        Assert.Equal(kind, detail["transition"]!.GetValue<string>());
        Assert.Null(detail["transitionDuration"]);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Outline_ShowsTransitionPerSlide()
    {
        Create();
        Edit(
            TestEnv.Op("add", "/slide[2]", type: "slide"),
            TestEnv.Op("set", "/slide[2]", props: TestEnv.Props(
                ("transition", "wipe"), ("transitionDuration", "1.25s"))));

        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        var slides = outline["slides"]!.AsArray();
        Assert.Null(slides[0]!["transition"]);
        Assert.Equal("wipe", slides[1]!["transition"]!.GetValue<string>());
        Assert.Equal("1.25s", slides[1]!["transitionDuration"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetNone_RemovesTheTransition()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("transition", "fade"))));
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("transition", "none"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Null(doc.PresentationPart!.SlideParts.Single().Slide!.Transition);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        Assert.Null(detail["transition"]);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("split", typeof(P.SplitTransition))]
    [InlineData("cut", typeof(P.CutTransition))]
    [InlineData("zoom", typeof(P.ZoomTransition))]
    public void SetExpandedKind_ReopenVerify(string kind, Type effectType)
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("transition", kind))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var transition = doc.PresentationPart!.SlideParts.Single().Slide!.Transition!;
            Assert.IsType(effectType, transition.ChildElements.Single());
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        Assert.Equal(kind, detail["transition"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetReveal_PersistsP14EffectAndReflects()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("transition", "reveal"), ("transitionDuration", "0.75s"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var transition = doc.PresentationPart!.SlideParts.Single().Slide!.Transition!;
            Assert.IsType<DocumentFormat.OpenXml.Office2010.PowerPoint.RevealTransition>(transition.ChildElements.Single());
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        Assert.Equal("reveal", detail["transition"]!.GetValue<string>());
        Assert.Equal("0.75s", detail["transitionDuration"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Morph_IsTypedUnsupportedWithCandidates()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("transition", "morph")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Equal(["none", "fade", "push", "wipe", "split", "reveal", "cut", "zoom"], error.Candidates!);
    }

    [Fact]
    public void UnknownKind_IsTypedUnsupportedWithCandidates()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("transition", "dissolve")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Equal(["none", "fade", "push", "wipe", "split", "reveal", "cut", "zoom"], error.Candidates!);
    }

    [Fact]
    public void DurationWithoutTransition_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("transitionDuration", "0.5s")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("transition", error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DurationOnExistingTransition_UpdatesInPlace()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("transition", "push"))));
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("transitionDuration", "750ms"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        Assert.Equal("push", detail["transition"]!.GetValue<string>());
        Assert.Equal("0.75s", detail["transitionDuration"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void InvalidDuration_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(
                ("transition", "fade"), ("transitionDuration", "fast")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }
}
