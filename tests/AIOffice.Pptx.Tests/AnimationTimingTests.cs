using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.7.0 animation timing completeness (additive on the existing set op): repeat
/// (none|N|untilClick|untilNext), rewind (rewind-at-end), autoReverse, and accel
/// (none|smooth|decelerate). Each control reopens validator-clean and is surfaced by
/// get and read --view structure.
/// </summary>
public sealed class AnimationTimingTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    /// <summary>Creates a deck with one shape carrying one fade animation; returns the animation path.</summary>
    private string CreateWithFade()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var shape = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "Box"))));
        var shapePath = shape["results"]![0]!["target"]!.GetValue<string>();
        Edit(TestEnv.Op("add", shapePath, type: "animation", props: TestEnv.Props(("effect", "fade"))));
        return "/slide[1]/animation[1]";
    }

    private P.CommonTimeNode EffectNode(PresentationDocument doc)
    {
        var timing = doc.PresentationPart!.SlideParts.Single().Slide!.Timing!;
        return timing.Descendants<P.CommonTimeNode>().First(n =>
            n.NodeType?.Value is { } t &&
            (t == P.TimeNodeValues.ClickEffect || t == P.TimeNodeValues.WithEffect || t == P.TimeNodeValues.AfterEffect));
    }

    [Fact]
    public void SetRepeatCount_WritesRepeatCountInThousandths_ReopenVerified()
    {
        var anim = CreateWithFade();
        Edit(TestEnv.Op("set", anim, props: TestEnv.Props(("repeat", 3))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Equal("3000", EffectNode(doc).RepeatCount!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", anim))));
        Assert.Equal("3", detail["repeat"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetRepeatUntilClick_WritesIndefinite_ReopenVerified()
    {
        var anim = CreateWithFade();
        Edit(TestEnv.Op("set", anim, props: TestEnv.Props(("repeat", "untilClick"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Equal("indefinite", EffectNode(doc).RepeatCount!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", anim))));
        Assert.Equal("indefinite", detail["repeat"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetRepeatNone_ClearsRepeatCount()
    {
        var anim = CreateWithFade();
        Edit(TestEnv.Op("set", anim, props: TestEnv.Props(("repeat", 4))));
        Edit(TestEnv.Op("set", anim, props: TestEnv.Props(("repeat", "none"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        Assert.Null(EffectNode(doc).RepeatCount);
    }

    [Fact]
    public void SetRewind_WritesFillRemove_ReopenVerified()
    {
        var anim = CreateWithFade();
        Edit(TestEnv.Op("set", anim, props: TestEnv.Props(("rewind", true))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Equal(P.TimeNodeFillValues.Remove, EffectNode(doc).Fill!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", anim))));
        Assert.True(detail["rewind"]!.GetValue<bool>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetAutoReverse_WritesAutoRev_ReopenVerified()
    {
        var anim = CreateWithFade();
        Edit(TestEnv.Op("set", anim, props: TestEnv.Props(("autoReverse", true))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.True(EffectNode(doc).AutoReverse!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", anim))));
        Assert.True(detail["autoReverse"]!.GetValue<bool>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetAccelSmooth_WritesAccelAndDecel_ReopenVerified()
    {
        var anim = CreateWithFade();
        Edit(TestEnv.Op("set", anim, props: TestEnv.Props(("accel", "smooth"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var node = EffectNode(doc);
            Assert.Equal(25_000, node.Acceleration!.Value);
            Assert.Equal(25_000, node.Deceleration!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", anim))));
        Assert.Equal("smooth", detail["accel"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetAccelDecelerate_WritesOnlyDecel_ReopenVerified()
    {
        var anim = CreateWithFade();
        Edit(TestEnv.Op("set", anim, props: TestEnv.Props(("accel", "decelerate"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var node = EffectNode(doc);
            Assert.Null(node.Acceleration);
            Assert.Equal(50_000, node.Deceleration!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", anim))));
        Assert.Equal("decelerate", detail["accel"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void TimingControlsCombine_AndSurviveATriggerChange()
    {
        var anim = CreateWithFade();
        // Set timing controls AND change the trigger in the same op — the trigger
        // rebuild reuses the effect par, so the timing attributes must survive.
        Edit(TestEnv.Op("set", anim, props: TestEnv.Props(
            ("repeat", 2), ("rewind", true), ("autoReverse", true), ("trigger", "afterPrevious"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var node = EffectNode(doc);
            Assert.Equal("2000", node.RepeatCount!.Value);
            Assert.Equal(P.TimeNodeFillValues.Remove, node.Fill!.Value);
            Assert.True(node.AutoReverse!.Value);
            Assert.Equal(P.TimeNodeValues.AfterEffect, node.NodeType!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", anim))));
        Assert.Equal("2", detail["repeat"]!.GetValue<string>());
        Assert.True(detail["rewind"]!.GetValue<bool>());
        Assert.True(detail["autoReverse"]!.GetValue<bool>());
        Assert.Equal("afterPrevious", detail["trigger"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Structure_ShowsTimingDetails()
    {
        var anim = CreateWithFade();
        Edit(TestEnv.Op("set", anim, props: TestEnv.Props(("repeat", "untilNext"), ("accel", "smooth"))));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var animation = data["slides"]![0]!["animations"]![0]!.AsObject();
        Assert.Equal("indefinite", animation["repeat"]!.GetValue<string>());
        Assert.Equal("smooth", animation["accel"]!.GetValue<string>());
    }

    [Fact]
    public void BadRepeat_Fails()
    {
        var anim = CreateWithFade();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", anim, props: TestEnv.Props(("repeat", "twice")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void BadAccel_Fails()
    {
        var anim = CreateWithFade();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", anim, props: TestEnv.Props(("accel", "fast")))]),
            ErrorCodes.InvalidArgs);
    }
}
