using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>M4 entrance animations: p:timing tree with PowerPoint preset class/id/subtype.</summary>
public sealed class AnimationTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    /// <summary>Creates deck.pptx with one textbox on slide 1 and returns its canonical shape path.</summary>
    private string CreateWithShape()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "Hello"))),
        ]));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private JsonObject AddAnimation(string shapePath, params (string Key, JsonNode? Value)[] props) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", shapePath, type: "animation", props: TestEnv.Props(props))]));

    private P.Timing Slide1Timing(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.Single().Slide!.Timing!;

    [Fact]
    public void AddAppear_BuildsTimingTree_WithVisibilitySetOnly()
    {
        var shapePath = CreateWithShape();
        var data = AddAnimation(shapePath, ("effect", "appear"));
        Assert.Equal("/slide[1]/animation[1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var root = timing.Descendants<P.CommonTimeNode>().First();
            Assert.Equal(P.TimeNodeValues.TmingRoot, root.NodeType!.Value);

            var mainSeq = timing.Descendants<P.SequenceTimeNode>().Single();
            Assert.Equal(P.TimeNodeValues.MainSequence, mainSeq.CommonTimeNode!.NodeType!.Value);

            var effect = timing.Descendants<P.CommonTimeNode>()
                .Single(c => c.NodeType?.Value == P.TimeNodeValues.ClickEffect);
            Assert.Equal(1, effect.PresetId!.Value);
            Assert.Equal(P.TimeNodePresetClassValues.Entrance, effect.PresetClass!.Value);

            var set = timing.Descendants<P.SetBehavior>().Single();
            Assert.Equal("style.visibility", set.Descendants<P.AttributeName>().Single().Text);
            Assert.Equal("visible", set.Descendants<P.StringVariantValue>().Single().Val!.Value);
            Assert.Empty(timing.Descendants<P.AnimateEffect>());
            Assert.Empty(timing.Descendants<P.Animate>());

            // The effect targets the shape by its stable id.
            var shapeId = shapePath.Split("@id=")[1].TrimEnd(']');
            Assert.Equal(shapeId, timing.Descendants<P.ShapeTarget>().First().ShapeId!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddFade_WritesAnimEffectFilterFade_WithDuration()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "fade"), ("duration", "0.75s"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var effect = timing.Descendants<P.CommonTimeNode>()
                .Single(c => c.NodeType?.Value == P.TimeNodeValues.ClickEffect);
            Assert.Equal(10, effect.PresetId!.Value);

            var animEffect = timing.Descendants<P.AnimateEffect>().Single();
            Assert.Equal("fade", animEffect.Filter!.Value);
            Assert.Equal(P.AnimateEffectTransitionValues.In, animEffect.Transition!.Value);
            Assert.Equal(
                "750",
                animEffect.GetFirstChild<P.CommonBehavior>()!.GetFirstChild<P.CommonTimeNode>()!.Duration!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddFlyIn_FromLeft_WritesSubtypeAndPositionAnims()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "flyIn"), ("direction", "left"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var effect = timing.Descendants<P.CommonTimeNode>()
                .Single(c => c.NodeType?.Value == P.TimeNodeValues.ClickEffect);
            Assert.Equal(2, effect.PresetId!.Value);
            Assert.Equal(8, effect.PresetSubtype!.Value); // from left

            var anims = timing.Descendants<P.Animate>().ToList();
            Assert.Equal(2, anims.Count); // ppt_x and ppt_y
            Assert.Equal(
                ["ppt_x", "ppt_y"],
                anims.Select(a => a.Descendants<P.AttributeName>().Single().Text));
            Assert.Contains(
                "0-#ppt_w/2",
                anims[0].Descendants<P.StringVariantValue>().Select(v => v.Val!.Value));
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddWipe_FromBottom_WritesWipeUpFilter()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "wipe"), ("direction", "bottom"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var effect = timing.Descendants<P.CommonTimeNode>()
                .Single(c => c.NodeType?.Value == P.TimeNodeValues.ClickEffect);
            Assert.Equal(22, effect.PresetId!.Value);
            Assert.Equal(1, effect.PresetSubtype!.Value);
            Assert.Equal("wipe(up)", timing.Descendants<P.AnimateEffect>().Single().Filter!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Triggers_MapToNodeTypes_AndClickStartsANewGroup()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "appear"));
        AddAnimation(shapePath, ("effect", "fade"), ("trigger", "withPrevious"));
        AddAnimation(shapePath, ("effect", "wipe"), ("trigger", "afterPrevious"), ("delay", "0.25s"));
        AddAnimation(shapePath, ("effect", "fade"), ("trigger", "click"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var effects = timing.Descendants<P.CommonTimeNode>()
                .Where(c => c.NodeType?.Value is { } t &&
                            (t == P.TimeNodeValues.ClickEffect ||
                             t == P.TimeNodeValues.WithEffect ||
                             t == P.TimeNodeValues.AfterEffect))
                .ToList();
            Assert.Equal(4, effects.Count);
            Assert.Equal(P.TimeNodeValues.ClickEffect, effects[0].NodeType!.Value);
            Assert.Equal(P.TimeNodeValues.WithEffect, effects[1].NodeType!.Value);
            Assert.Equal(P.TimeNodeValues.AfterEffect, effects[2].NodeType!.Value);
            Assert.Equal(P.TimeNodeValues.ClickEffect, effects[3].NodeType!.Value);

            // Two click groups under mainSeq: [appear+with+after] and [fade].
            var mainSeqChildren = timing.Descendants<P.SequenceTimeNode>().Single()
                .CommonTimeNode!.GetFirstChild<P.ChildTimeNodeList>()!;
            Assert.Equal(2, mainSeqChildren.Elements<P.ParallelTimeNode>().Count());

            // The afterPrevious effect carries its delay in ms.
            Assert.Equal(
                "250",
                effects[2].GetFirstChild<P.StartConditionList>()!.GetFirstChild<P.Condition>()!.Delay!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[3]"))));
        Assert.Equal("afterPrevious", detail["trigger"]!.GetValue<string>());
        Assert.Equal("0.25s", detail["delay"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_AnimationPath_ReportsTargetEffectAndTrigger()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "flyIn"), ("direction", "top"), ("duration", "1s"));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[1]"))));
        Assert.Equal("/slide[1]/animation[1]", detail["path"]!.GetValue<string>());
        Assert.Equal(shapePath, detail["target"]!.GetValue<string>());
        Assert.Equal("flyIn", detail["effect"]!.GetValue<string>());
        Assert.Equal("click", detail["trigger"]!.GetValue<string>());
        Assert.Equal("1s", detail["duration"]!.GetValue<string>());
        Assert.Equal("top", detail["direction"]!.GetValue<string>());
    }

    [Fact]
    public void Structure_ListsAnimationsPerSlide()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "fade"));
        AddAnimation(shapePath, ("effect", "wipe"), ("trigger", "withPrevious"));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var animations = data["slides"]![0]!["animations"]!.AsArray();
        Assert.Equal(2, animations.Count);
        Assert.Equal("/slide[1]/animation[1]", animations[0]!["path"]!.GetValue<string>());
        Assert.Equal(shapePath, animations[0]!["target"]!.GetValue<string>());
        Assert.Equal("fade", animations[0]!["effect"]!.GetValue<string>());
        Assert.Equal("withPrevious", animations[1]!["trigger"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_MiddleAnimation_RelinksIndices_AndStaysValid()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "appear"));
        AddAnimation(shapePath, ("effect", "fade"), ("trigger", "click"));
        AddAnimation(shapePath, ("effect", "wipe"), ("trigger", "click"));

        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("remove", "/slide[1]/animation[2]")]));
        Assert.Equal("/slide[1]/animation[2]", data["results"]![0]!["target"]!.GetValue<string>());

        // The former third animation is now animation[2]; the seq re-linked cleanly.
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[2]"))));
        Assert.Equal("wipe", detail["effect"]!.GetValue<string>());
        TestEnv.AssertFail(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[3]"))),
            ErrorCodes.InvalidPath);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Remove_LastAnimation_DropsTheWholeTimingTree()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "fade"));

        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/slide[1]/animation[1]")]));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Null(doc.PresentationPart!.SlideParts.Single().Slide!.Timing);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void UnknownEffect_IsTypedUnsupported_ListingAllEleven()
    {
        var shapePath = CreateWithShape();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", shapePath, type: "animation", props: TestEnv.Props(("effect", "bounce")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Equal(
            ["appear", "fade", "flyIn", "wipe", "pulse", "grow", "spin", "colorPulse", "fadeOut", "flyOut", "wipeOut"],
            error.Candidates!);
    }

    [Fact]
    public void AddAnimation_OnSlidePath_IsInvalidArgs()
    {
        CreateWithShape();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "animation", props: TestEnv.Props(("effect", "fade")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Direction_OnFade_IsInvalidArgs()
    {
        var shapePath = CreateWithShape();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "add", shapePath, type: "animation",
            props: TestEnv.Props(("effect", "fade"), ("direction", "left")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }
}
