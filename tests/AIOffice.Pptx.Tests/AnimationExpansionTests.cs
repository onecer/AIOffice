using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>M5 animation expansion: emphasis (emph) and exit preset classes.</summary>
public sealed class AnimationExpansionTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private string CreateWithShape()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "Hello"), ("fill", "4472C4"))),
        ]));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private JsonObject AddAnimation(string shapePath, params (string Key, JsonNode? Value)[] props) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", shapePath, type: "animation", props: TestEnv.Props(props))]));

    private P.Timing Slide1Timing(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.Single().Slide!.Timing!;

    private static P.CommonTimeNode EffectNode(P.Timing timing) => timing.Descendants<P.CommonTimeNode>()
        .Single(c => c.NodeType?.Value is { } t &&
                     (t == P.TimeNodeValues.ClickEffect ||
                      t == P.TimeNodeValues.WithEffect ||
                      t == P.TimeNodeValues.AfterEffect));

    // ----- emphasis -------------------------------------------------------------

    [Fact]
    public void Pulse_IsEmphPreset35_AnimScaleAutoReversing_NoVisibilitySet()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "pulse"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var effect = EffectNode(timing);
            Assert.Equal(P.TimeNodePresetClassValues.Emphasis, effect.PresetClass!.Value);
            Assert.Equal(35, effect.PresetId!.Value);

            var scale = timing.Descendants<P.AnimateScale>().Single();
            Assert.Equal(106_000, scale.ToPosition!.X!.Value);
            Assert.Equal(106_000, scale.ToPosition!.Y!.Value);
            Assert.True(scale.CommonBehavior!.CommonTimeNode!.AutoReverse!.Value);

            // Emphasis never touches visibility.
            Assert.Empty(timing.Descendants<P.SetBehavior>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Grow_IsEmphPreset6_AnimScaleTo150()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "grow"), ("duration", "1s"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var effect = EffectNode(timing);
            Assert.Equal(P.TimeNodePresetClassValues.Emphasis, effect.PresetClass!.Value);
            Assert.Equal(6, effect.PresetId!.Value);

            var scale = timing.Descendants<P.AnimateScale>().Single();
            Assert.Equal(150_000, scale.ToPosition!.X!.Value);
            Assert.Null(scale.CommonBehavior!.CommonTimeNode!.AutoReverse);
            Assert.Equal("1000", scale.CommonBehavior!.CommonTimeNode!.Duration!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Spin_IsEmphPreset8_AnimRotFullTurnOnR()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "spin"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            Assert.Equal(8, EffectNode(timing).PresetId!.Value);

            var rotation = timing.Descendants<P.AnimateRotation>().Single();
            Assert.Equal(21_600_000, rotation.By!.Value); // 360 degrees
            Assert.Equal("r", rotation.Descendants<P.AttributeName>().Single().Text);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ColorPulse_IsEmphPreset36_AnimClrToTheGivenColor()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "colorPulse"), ("color", "FF0000"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            Assert.Equal(36, EffectNode(timing).PresetId!.Value);

            var color = timing.Descendants<P.AnimateColor>().Single();
            Assert.Equal(P.AnimateColorSpaceValues.Rgb, color.ColorSpace!.Value);
            Assert.Equal("fillcolor", color.Descendants<P.AttributeName>().Single().Text);
            Assert.Equal("FF0000", color.ToColor!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.True(color.CommonBehavior!.CommonTimeNode!.AutoReverse!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ----- exits ------------------------------------------------------------------

    [Fact]
    public void FadeOut_IsExitPreset10_TransitionOut_ThenHides()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "fadeOut"), ("duration", "0.5s"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var effect = EffectNode(timing);
            Assert.Equal(P.TimeNodePresetClassValues.Exit, effect.PresetClass!.Value);
            Assert.Equal(10, effect.PresetId!.Value);

            var animEffect = timing.Descendants<P.AnimateEffect>().Single();
            Assert.Equal("fade", animEffect.Filter!.Value);
            Assert.Equal(P.AnimateEffectTransitionValues.Out, animEffect.Transition!.Value);

            // The hide set fires at the end of the effect (dur - 1 ms).
            var set = timing.Descendants<P.SetBehavior>().Single();
            Assert.Equal("hidden", set.Descendants<P.StringVariantValue>().Single().Val!.Value);
            Assert.Equal("499", set.GetFirstChild<P.CommonBehavior>()!.GetFirstChild<P.CommonTimeNode>()!
                .GetFirstChild<P.StartConditionList>()!.GetFirstChild<P.Condition>()!.Delay!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void FlyOut_ToRight_AnimatesToOffSlideAndHides()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "flyOut"), ("direction", "right"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var effect = EffectNode(timing);
            Assert.Equal(P.TimeNodePresetClassValues.Exit, effect.PresetClass!.Value);
            Assert.Equal(2, effect.PresetId!.Value);
            Assert.Equal(2, effect.PresetSubtype!.Value); // to the right

            var anims = timing.Descendants<P.Animate>().ToList();
            Assert.Equal(2, anims.Count);
            var xValues = anims[0].Descendants<P.StringVariantValue>().Select(v => v.Val!.Value).ToList();
            Assert.Equal(["#ppt_x", "1+#ppt_w/2"], xValues); // from here, to off-slide

            var set = timing.Descendants<P.SetBehavior>().Single();
            Assert.Equal("hidden", set.Descendants<P.StringVariantValue>().Single().Val!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void WipeOut_ToLeft_WritesSubtypeAndOutFilter()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "wipeOut"), ("direction", "left"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var effect = EffectNode(timing);
            Assert.Equal(22, effect.PresetId!.Value);
            Assert.Equal(2, effect.PresetSubtype!.Value);

            var animEffect = timing.Descendants<P.AnimateEffect>().Single();
            Assert.Equal("wipe(left)", animEffect.Filter!.Value);
            Assert.Equal(P.AnimateEffectTransitionValues.Out, animEffect.Transition!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ----- mixed sequences -----------------------------------------------------------

    [Fact]
    public void MixedClasses_SerializeInOrder_AndReadBackWithClasses()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "fade"));
        AddAnimation(shapePath, ("effect", "pulse"), ("trigger", "withPrevious"));
        AddAnimation(shapePath, ("effect", "fadeOut"), ("trigger", "afterPrevious"));
        AddAnimation(shapePath, ("effect", "spin"), ("trigger", "click"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);
            var effects = timing.Descendants<P.CommonTimeNode>()
                .Where(c => c.PresetClass is not null)
                .ToList();
            Assert.Equal(4, effects.Count);
            Assert.Equal(P.TimeNodePresetClassValues.Entrance, effects[0].PresetClass!.Value);
            Assert.Equal(P.TimeNodePresetClassValues.Emphasis, effects[1].PresetClass!.Value);
            Assert.Equal(P.TimeNodePresetClassValues.Exit, effects[2].PresetClass!.Value);
            Assert.Equal(P.TimeNodePresetClassValues.Emphasis, effects[3].PresetClass!.Value);

            // Two click groups: [fade + pulse + fadeOut] and [spin].
            var mainSeqChildren = timing.Descendants<P.SequenceTimeNode>().Single()
                .CommonTimeNode!.GetFirstChild<P.ChildTimeNodeList>()!;
            Assert.Equal(2, mainSeqChildren.Elements<P.ParallelTimeNode>().Count());
        }

        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var animations = structure["slides"]![0]!["animations"]!.AsArray();
        Assert.Equal(
            ["entrance", "emphasis", "exit", "emphasis"],
            animations.Select(a => a!["class"]!.GetValue<string>()).ToList());
        Assert.Equal(
            ["fade", "pulse", "fadeOut", "spin"],
            animations.Select(a => a!["effect"]!.GetValue<string>()).ToList());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_ExitAnimation_ReportsClassEffectAndDirection()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "flyOut"), ("direction", "left"), ("duration", "0.5s"));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[1]"))));
        Assert.Equal("exit", detail["class"]!.GetValue<string>());
        Assert.Equal("flyOut", detail["effect"]!.GetValue<string>());
        Assert.Equal("left", detail["direction"]!.GetValue<string>());
        Assert.Equal("0.5s", detail["duration"]!.GetValue<string>());
        Assert.Equal(shapePath, detail["target"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_EmphasisAnimation_PrunesCleanly()
    {
        var shapePath = CreateWithShape();
        AddAnimation(shapePath, ("effect", "pulse"));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/slide[1]/animation[1]")]));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Null(doc.PresentationPart!.SlideParts.Single().Slide!.Timing);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ----- prop validation ---------------------------------------------------------

    [Fact]
    public void Direction_OnPulse_IsInvalidArgs()
    {
        var shapePath = CreateWithShape();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", shapePath, type: "animation",
                props: TestEnv.Props(("effect", "pulse"), ("direction", "left")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Color_OnNonColorPulse_IsInvalidArgs()
    {
        var shapePath = CreateWithShape();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", shapePath, type: "animation",
                props: TestEnv.Props(("effect", "grow"), ("color", "FF0000")))]),
            ErrorCodes.InvalidArgs);
    }
}
