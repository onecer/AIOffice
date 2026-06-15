using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.4.0 animation trigger-on-object-click: triggerOn:"@M" makes an effect play when shape M is
/// clicked. The effect par lives in an interactive p:seq (nodeType=interactiveSeq) keyed to M's
/// onClick, not the main sequence. Validator-clean; read/get surface the trigger.
/// </summary>
public sealed class AnimationTriggerTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    /// <summary>Creates deck.pptx with two shapes and returns their (animatedPath, triggerPath, triggerId).</summary>
    private (string Animated, string Trigger, uint TriggerId) CreateWithTwoShapes()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var a = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "Target"))));
        var b = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "Button"))));
        var animated = a["results"]![0]!["target"]!.GetValue<string>();
        var trigger = b["results"]![0]!["target"]!.GetValue<string>();
        var triggerId = uint.Parse(trigger.Split("@id=")[1].TrimEnd(']'));
        return (animated, trigger, triggerId);
    }

    private P.Timing Slide1Timing(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.Single().Slide!.Timing!;

    [Fact]
    public void TriggerOn_BuildsInteractiveSequence_KeyedToTheTriggerShapesOnClick_AndValidates()
    {
        var (animated, _, triggerId) = CreateWithTwoShapes();
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(
            ("effect", "fade"), ("triggerOn", $"@{triggerId}"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = Slide1Timing(doc);

            // There is an interactive sequence (alongside the main sequence) keyed to the trigger.
            var interactive = timing.Descendants<P.SequenceTimeNode>()
                .Single(s => s.CommonTimeNode?.NodeType?.Value == P.TimeNodeValues.InteractiveSequence);

            var startCondition = interactive.CommonTimeNode!
                .GetFirstChild<P.StartConditionList>()!.GetFirstChild<P.Condition>()!;
            Assert.Equal(P.TriggerEventValues.OnClick, startCondition.Event!.Value);
            Assert.Equal(
                triggerId.ToString(),
                startCondition.TargetElement!.GetFirstChild<P.ShapeTarget>()!.ShapeId!.Value);

            // The fade effect targets the animated shape (not the trigger).
            var animatedId = animated.Split("@id=")[1].TrimEnd(']');
            var animEffect = interactive.Descendants<P.AnimateEffect>().Single();
            Assert.Equal("fade", animEffect.Filter!.Value);
            Assert.Equal(animatedId, animEffect.Descendants<P.ShapeTarget>().First().ShapeId!.Value);

            // The main sequence still exists and is empty of effects (only the interactive one fires).
            var mainSeq = timing.Descendants<P.SequenceTimeNode>()
                .Single(s => s.CommonTimeNode?.NodeType?.Value == P.TimeNodeValues.MainSequence);
            Assert.Empty(mainSeq.Descendants<P.AnimateEffect>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_TriggeredAnimation_ReportsTriggerOnPathAndShapeId()
    {
        var (animated, trigger, triggerId) = CreateWithTwoShapes();
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(
            ("effect", "appear"), ("triggerOn", $"@{triggerId}"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[1]"))));
        Assert.Equal(animated, detail["target"]!.GetValue<string>());
        Assert.Equal(trigger, detail["triggerOn"]!.GetValue<string>());
        Assert.Equal(triggerId, detail["triggerShapeId"]!.GetValue<uint>());
    }

    [Fact]
    public void Structure_ShowsTheTrigger()
    {
        var (animated, trigger, triggerId) = CreateWithTwoShapes();
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(
            ("effect", "fade"), ("triggerOn", $"@{triggerId}"))));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var anim = data["slides"]![0]!["animations"]![0]!;
        Assert.Equal("fade", anim["effect"]!.GetValue<string>());
        Assert.Equal(trigger, anim["triggerOn"]!.GetValue<string>());
    }

    [Fact]
    public void TwoEffects_SameTrigger_ShareOneInteractiveSequence()
    {
        var (animated, _, triggerId) = CreateWithTwoShapes();
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(
            ("effect", "fade"), ("triggerOn", $"@{triggerId}"))));
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(
            ("effect", "spin"), ("triggerOn", $"@{triggerId}"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var interactiveSeqs = Slide1Timing(doc).Descendants<P.SequenceTimeNode>()
                .Where(s => s.CommonTimeNode?.NodeType?.Value == P.TimeNodeValues.InteractiveSequence)
                .ToList();
            Assert.Single(interactiveSeqs); // one seq, two effects inside it
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void TriggerOn_CoexistsWithAMainSequenceAnimation_AndStaysValid()
    {
        var (animated, _, triggerId) = CreateWithTwoShapes();
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(("effect", "appear"))));
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(
            ("effect", "fade"), ("triggerOn", $"@{triggerId}"))));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        Assert.Equal(2, data["slides"]![0]!["animations"]!.AsArray().Count);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void TriggerOn_WorksOnMotionPath()
    {
        var (animated, _, triggerId) = CreateWithTwoShapes();
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(
            ("effect", "motionPath"), ("path", "line"), ("triggerOn", $"@{triggerId}"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var interactive = Slide1Timing(doc).Descendants<P.SequenceTimeNode>()
                .Single(s => s.CommonTimeNode?.NodeType?.Value == P.TimeNodeValues.InteractiveSequence);
            Assert.Single(interactive.Descendants<P.AnimateMotion>());
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[1]"))));
        Assert.Equal("motionPath", detail["effect"]!.GetValue<string>());
        Assert.Equal(triggerId, detail["triggerShapeId"]!.GetValue<uint>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Remove_TriggeredAnimation_PrunesTheInteractiveSequence_AndStaysValid()
    {
        var (animated, _, triggerId) = CreateWithTwoShapes();
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(
            ("effect", "fade"), ("triggerOn", $"@{triggerId}"))));

        Edit(TestEnv.Op("remove", "/slide[1]/animation[1]"));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = doc.PresentationPart!.SlideParts.Single().Slide!.Timing;
            // No interactive sequence remains (the empty seq was pruned).
            var interactiveSeqs = (timing?.Descendants<P.SequenceTimeNode>() ?? [])
                .Where(s => s.CommonTimeNode?.NodeType?.Value == P.TimeNodeValues.InteractiveSequence)
                .ToList();
            Assert.Empty(interactiveSeqs);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ----- guards -----------------------------------------------------------------

    [Fact]
    public void TriggerOn_UnknownShape_IsInvalidPath_WithShapeIdCandidates()
    {
        var (animated, _, _) = CreateWithTwoShapes();
        var error = TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", animated, type: "animation",
                props: TestEnv.Props(("effect", "fade"), ("triggerOn", "@999")))]),
            ErrorCodes.InvalidPath);
        Assert.NotEmpty(error.Candidates!);
        Assert.All(error.Candidates!, c => Assert.StartsWith("@", c, StringComparison.Ordinal));
    }

    [Fact]
    public void TriggerOn_TheAnimatedShapeItself_IsInvalidArgs()
    {
        var (animated, _, _) = CreateWithTwoShapes();
        var animatedId = animated.Split("@id=")[1].TrimEnd(']');
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", animated, type: "animation",
                props: TestEnv.Props(("effect", "fade"), ("triggerOn", $"@{animatedId}")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void TriggerOn_BadFormat_IsInvalidArgs()
    {
        var (animated, _, _) = CreateWithTwoShapes();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", animated, type: "animation",
                props: TestEnv.Props(("effect", "fade"), ("triggerOn", "button")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void SetTrigger_OnTriggeredAnimation_IsInvalidArgs()
    {
        var (animated, _, triggerId) = CreateWithTwoShapes();
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(
            ("effect", "fade"), ("triggerOn", $"@{triggerId}"))));

        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", "/slide[1]/animation[1]",
                props: TestEnv.Props(("trigger", "withPrevious")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void SetDuration_OnTriggeredAnimation_StillWorks()
    {
        var (animated, _, triggerId) = CreateWithTwoShapes();
        Edit(TestEnv.Op("add", animated, type: "animation", props: TestEnv.Props(
            ("effect", "fade"), ("triggerOn", $"@{triggerId}"))));

        Edit(TestEnv.Op("set", "/slide[1]/animation[1]", props: TestEnv.Props(("duration", "1.5s"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[1]"))));
        Assert.Equal("1.5s", detail["duration"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
