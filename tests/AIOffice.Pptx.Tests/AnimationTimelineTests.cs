using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// M6 animation timeline editing: reorder (move before/after another animation)
/// and retime (set trigger/delay/duration) keep the p:timing tree valid and the
/// /slide[i]/animation[k] indices consistent.
/// </summary>
public sealed class AnimationTimelineTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    /// <summary>Creates a deck with one shape and returns its canonical path.</summary>
    private string CreateWithShape()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "animate me"))),
        ]));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private string Effect(int index) =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", $"/slide[1]/animation[{index}]"))))
            ["effect"]!.GetValue<string>();

    // ---- retime --------------------------------------------------------------

    [Fact]
    public void SetDuration_RetimesTheEffect_ReopenVerified()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"), ("duration", "0.5s"))));

        Edit(TestEnv.Op("set", "/slide[1]/animation[1]", props: TestEnv.Props(("duration", "1.2s"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[1]"))));
        Assert.Equal("1.2s", detail["duration"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetDelay_RewritesTheStartCondition_ReopenVerified()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));

        Edit(TestEnv.Op("set", "/slide[1]/animation[1]", props: TestEnv.Props(("delay", "0.2s"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[1]"))));
        Assert.Equal("0.2s", detail["delay"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetTrigger_ToWithPrevious_RewritesNodeTypeAndRegroups()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "wipe"), ("trigger", "click"))));

        // Two click groups initially; making the second withPrevious folds it into the first group.
        Edit(TestEnv.Op("set", "/slide[1]/animation[2]", props: TestEnv.Props(("trigger", "withPrevious"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[2]"))));
        Assert.Equal("withPrevious", detail["trigger"]!.GetValue<string>());
        Assert.Equal("wipe", detail["effect"]!.GetValue<string>()); // order preserved

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var timing = doc.PresentationPart!.SlideParts.Single().Slide!.Timing!;
            var mainSeqChildren = timing.Descendants<P.SequenceTimeNode>().Single()
                .CommonTimeNode!.GetFirstChild<P.ChildTimeNodeList>()!;
            Assert.Single(mainSeqChildren.Elements<P.ParallelTimeNode>()); // one group now
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetTrigger_AndDelayAndDuration_InOneOp_AllApply()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "wipe"), ("trigger", "click"))));

        Edit(TestEnv.Op("set", "/slide[1]/animation[2]", props: TestEnv.Props(
            ("trigger", "afterPrevious"), ("delay", "0.3s"), ("duration", "0.9s"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[2]"))));
        Assert.Equal("afterPrevious", detail["trigger"]!.GetValue<string>());
        Assert.Equal("0.3s", detail["delay"]!.GetValue<string>());
        Assert.Equal("0.9s", detail["duration"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetDurationOnAppear_IsInvalidArgs()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "appear"))));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/slide[1]/animation[1]", props: TestEnv.Props(("duration", "1s"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void SetUnknownAnimationProp_IsInvalidArgs()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/slide[1]/animation[1]", props: TestEnv.Props(("effect", "wipe"))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("trigger", error.Candidates!);
    }

    [Fact]
    public void SetUnknownTrigger_IsInvalidArgs()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/slide[1]/animation[1]", props: TestEnv.Props(("trigger", "hover"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- reorder -------------------------------------------------------------

    [Fact]
    public void MoveAnimation_BeforeAnother_ReordersTheTimeline()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "wipe"))));
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "spin"))));
        Assert.Equal(["fade", "wipe", "spin"], new[] { Effect(1), Effect(2), Effect(3) });

        // Move spin (3) before fade (1): order becomes spin, fade, wipe.
        var data = Edit(TestEnv.Op("move", "/slide[1]/animation[3]", position: "before /slide[1]/animation[1]"));
        Assert.Equal("/slide[1]/animation[1]", data["results"]![0]!["target"]!.GetValue<string>());
        Assert.Equal(["spin", "fade", "wipe"], new[] { Effect(1), Effect(2), Effect(3) });
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MoveAnimation_AfterAnother_ReordersTheTimeline()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "wipe"))));
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "spin"))));

        // Move fade (1) after wipe (2): order becomes wipe, fade, spin.
        var data = Edit(TestEnv.Op("move", "/slide[1]/animation[1]", position: "after /slide[1]/animation[2]"));
        Assert.Equal("/slide[1]/animation[2]", data["results"]![0]!["target"]!.GetValue<string>());
        Assert.Equal(["wipe", "fade", "spin"], new[] { Effect(1), Effect(2), Effect(3) });
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MoveAnimation_PreservesEachTriggerWhenRegrouping()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "wipe"), ("trigger", "withPrevious"))));
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "spin"), ("trigger", "click"))));

        // Move spin to the front; it keeps its click trigger and opens the first group.
        Edit(TestEnv.Op("move", "/slide[1]/animation[3]", position: "before /slide[1]/animation[1]"));

        var first = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[1]"))));
        Assert.Equal("spin", first["effect"]!.GetValue<string>());
        Assert.Equal("click", first["trigger"]!.GetValue<string>());

        var second = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[2]"))));
        Assert.Equal("fade", second["effect"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MoveAnimation_NoOpWhenAnchoredToItself()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "wipe"))));

        var data = Edit(TestEnv.Op("move", "/slide[1]/animation[1]", position: "before /slide[1]/animation[1]"));
        Assert.Equal("/slide[1]/animation[1]", data["results"]![0]!["target"]!.GetValue<string>());
        Assert.Equal(["fade", "wipe"], new[] { Effect(1), Effect(2) });
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MoveAnimation_WithoutAnchor_IsInvalidArgs()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("move", "/slide[1]/animation[1]", position: "1"),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void MoveAnimation_AnchorOnAnotherSlide_IsInvalidArgs()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"))));
        Edit(TestEnv.Op("add", "/slide[2]", type: "slide"));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("move", "/slide[1]/animation[1]", position: "before /slide[2]/animation[1]"),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Structure_ShowsTheOrderedTimelineWithTimings()
    {
        var shape = CreateWithShape();
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade"), ("duration", "0.5s"))));
        Edit(TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(
            ("effect", "wipe"), ("trigger", "afterPrevious"), ("delay", "0.25s"))));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var animations = data["slides"]![0]!["animations"]!.AsArray();
        Assert.Equal(2, animations.Count);
        Assert.Equal(1, animations[0]!["index"]!.GetValue<int>());
        Assert.Equal("0.5s", animations[0]!["duration"]!.GetValue<string>());
        Assert.Equal("afterPrevious", animations[1]!["trigger"]!.GetValue<string>());
        Assert.Equal("0.25s", animations[1]!["delay"]!.GetValue<string>());
    }
}
