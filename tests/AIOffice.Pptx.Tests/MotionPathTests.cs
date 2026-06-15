using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// Motion-path animations (v1.3.0, additive "path" preset class): an
/// <c>{"effect":"motionPath"}</c> op builds a <c>p:animMotion</c> in the timing
/// tree with the right path string, surfaces in read --view structure / get, and
/// stays validator-clean across line/arc/circle/custom paths.
/// </summary>
public sealed class MotionPathTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    /// <summary>Creates a deck with one box shape (id 2) and returns its shape path.</summary>
    private string CreateWithShape()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("text", "Mover"), ("x", "2cm"), ("y", "2cm"), ("w", "4cm"), ("h", "2cm")))]));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private JsonObject AddMotion(string shapePath, params (string Key, JsonNode? Value)[] props)
    {
        var p = TestEnv.Props(("effect", "motionPath"));
        foreach (var (key, value) in props)
        {
            p[key] = value;
        }

        return TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", shapePath, type: "animation", props: p)]));
    }

    private P.AnimateMotion SingleMotion()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        return doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.AnimateMotion>().Single();
    }

    [Fact]
    public void Line_WritesAnimMotionWithRelativePathAndValidates()
    {
        var shape = CreateWithShape();
        AddMotion(shape, ("path", "line"), ("direction", "right"), ("duration", "2s"));

        var motion = SingleMotion();
        Assert.Equal("M 0 0 L 0.3 0 E", motion.Path!.Value);
        Assert.Equal(P.AnimateMotionPathEditModeValues.Relative, motion.PathEditMode!.Value);
        Assert.Equal("2", motion.Descendants<P.ShapeTarget>().Single().ShapeId!.Value); // targets shape id 2

        // The hosting cTn is in the "path" preset class.
        var node = motion.Ancestors<P.CommonTimeNode>().First(c => c.PresetClass is not null);
        Assert.Equal(P.TimeNodePresetClassValues.Path, node.PresetClass!.Value);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("left", "M 0 0 L -0.3 0 E")]
    [InlineData("top", "M 0 0 L 0 -0.3 E")]
    [InlineData("bottom", "M 0 0 L 0 0.3 E")]
    public void Line_HonorsDirection(string direction, string expected)
    {
        var shape = CreateWithShape();
        AddMotion(shape, ("path", "line"), ("direction", direction));
        Assert.Equal(expected, SingleMotion().Path!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Arc_WritesQuadraticPathAndValidates()
    {
        var shape = CreateWithShape();
        AddMotion(shape, ("path", "arc"), ("direction", "top"));

        var motion = SingleMotion();
        Assert.Contains("Q", motion.Path!.Value, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Circle_WritesClosedLoopAndValidates()
    {
        var shape = CreateWithShape();
        AddMotion(shape, ("path", "circle"));

        var motion = SingleMotion();
        Assert.Contains("Z", motion.Path!.Value, StringComparison.Ordinal); // a closed loop
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Custom_FromFractionPoints_BuildsMoveAndLineSegments()
    {
        var shape = CreateWithShape();
        AddMotion(shape, ("path", "custom"), ("points", new JsonArray(
            new JsonArray(0, 0), new JsonArray(0.5, 0.2), new JsonArray(0.5, -0.1))));

        var path = SingleMotion().Path!.Value!;
        Assert.StartsWith("M 0 0 L 0.5 0.2 L 0.5 -0.1", path, StringComparison.Ordinal);
        Assert.EndsWith("E", path, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Custom_FromLengthPoints_ConvertsToSlideFractions()
    {
        var shape = CreateWithShape();
        // The default slide is 33.867cm wide; 16.9335cm is exactly 0.5 of the width.
        AddMotion(shape, ("path", "custom"), ("points", new JsonArray(
            new JsonArray("0cm", "0cm"), new JsonArray("16.9335cm", "0cm"))));

        var path = SingleMotion().Path!.Value!;
        Assert.Contains("L 0.5 0", path, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Triggers_AreHonored()
    {
        var shape = CreateWithShape();
        AddMotion(shape, ("path", "line"), ("trigger", "afterPrevious"));
        var node = SingleMotion().Ancestors<P.CommonTimeNode>().First(c => c.PresetClass is not null);
        Assert.Equal(P.TimeNodeValues.AfterEffect, node.NodeType!.Value);
    }

    [Fact]
    public void Get_ReportsMotionPathClassEffectAndKind()
    {
        var shape = CreateWithShape();
        AddMotion(shape, ("path", "arc"), ("direction", "right"));

        var data = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/animation[1]"))));
        Assert.Equal("path", data["class"]!.GetValue<string>());
        Assert.Equal("motionPath", data["effect"]!.GetValue<string>());
        Assert.Equal("arc", data["direction"]!.GetValue<string>()); // path kind in the direction slot
    }

    [Fact]
    public void Structure_ListsMotionAmongAnimations()
    {
        var shape = CreateWithShape();
        AddMotion(shape, ("path", "circle"));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var animations = data["slides"]![0]!["animations"]!.AsArray();
        Assert.Contains(animations, a => a!["effect"]!.GetValue<string>() == "motionPath");
    }

    [Fact]
    public void Remove_DropsMotionAndKeepsValid()
    {
        var shape = CreateWithShape();
        AddMotion(shape, ("path", "line"));
        AddMotion(shape, ("path", "arc"));

        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("remove", "/slide[1]/animation[1]")]));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Single(doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.AnimateMotion>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Coexists_WithEntranceAnimations()
    {
        var shape = CreateWithShape();
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(("effect", "fade")))]));
        AddMotion(shape, ("path", "line"));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var animations = data["slides"]![0]!["animations"]!.AsArray();
        Assert.Equal(2, animations.Count);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void UnknownPath_IsTypedUnsupported()
    {
        var shape = CreateWithShape();
        var error = TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(
                    ("effect", "motionPath"), ("path", "spiral")))]),
            ErrorCodes.UnsupportedFeature);
        Assert.Contains("line", error.Suggestion, StringComparison.Ordinal);
        Assert.Equal(["line", "arc", "circle", "custom"], error.Candidates!);
    }

    [Fact]
    public void Custom_WithoutPoints_IsInvalidArgs()
    {
        var shape = CreateWithShape();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(
                    ("effect", "motionPath"), ("path", "custom")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Direction_OnCircle_IsInvalidArgs()
    {
        var shape = CreateWithShape();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("add", shape, type: "animation", props: TestEnv.Props(
                    ("effect", "motionPath"), ("path", "circle"), ("direction", "left")))]),
            ErrorCodes.InvalidArgs);
    }
}
