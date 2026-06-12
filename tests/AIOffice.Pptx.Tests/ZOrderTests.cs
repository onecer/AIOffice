using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>M3 z-order: move front/back/forward/backward over spTree paint order.</summary>
public sealed class ZOrderTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    /// <summary>Creates a deck with three overlapping shapes A, B, C and returns their canonical paths.</summary>
    private (string A, string B, string C) CreateStack()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var added = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "A"), ("fill", "FF0000"))),
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "B"), ("fill", "00FF00"))),
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "C"), ("fill", "0000FF"))),
        ]));
        var results = added["results"]!.AsArray();
        return (
            results[0]!["target"]!.GetValue<string>(),
            results[1]!["target"]!.GetValue<string>(),
            results[2]!["target"]!.GetValue<string>());
    }

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private int ZIndexOf(string path) =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["zIndex"]!.GetValue<int>();

    [Fact]
    public void MoveFront_PaintsLast()
    {
        var (a, b, c) = CreateStack();
        Assert.Equal(1, ZIndexOf(a));

        var data = Edit(TestEnv.Op("move", a, position: "front"));
        Assert.Equal(a, data["results"]![0]!["target"]!.GetValue<string>());

        Assert.Equal(3, ZIndexOf(a));
        Assert.Equal(1, ZIndexOf(b));
        Assert.Equal(2, ZIndexOf(c));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MoveBack_PaintsFirst()
    {
        var (a, b, c) = CreateStack();
        Edit(TestEnv.Op("move", c, position: "back"));

        Assert.Equal(1, ZIndexOf(c));
        Assert.Equal(2, ZIndexOf(a));
        Assert.Equal(3, ZIndexOf(b));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MoveForwardAndBackward_SwapNeighbors()
    {
        var (a, b, c) = CreateStack();

        Edit(TestEnv.Op("move", a, position: "forward"));
        Assert.Equal(2, ZIndexOf(a));
        Assert.Equal(1, ZIndexOf(b));

        Edit(TestEnv.Op("move", c, position: "backward"));
        Assert.Equal(2, ZIndexOf(c));
        Assert.Equal(3, ZIndexOf(a));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MoveAtExtreme_IsANoOp()
    {
        var (a, _, c) = CreateStack();

        Edit(TestEnv.Op("move", a, position: "back"));
        Assert.Equal(1, ZIndexOf(a));

        Edit(TestEnv.Op("move", c, position: "front"));
        Assert.Equal(3, ZIndexOf(c));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SvgPaintOrder_FollowsZOrder()
    {
        var (a, _, _) = CreateStack();
        Edit(TestEnv.Op("move", a, position: "front"));

        var svg = TestEnv.AssertOk(_handler.Render(
            _ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))))["slides"]![0]!["svg"]!.GetValue<string>();

        // SVG paints in document order: A (moved to front) must now come last.
        var posA = svg.IndexOf(">A</text>", StringComparison.Ordinal);
        var posB = svg.IndexOf(">B</text>", StringComparison.Ordinal);
        var posC = svg.IndexOf(">C</text>", StringComparison.Ordinal);
        Assert.True(posB < posC && posC < posA, $"paint order wrong: A@{posA} B@{posB} C@{posC}");
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void StableIdPath_SurvivesZOrderMoves()
    {
        var (a, b, _) = CreateStack();
        Edit(TestEnv.Op("move", a, position: "front"));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", a))));
        Assert.Equal("A", detail["text"]!.GetValue<string>());
        Assert.Equal(a, detail["path"]!.GetValue<string>());

        // The ordinal path shifted, but the id path still resolves shape B too.
        Assert.Equal("B", TestEnv.AssertOk(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", b))))["text"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void UnknownPosition_IsInvalidArgsWithCandidates()
    {
        var (a, _, _) = CreateStack();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("move", a, position: "top")]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Equal(["front", "back", "forward", "backward"], error.Candidates!);
    }

    [Fact]
    public void ChartFrame_MovesInZOrderViaChartPath()
    {
        var (_, _, _) = CreateStack();
        Edit(TestEnv.Op("add", "/slide[1]", type: "chart", props: TestEnv.Props(
            ("kind", "pie"),
            ("categories", new JsonArray("X", "Y")),
            ("series", new JsonArray(new JsonObject { ["values"] = new JsonArray(1, 2) })))));

        var data = Edit(TestEnv.Op("move", "/slide[1]/chart[1]", position: "back"));
        var target = data["results"]![0]!["target"]!.GetValue<string>();
        Assert.StartsWith("/slide[1]/shape[@id=", target, StringComparison.Ordinal);
        Assert.Equal(1, ZIndexOf(target));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
