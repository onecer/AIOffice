using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

public sealed class ReadTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void CreateThreeSlideDeck()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Alpha"))));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("title", "Beta"))),
            TestEnv.Op("add", "/slide[3]", type: "slide", props: TestEnv.Props(("title", "Gamma"))),
            TestEnv.Op("add", "/slide[2]", type: "shape", props: TestEnv.Props(
                ("text", JsonValue.Create("two words")),
                ("x", JsonValue.Create(2)),
                ("y", JsonValue.Create(4)),
                ("w", JsonValue.Create(8)),
                ("h", JsonValue.Create(2)))),
        ]));
    }

    [Fact]
    public void Outline_ListsSlidesAndShapeTexts()
    {
        CreateThreeSlideDeck();
        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));

        Assert.Equal("outline", data["view"]!.GetValue<string>());
        var slides = data["slides"]!.AsArray();
        Assert.Equal(3, slides.Count);
        Assert.Equal("/slide[2]", slides[1]!["path"]!.GetValue<string>());
        Assert.Equal(2, slides[1]!["shapeCount"]!.GetValue<int>());
        Assert.Equal("Beta", slides[1]!["shapes"]![0]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void TextView_JoinsAllSlideText()
    {
        CreateThreeSlideDeck();
        var text = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))))["text"]!.GetValue<string>();

        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("two words", text, StringComparison.Ordinal);
        Assert.Contains("Gamma", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Stats_CountSlidesShapesWordsCharacters()
    {
        CreateThreeSlideDeck();
        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "stats"))));

        Assert.Equal(3, data["slides"]!.GetValue<int>());
        Assert.Equal(4, data["shapes"]!.GetValue<int>());
        Assert.Equal(5, data["words"]!.GetValue<int>()); // Alpha + Beta + Gamma + "two words"
    }

    [Fact]
    public void Structure_ReportsSlideSizeAndShapeGeometry()
    {
        CreateThreeSlideDeck();
        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));

        Assert.Equal(33.87, data["slideWidthCm"]!.GetValue<double>());
        Assert.Equal(19.05, data["slideHeightCm"]!.GetValue<double>());

        var box = data["slides"]!.AsArray()[1]!["shapes"]!.AsArray()[1]!;
        Assert.Equal(2.0, box["x"]!.GetValue<double>());
        Assert.Equal(4.0, box["y"]!.GetValue<double>());
        Assert.Equal(8.0, box["w"]!.GetValue<double>());
        Assert.Equal(2.0, box["h"]!.GetValue<double>());
    }

    [Fact]
    public void Range_LimitsSlidesButKeepsRealIndices()
    {
        CreateThreeSlideDeck();
        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"), ("range", "2..3"))));

        var slides = data["slides"]!.AsArray();
        Assert.Equal(2, slides.Count);
        Assert.Equal("/slide[2]", slides[0]!["path"]!.GetValue<string>());
        Assert.Equal("/slide[3]", slides[1]!["path"]!.GetValue<string>());
    }

    [Fact]
    public void UnknownView_FailsWithCandidates()
    {
        CreateThreeSlideDeck();
        var envelope = _handler.Read(_ws.Ctx("deck.pptx", ("view", "summary")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("outline", error.Candidates!);
    }
}
