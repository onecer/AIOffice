using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

public sealed class QueryGetTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void CreateDeckWithShapes()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Deck Title"))));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("text", JsonValue.Create("Quarterly Q3 numbers")),
                ("fill", JsonValue.Create("FF0000")))),
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("text", JsonValue.Create("Notes")),
                ("fill", JsonValue.Create("0000FF")))),
        ]));
    }

    [Fact]
    public void Query_ContainsText_ReturnsStableIdPaths()
    {
        CreateDeckWithShapes();
        var data = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape:contains('q3')"))));

        Assert.Equal(1, data["count"]!.GetValue<int>());
        var match = data["matches"]![0]!;
        Assert.Matches(new Regex(@"^/slide\[1\]/shape\[@id=[0-9]+\]$"), match["path"]!.GetValue<string>());
        Assert.Matches(new Regex(@"^/slide\[1\]/shape\[[0-9]+\]$"), match["ordinalPath"]!.GetValue<string>());
        Assert.Contains("Q3", match["text"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void Query_ByFill_FiltersShapes()
    {
        CreateDeckWithShapes();
        var data = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape[fill=FF0000]"))));

        Assert.Equal(1, data["count"]!.GetValue<int>());
        Assert.Contains("Q3", data["matches"]![0]!["text"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void Query_SlideElement_MatchesByContainedText()
    {
        CreateDeckWithShapes();
        var data = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "slide:contains('Notes')"))));

        Assert.Equal(1, data["count"]!.GetValue<int>());
        Assert.Equal("/slide[1]", data["matches"]![0]!["path"]!.GetValue<string>());
        Assert.Equal("slide", data["matches"]![0]!["kind"]!.GetValue<string>());
    }

    [Fact]
    public void Query_UnknownElement_FailsWithCandidates()
    {
        CreateDeckWithShapes();
        var envelope = _handler.Query(_ws.Ctx("deck.pptx", ("selector", "cell[value>100]")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.NotNull(error.Candidates);
        Assert.Contains("shape", error.Candidates!);
    }

    [Fact]
    public void Get_AcceptsOrdinalAndStableIdForms()
    {
        CreateDeckWithShapes();
        var byOrdinal = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/shape[2]"))));
        var idPath = byOrdinal["path"]!.GetValue<string>();
        var byId = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", idPath))));

        Assert.Equal(byOrdinal["id"]!.GetValue<uint>(), byId["id"]!.GetValue<uint>());
        Assert.Equal("Quarterly Q3 numbers", byId["text"]!.GetValue<string>());
        Assert.Equal("FF0000", byId["fill"]!.GetValue<string>());
        Assert.Equal(10.0, byId["w"]!.GetValue<double>()); // default width 10cm
    }

    [Fact]
    public void Get_Slide_SummarizesShapes()
    {
        CreateDeckWithShapes();
        var data = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));

        Assert.Equal(1, data["index"]!.GetValue<int>());
        Assert.Equal(3, data["shapeCount"]!.GetValue<int>());
        Assert.Equal(3, data["shapes"]!.AsArray().Count);
    }

    [Fact]
    public void Get_OutOfRangeShape_IsInvalidPathWithCandidates()
    {
        CreateDeckWithShapes();
        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/shape[99]")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.NotNull(error.Candidates);
        Assert.Equal(3, error.Candidates!.Count);
        Assert.All(error.Candidates!, c => Assert.StartsWith("/slide[1]/shape[@id=", c, StringComparison.Ordinal));
    }

    [Fact]
    public void Get_OutOfRangeSlide_IsInvalidPathWithCandidates()
    {
        CreateDeckWithShapes();
        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[9]")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.Equal(new[] { "/slide[1]" }, error.Candidates!);
    }

    [Fact]
    public void Get_NotesAddressing_IsReservedUnsupportedFeature()
    {
        CreateDeckWithShapes();
        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/notes[1]")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("/slide[", error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Get_RunLevel_IsTypedUnsupportedFeature()
    {
        CreateDeckWithShapes();
        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/shape[1]/p[1]/run[1]")));

        TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
    }
}
