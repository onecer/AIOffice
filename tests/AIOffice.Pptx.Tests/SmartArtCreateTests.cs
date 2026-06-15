using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.2.0 SmartArt CREATION: add a list/process/hierarchy/orgChart/cycle diagram, then
/// reopen-verify the data-model node texts, the layout reference and validator cleanliness.
/// SmartArt remains read-only for edits (the M7 contract), so editing an existing diagram
/// stays unsupported_feature.
/// </summary>
public sealed class SmartArtCreateTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public SmartArtCreateTests() => TestEnv.AssertOk(_handler.Create(_ws.Ctx("d.pptx")));

    public void Dispose() => _ws.Dispose();

    private static JsonArray Nodes(params (string Text, int Level)[] nodes)
    {
        var array = new JsonArray();
        foreach (var (text, level) in nodes)
        {
            array.Add(new JsonObject { ["text"] = text, ["level"] = level });
        }

        return array;
    }

    private string AddSmartArt(string layout, JsonArray nodes, params (string Key, JsonNode? Value)[] extra)
    {
        var props = TestEnv.Props(("layout", JsonValue.Create(layout)), ("nodes", nodes));
        foreach (var (key, value) in extra)
        {
            props[key] = value;
        }

        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("d.pptx"), [TestEnv.Op("add", "/slide[1]", type: "smartart", props: props)]));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private JsonObject Get(string path) => TestEnv.AssertOk(_handler.Get(_ws.Ctx("d.pptx", ("path", path))));

    [Theory]
    [InlineData("list")]
    [InlineData("process")]
    [InlineData("hierarchy")]
    [InlineData("orgChart")]
    [InlineData("cycle")]
    public void EachLayout_ReopenVerifiesDataModelAndLayoutRef(string layout)
    {
        var path = AddSmartArt(layout, Nodes(("Alpha", 0), ("Beta", 0), ("Gamma", 0)));
        Assert.Equal("/slide[1]/smartart[1]", path);

        var detail = Get(path);
        Assert.Equal("smartart", detail["kind"]!.GetValue<string>());
        Assert.True(detail["readOnly"]!.GetValue<bool>());
        Assert.Equal(3, detail["nodeCount"]!.GetValue<int>());

        // The layout reference survives the round-trip (a non-empty built-in layout id).
        Assert.False(string.IsNullOrEmpty(detail["layout"]!.GetValue<string>()));

        var texts = detail["texts"]!.AsArray().Select(t => t!["text"]!.GetValue<string>()).ToList();
        Assert.Equal(["Alpha", "Beta", "Gamma"], texts);

        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void NestedNodes_BuildAParentChildTree()
    {
        var path = AddSmartArt("hierarchy", Nodes(("CEO", 0), ("VP Eng", 1), ("Lead", 2), ("VP Sales", 1)));
        var detail = Get(path);
        Assert.Equal(4, detail["nodeCount"]!.GetValue<int>());

        var roots = detail["texts"]!.AsArray();
        Assert.Single(roots);
        Assert.Equal("CEO", roots[0]!["text"]!.GetValue<string>());

        var ceoChildren = roots[0]!["children"]!.AsArray();
        Assert.Equal(["VP Eng", "VP Sales"], ceoChildren.Select(c => c!["text"]!.GetValue<string>()).ToList());

        // VP Eng owns Lead (level 2 attaches to the nearest level-1 node before it).
        var vpEngChildren = ceoChildren[0]!["children"]!.AsArray();
        Assert.Equal(["Lead"], vpEngChildren.Select(c => c!["text"]!.GetValue<string>()).ToList());

        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void Geometry_HonorsProvidedXYWH()
    {
        var path = AddSmartArt(
            "cycle",
            Nodes(("One", 0)),
            ("x", JsonValue.Create("3cm")),
            ("y", JsonValue.Create("4cm")),
            ("w", JsonValue.Create("20cm")),
            ("h", JsonValue.Create("10cm")));

        var detail = Get(path);
        Assert.Equal(3, detail["x"]!.GetValue<double>());
        Assert.Equal(4, detail["y"]!.GetValue<double>());
        Assert.Equal(20, detail["w"]!.GetValue<double>());
        Assert.Equal(10, detail["h"]!.GetValue<double>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void TwoDiagrams_GetDistinctOneBasedIndices()
    {
        var first = AddSmartArt("list", Nodes(("A", 0)));
        var second = AddSmartArt("process", Nodes(("B", 0)));
        Assert.Equal("/slide[1]/smartart[1]", first);
        Assert.Equal("/slide[1]/smartart[2]", second);

        Assert.Equal("A", Get(first)["texts"]!.AsArray()[0]!["text"]!.GetValue<string>());
        Assert.Equal("B", Get(second)["texts"]!.AsArray()[0]!["text"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void SlideDetail_ListsTheSmartArt()
    {
        AddSmartArt("list", Nodes(("X", 0)));
        var slide = Get("/slide[1]");
        var smartArt = slide["smartArt"]!.AsArray().Select(s => s!.GetValue<string>()).ToList();
        Assert.Contains("/slide[1]/smartart[1]", smartArt);
    }

    [Fact]
    public void UnknownLayout_IsUnsupportedFeatureWithCandidates()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "smartart", props: TestEnv.Props(
                ("layout", JsonValue.Create("venn")), ("nodes", Nodes(("A", 0)))))]);
        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("list", error.Candidates!);
        Assert.Contains("cycle", error.Candidates!);
    }

    [Fact]
    public void MissingNodes_IsInvalidArgs()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "smartart", props: TestEnv.Props(("layout", JsonValue.Create("list"))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void FirstNodeNotLevelZero_IsInvalidArgs()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "smartart", props: TestEnv.Props(
                ("layout", JsonValue.Create("list")), ("nodes", Nodes(("Orphan", 1)))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void UnknownProp_IsInvalidArgs()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "smartart", props: TestEnv.Props(
                ("layout", JsonValue.Create("list")), ("nodes", Nodes(("A", 0))), ("bogus", JsonValue.Create(1))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void CreatedSmartArt_RemainsReadOnlyForEdits()
    {
        var path = AddSmartArt("process", Nodes(("A", 0)));
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("set", path, props: TestEnv.Props(("layout", JsonValue.Create("cycle"))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
    }

    [Fact]
    public void RoundTrip_OpenAndSaveLeavesSmartArtReadable()
    {
        var path = AddSmartArt("orgChart", Nodes(("Root", 0), ("Child", 1)));

        // A no-op edit reopens and resaves the deck; the diagram must survive intact.
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", JsonValue.Create("note"))))]));

        var detail = Get(path);
        Assert.Equal(2, detail["nodeCount"]!.GetValue<int>());
        Assert.Equal("Root", detail["texts"]!.AsArray()[0]!["text"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }
}
