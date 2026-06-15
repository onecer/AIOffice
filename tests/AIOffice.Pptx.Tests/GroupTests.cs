using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.2.0 shape grouping (additive add-types group / ungroup): wrap shapes in a p:grpSp,
/// list its children, ungroup to restore absolute coordinates, and reorder the group in
/// z-order. Every mutating step reopen-verifies validator-clean.
/// </summary>
public sealed class GroupTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();
    private readonly string _aPath;
    private readonly string _bPath;
    private readonly string _cPath;
    private readonly uint _aId;
    private readonly uint _bId;
    private readonly uint _cId;

    public GroupTests()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("d.pptx")));
        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("d.pptx"), [
            Box("A", "2cm", "2cm"),
            Box("B", "10cm", "8cm"),
            Box("C", "18cm", "14cm"),
        ]));
        _aPath = data["results"]![0]!["target"]!.GetValue<string>();
        _bPath = data["results"]![1]!["target"]!.GetValue<string>();
        _cPath = data["results"]![2]!["target"]!.GetValue<string>();
        _aId = IdOf(_aPath);
        _bId = IdOf(_bPath);
        _cId = IdOf(_cPath);
    }

    public void Dispose() => _ws.Dispose();

    private static EditOp Box(string text, string x, string y) =>
        TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("text", JsonValue.Create(text)), ("name", JsonValue.Create("Box " + text)),
            ("x", JsonValue.Create(x)), ("y", JsonValue.Create(y)),
            ("w", JsonValue.Create("4cm")), ("h", JsonValue.Create("3cm"))));

    private static uint IdOf(string shapePath) =>
        uint.Parse(shapePath.Split("@id=")[1].TrimEnd(']'), System.Globalization.CultureInfo.InvariantCulture);

    private string Group(params string[] refs)
    {
        var array = new JsonArray();
        foreach (var r in refs)
        {
            array.Add(r);
        }

        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "group", props: TestEnv.Props(("shapes", array)))]));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private JsonObject Get(string path) => TestEnv.AssertOk(_handler.Get(_ws.Ctx("d.pptx", ("path", path))));

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("d.pptx"), ops));

    [Fact]
    public void Group_WrapsTheNamedShapes()
    {
        var groupPath = Group("@" + _aId, "@" + _bId);
        Assert.StartsWith("/slide[1]/group[@id=", groupPath, StringComparison.Ordinal);

        var detail = Get(groupPath);
        Assert.Equal("group", detail["kind"]!.GetValue<string>());
        Assert.Equal(2, detail["childCount"]!.GetValue<int>());

        var childIds = detail["children"]!.AsArray().Select(c => c!["id"]!.GetValue<uint>()).ToList();
        Assert.Contains(_aId, childIds);
        Assert.Contains(_bId, childIds);

        // The slide now shows the group plus the ungrouped C (3 -> 2 top-level shapes).
        var slide = Get("/slide[1]");
        Assert.Equal(2, slide["shapeCount"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void GroupBox_IsTheUnionOfMembers()
    {
        var groupPath = Group("@" + _aId, "@" + _bId);
        var detail = Get(groupPath);
        // A at (2,2) 4x3 and B at (10,8) 4x3 -> union top-left (2,2), bottom-right (14,11).
        Assert.Equal(2, detail["x"]!.GetValue<double>());
        Assert.Equal(2, detail["y"]!.GetValue<double>());
        Assert.Equal(12, detail["w"]!.GetValue<double>());
        Assert.Equal(9, detail["h"]!.GetValue<double>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void GroupChild_IsAddressableAndKeepsAbsoluteCoords()
    {
        var groupPath = Group("@" + _aId, "@" + _bId);
        var childA = Get(groupPath + "/shape[@id=" + _aId + "]");
        Assert.Equal("A", childA["text"]!.GetValue<string>());
        Assert.Equal(2, childA["x"]!.GetValue<double>());
        Assert.Equal(2, childA["y"]!.GetValue<double>());
    }

    [Fact]
    public void Ungroup_RestoresChildrenWithAbsoluteCoords()
    {
        var groupPath = Group("@" + _aId, "@" + _bId);
        Edit(TestEnv.Op("add", groupPath, type: "ungroup"));

        // Both boxes are back as top-level shapes at their original positions.
        var slide = Get("/slide[1]");
        Assert.Equal(3, slide["shapeCount"]!.GetValue<int>());

        var a = Get(_aPath);
        Assert.Equal(2, a["x"]!.GetValue<double>());
        Assert.Equal(2, a["y"]!.GetValue<double>());
        var b = Get(_bPath);
        Assert.Equal(10, b["x"]!.GetValue<double>());
        Assert.Equal(8, b["y"]!.GetValue<double>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void Group_ThenUngroup_RoundTripsTheDeck()
    {
        var groupPath = Group("@" + _aId, "@" + _bId, "@" + _cId);
        Assert.Equal(3, Get(groupPath)["childCount"]!.GetValue<int>());

        Edit(TestEnv.Op("add", groupPath, type: "ungroup"));
        Assert.Equal(3, Get("/slide[1]")["shapeCount"]!.GetValue<int>());
        // Text survives the wrap+dissolve cycle.
        Assert.Equal("C", Get(_cPath)["text"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void Group_ByName()
    {
        var groupPath = Group("Box A", "Box C");
        var detail = Get(groupPath);
        var childIds = detail["children"]!.AsArray().Select(c => c!["id"]!.GetValue<uint>()).ToList();
        Assert.Contains(_aId, childIds);
        Assert.Contains(_cId, childIds);
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void GroupMoveZOrder_Reorders()
    {
        var groupPath = Group("@" + _aId, "@" + _bId); // group + C are the two top-level shapes
        // The group sits first (paint order 1); move it to front.
        Edit(TestEnv.Op("move", groupPath, position: "front"));

        var slide = Get("/slide[1]");
        var shapes = slide["shapes"]!.AsArray();
        // The last-painted (topmost) top-level shape is now the group.
        Assert.Equal("group", shapes[^1]!["kind"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void SetGroupName_Works()
    {
        var groupPath = Group("@" + _aId, "@" + _bId);
        Edit(TestEnv.Op("set", groupPath, props: TestEnv.Props(("name", JsonValue.Create("My Cluster")))));
        Assert.Equal("My Cluster", Get(groupPath)["name"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void SetGroupChildFill_Works()
    {
        var groupPath = Group("@" + _aId, "@" + _bId);
        Edit(TestEnv.Op("set", groupPath + "/shape[@id=" + _aId + "]", props: TestEnv.Props(("fill", JsonValue.Create("00AA00")))));
        Assert.Equal("00AA00", Get(groupPath + "/shape[@id=" + _aId + "]")["fill"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void RemoveGroupChild_DropsJustThatChild()
    {
        var groupPath = Group("@" + _aId, "@" + _bId);
        Edit(TestEnv.Op("remove", groupPath + "/shape[@id=" + _aId + "]"));
        Assert.Equal(1, Get(groupPath)["childCount"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void RemoveGroup_DeletesTheWholeGroup()
    {
        var groupPath = Group("@" + _aId, "@" + _bId);
        Edit(TestEnv.Op("remove", groupPath));
        // Only C is left (the group and its two children are gone).
        Assert.Equal(1, Get("/slide[1]")["shapeCount"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void Svg_RendersGroupChildren()
    {
        var groupPath = Group("@" + _aId, "@" + _bId);
        var svg = TestEnv.AssertOk(_handler.Render(
            _ws.Ctx("d.pptx", ("to", "svg"), ("scope", "/slide[1]"))))["slides"]![0]!["svg"]!.GetValue<string>();
        // The child's data-aio-path is the group-child path.
        Assert.Contains("/slide[1]/group[@id=" + IdOf(groupPath) + "]/shape[@id=" + _aId + "]", svg);
    }

    [Fact]
    public void GroupWithOneShape_IsInvalidArgs()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "group", props: TestEnv.Props(("shapes", new JsonArray("@" + _aId))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void GroupUnknownShape_IsInvalidPathWithCandidates()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "group", props: TestEnv.Props(("shapes", new JsonArray("@" + _aId, "@9999"))))]);
        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.Contains("@" + _aId, error.Candidates!);
    }

    [Fact]
    public void UngroupOnNonGroup_IsInvalidArgs()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", _aPath, type: "ungroup")]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void UngroupUnknownGroup_IsInvalidPath()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]/group[@id=9999]", type: "ungroup")]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
    }
}
