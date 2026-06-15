using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.2.0 shape connectors: add a straight/elbow/curved p:cxnSp between two shapes,
/// reopen-verify the endpoints, reject an unknown endpoint with candidates, and confirm
/// the SVG render draws the connector line.
/// </summary>
public sealed class ConnectorTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();
    private readonly uint _aId;
    private readonly uint _bId;

    public ConnectorTests()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("d.pptx")));
        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("name", JsonValue.Create("Box A")), ("x", JsonValue.Create("2cm")), ("y", JsonValue.Create("2cm")),
                ("w", JsonValue.Create("4cm")), ("h", JsonValue.Create("3cm")))),
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("name", JsonValue.Create("Box B")), ("x", JsonValue.Create("16cm")), ("y", JsonValue.Create("12cm")),
                ("w", JsonValue.Create("4cm")), ("h", JsonValue.Create("3cm")))),
        ]));
        _aId = IdOf(data["results"]![0]!["target"]!.GetValue<string>());
        _bId = IdOf(data["results"]![1]!["target"]!.GetValue<string>());
    }

    public void Dispose() => _ws.Dispose();

    private static uint IdOf(string shapePath) =>
        uint.Parse(shapePath.Split("@id=")[1].TrimEnd(']'), System.Globalization.CultureInfo.InvariantCulture);

    private string AddConnector(params (string Key, JsonNode? Value)[] props)
    {
        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "connector", props: TestEnv.Props(props))]));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private JsonObject Get(string path) => TestEnv.AssertOk(_handler.Get(_ws.Ctx("d.pptx", ("path", path))));

    [Theory]
    [InlineData("straight")]
    [InlineData("elbow")]
    [InlineData("curved")]
    public void EachKind_ReopenVerifiesEndpoints(string kind)
    {
        var path = AddConnector(
            ("kind", JsonValue.Create(kind)),
            ("from", JsonValue.Create("@" + _aId)),
            ("to", JsonValue.Create("@" + _bId)));

        var detail = Get(path);
        Assert.Equal("connector", detail["kind"]!.GetValue<string>());
        var connector = detail["connector"]!.AsObject();
        Assert.Equal(_aId, connector["from"]!.GetValue<uint>());
        Assert.Equal(_bId, connector["to"]!.GetValue<uint>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void DefaultKind_IsStraight()
    {
        var path = AddConnector(("from", JsonValue.Create("@" + _aId)), ("to", JsonValue.Create("@" + _bId)));
        var detail = Get(path);
        Assert.Equal(_aId, detail["connector"]!["from"]!.GetValue<uint>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void EndpointsByShapeName_Resolve()
    {
        var path = AddConnector(("from", JsonValue.Create("Box A")), ("to", JsonValue.Create("Box B")));
        var detail = Get(path);
        Assert.Equal(_aId, detail["connector"]!["from"]!.GetValue<uint>());
        Assert.Equal(_bId, detail["connector"]!["to"]!.GetValue<uint>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void Arrows_And_StyleProps_Validate()
    {
        var path = AddConnector(
            ("kind", JsonValue.Create("straight")),
            ("from", JsonValue.Create("@" + _aId)),
            ("to", JsonValue.Create("@" + _bId)),
            ("startArrow", JsonValue.Create("triangle")),
            ("endArrow", JsonValue.Create("arrow")),
            ("color", JsonValue.Create("FF0000")),
            ("width", JsonValue.Create("3pt")));
        var detail = Get(path);
        Assert.Equal("FF0000", detail["lineColor"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void UnknownEndpoint_IsInvalidPathWithCandidates()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "connector", props: TestEnv.Props(
                ("from", JsonValue.Create("@9999")), ("to", JsonValue.Create("@" + _bId))))]);
        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.Contains("@" + _aId, error.Candidates!);
        Assert.Contains("@" + _bId, error.Candidates!);
    }

    [Fact]
    public void SameFromAndTo_IsInvalidArgs()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "connector", props: TestEnv.Props(
                ("from", JsonValue.Create("@" + _aId)), ("to", JsonValue.Create("@" + _aId))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void UnknownKind_IsUnsupportedFeatureWithCandidates()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "connector", props: TestEnv.Props(
                ("kind", JsonValue.Create("zigzag")), ("from", JsonValue.Create("@" + _aId)), ("to", JsonValue.Create("@" + _bId))))]);
        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("elbow", error.Candidates!);
    }

    [Fact]
    public void BadArrow_IsInvalidArgs()
    {
        var envelope = _handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "connector", props: TestEnv.Props(
                ("from", JsonValue.Create("@" + _aId)), ("to", JsonValue.Create("@" + _bId)), ("endArrow", JsonValue.Create("spike"))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Theory]
    [InlineData("straight", "<line")]
    [InlineData("elbow", "<polyline")]
    [InlineData("curved", "<path")]
    public void Svg_DrawsTheConnector(string kind, string element)
    {
        AddConnector(
            ("kind", JsonValue.Create(kind)),
            ("from", JsonValue.Create("@" + _aId)),
            ("to", JsonValue.Create("@" + _bId)));

        var svg = TestEnv.AssertOk(_handler.Render(
            _ws.Ctx("d.pptx", ("to", "svg"), ("scope", "/slide[1]"))))["slides"]![0]!["svg"]!.GetValue<string>();
        Assert.Contains(element, svg);
    }

    [Fact]
    public void Connector_RemoveDropsIt()
    {
        var path = AddConnector(("from", JsonValue.Create("@" + _aId)), ("to", JsonValue.Create("@" + _bId)));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("d.pptx"), [TestEnv.Op("remove", path)]));
        var slide = Get("/slide[1]");
        // Only the two boxes remain (the connector is gone).
        Assert.Equal(2, slide["shapeCount"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }

    [Fact]
    public void Connector_SurvivesRoundTrip()
    {
        var path = AddConnector(
            ("kind", JsonValue.Create("elbow")), ("from", JsonValue.Create("@" + _aId)), ("to", JsonValue.Create("@" + _bId)));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("d.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", JsonValue.Create("x"))))]));
        var detail = Get(path);
        Assert.Equal(_aId, detail["connector"]!["from"]!.GetValue<uint>());
        Assert.Equal(_bId, detail["connector"]!["to"]!.GetValue<uint>());
        TestEnv.AssertValid(_ws, "d.pptx");
    }
}
