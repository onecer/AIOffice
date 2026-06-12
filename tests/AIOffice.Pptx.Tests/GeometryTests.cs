using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>M3 preset geometries: prstGeom wiring, line connectors and truthful SVG.</summary>
public sealed class GeometryTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
    }

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private string RenderSvg()
    {
        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        return data["slides"]![0]!["svg"]!.GetValue<string>();
    }

    [Theory]
    [InlineData("rect", "rect", "<rect ")]
    [InlineData("roundRect", "roundRect", "rx=")]
    [InlineData("ellipse", "ellipse", "<ellipse ")]
    [InlineData("triangle", "triangle", "<polygon ")]
    [InlineData("diamond", "diamond", "<polygon ")]
    [InlineData("arrow", "rightArrow", "<polygon ")]
    public void AddPresetGeometry_ReopensAndRendersTruthfully(string token, string preset, string svgMarker)
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", token),
            ("x", JsonValue.Create(4)), ("y", JsonValue.Create(4)),
            ("w", JsonValue.Create(8)), ("h", JsonValue.Create(6)),
            ("fill", "3366CC"))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var shape = doc.PresentationPart!.SlideParts.Single().Slide!
                .Descendants<P.Shape>()
                .Single(s => s.ShapeProperties?.GetFirstChild<A.SolidFill>() is not null);
            Assert.Equal(preset, shape.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.Preset!.InnerText);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal(token, detail["geometry"]!.GetValue<string>());

        Assert.Contains(svgMarker, RenderSvg(), StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddLine_BuildsConnectorWithFlipV()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "line"),
            ("x", JsonValue.Create(2)), ("y", JsonValue.Create(3)),
            ("w", JsonValue.Create(10)), ("h", JsonValue.Create(5)),
            ("flip", "v"),
            ("fill", "CC0000"))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var connector = doc.PresentationPart!.SlideParts.Single().Slide!
                .Descendants<P.ConnectionShape>().Single();
            var properties = connector.ShapeProperties!;
            Assert.Equal("line", properties.GetFirstChild<A.PresetGeometry>()!.Preset!.InnerText);
            Assert.True(properties.Transform2D!.VerticalFlip!.Value);
            Assert.Equal(
                "CC0000",
                properties.GetFirstChild<A.Outline>()!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal("connector", detail["kind"]!.GetValue<string>());
        Assert.Equal("line", detail["geometry"]!.GetValue<string>());
        Assert.Equal("v", detail["flip"]!.GetValue<string>());
        Assert.Equal("CC0000", detail["lineColor"]!.GetValue<string>());

        // 2cm=75.6px 3cm=113.4px (96dpi); flipV runs bottom-left (302.4) to top-right (113.4).
        var svg = RenderSvg();
        Assert.Contains(
            "<line x1=\"75.6\" y1=\"302.4\" x2=\"453.5\" y2=\"113.4\" stroke=\"#cc0000\"",
            svg,
            StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddLine_WithText_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "line"), ("text", "nope")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void UnknownGeometry_IsTypedUnsupportedWithCandidates()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("shape", "hexagon")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Equal(["rect", "roundRect", "ellipse", "triangle", "diamond", "arrow", "line"], error.Candidates!);
    }

    [Fact]
    public void InvalidFlip_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "line"), ("flip", "diagonal")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Equal(["h", "v", "hv"], error.Candidates!);
    }

    [Fact]
    public void GeometryShape_StillHoldsTextAndFont()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "ellipse"), ("text", "Inside"), ("fontSize", JsonValue.Create(20)))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal("Inside", detail["text"]!.GetValue<string>());
        Assert.Equal(20, detail["font"]!["sizePt"]!.GetValue<double>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void DefaultTextbox_ReportsRectGeometry_AndStaysTextbox()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "plain"))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var shape = doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Single();
            Assert.True(shape.NonVisualShapeProperties!.NonVisualShapeDrawingProperties!.TextBox!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal("rect", detail["geometry"]!.GetValue<string>());
        Assert.Null(detail["flip"]);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
