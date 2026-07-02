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

    // ----- adjust handles (v1.22): a:prstGeom/a:avLst/a:gd guides ------------------

    /// <summary>The single shape's preset geometry, read straight from the saved OOXML.</summary>
    private static A.PresetGeometry PresetGeometryOfSingleShape(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<P.Shape>().Single()
            .ShapeProperties!.GetFirstChild<A.PresetGeometry>()!;

    [Fact]
    public void AddWithoutAdjust_KeepsEmptyAvLst_AndGetEmitsNoAdjustKey()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "roundRect"), ("fill", "3366CC"))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var geometry = PresetGeometryOfSingleShape(doc);
            Assert.NotNull(geometry.AdjustValueList);
            Assert.Empty(geometry.AdjustValueList!.Elements<A.ShapeGuide>()); // byte-stable v1.21 shape
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.False(detail.ContainsKey("adjust")); // empty avLst projects no adjust key
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("roundRect", 16667)]
    [InlineData("triangle", 25000)]
    public void AddSingleGuidePreset_WritesAdjGuide_AndRoundTrips(string token, int adjust)
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", token), ("adjust", JsonValue.Create(adjust)))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var guide = Assert.Single(PresetGeometryOfSingleShape(doc).AdjustValueList!.Elements<A.ShapeGuide>());
            Assert.Equal("adj", guide.Name!.Value);
            Assert.Equal($"val {adjust}", guide.Formula!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal(adjust, detail["adjust"]!.GetValue<long>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddArrowAdjust_WritesBothGuides_AndRoundTrips()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "arrow"),
            ("adjust", TestEnv.Props(("adj1", JsonValue.Create(60000)), ("adj2", JsonValue.Create(40000)))))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var guides = PresetGeometryOfSingleShape(doc).AdjustValueList!.Elements<A.ShapeGuide>().ToList();
            Assert.Equal(["adj1", "adj2"], guides.Select(g => g.Name!.Value!));
            Assert.Equal(["val 60000", "val 40000"], guides.Select(g => g.Formula!.Value!));
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal(60000, detail["adjust"]!["adj1"]!.GetValue<long>());
        Assert.Equal(40000, detail["adjust"]!["adj2"]!.GetValue<long>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddArrowAdjust_SingleGuide_WritesExactlyThatGuide()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "arrow"),
            ("adjust", TestEnv.Props(("adj2", JsonValue.Create(25000)))))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var guide = Assert.Single(PresetGeometryOfSingleShape(doc).AdjustValueList!.Elements<A.ShapeGuide>());
            Assert.Equal("adj2", guide.Name!.Value);
            Assert.Equal("val 25000", guide.Formula!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Null(detail["adjust"]!["adj1"]);
        Assert.Equal(25000, detail["adjust"]!["adj2"]!.GetValue<long>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetAdjustTwice_ReplacesNotAccumulates()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("shape", "roundRect"))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", canonical, props: TestEnv.Props(("adjust", JsonValue.Create(30000)))));
        Edit(TestEnv.Op("set", canonical, props: TestEnv.Props(("adjust", JsonValue.Create(60000)))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var guide = Assert.Single(PresetGeometryOfSingleShape(doc).AdjustValueList!.Elements<A.ShapeGuide>());
            Assert.Equal("adj", guide.Name!.Value);
            Assert.Equal("val 60000", guide.Formula!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal(60000, detail["adjust"]!.GetValue<long>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetArrowAdjust_SingleGuide_ReplacesBothExistingGuides()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "arrow"),
            ("adjust", TestEnv.Props(("adj1", JsonValue.Create(60000)), ("adj2", JsonValue.Create(40000)))))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", canonical, props: TestEnv.Props(
            ("adjust", TestEnv.Props(("adj1", JsonValue.Create(50000)))))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var guide = Assert.Single(PresetGeometryOfSingleShape(doc).AdjustValueList!.Elements<A.ShapeGuide>());
            Assert.Equal("adj1", guide.Name!.Value);
            Assert.Equal("val 50000", guide.Formula!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal(50000, detail["adjust"]!["adj1"]!.GetValue<long>());
        Assert.Null(detail["adjust"]!["adj2"]);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ForeignDeckGuides_ProjectOnlyPinnedNames()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("shape", "roundRect"))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        // Simulate a PowerPoint-authored deck: hand-write non-default guides,
        // including a foreign guide name aioffice must ignore.
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true))
        {
            var avLst = PresetGeometryOfSingleShape(doc).AdjustValueList!;
            avLst.Append(new A.ShapeGuide { Name = "adj", Formula = "val 35000" });
            avLst.Append(new A.ShapeGuide { Name = "someFuture", Formula = "val 123" });
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal(35000, detail["adjust"]!.GetValue<long>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ForeignArrowGuides_ProjectOnlySuppliedPinnedNames()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("shape", "arrow"))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true))
        {
            var avLst = PresetGeometryOfSingleShape(doc).AdjustValueList!;
            avLst.Append(new A.ShapeGuide { Name = "adj2", Formula = "val 25000" });
            avLst.Append(new A.ShapeGuide { Name = "someFuture", Formula = "val 9" });
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Null(detail["adjust"]!["adj1"]);
        Assert.Equal(25000, detail["adjust"]!["adj2"]!.GetValue<long>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("rect")]
    [InlineData("ellipse")]
    [InlineData("diamond")]
    public void AdjustOnNonAdjustablePreset_IsInvalidArgsWithCandidates(string token)
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", token), ("adjust", JsonValue.Create(16667))))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Equal(["roundRect", "arrow", "triangle"], error.Candidates!);
    }

    [Fact]
    public void AdjustOnLine_AddAndSet_AreInvalidArgs()
    {
        Create();
        var addEnvelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "line"), ("adjust", JsonValue.Create(16667))))]);
        var addError = TestEnv.AssertFail(addEnvelope, ErrorCodes.InvalidArgs);
        Assert.Equal(["roundRect", "arrow", "triangle"], addError.Candidates!);

        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("shape", "line"))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();
        var setEnvelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", canonical, props: TestEnv.Props(("adjust", JsonValue.Create(16667))))]);
        var setError = TestEnv.AssertFail(setEnvelope, ErrorCodes.InvalidArgs);
        Assert.Equal(["roundRect", "arrow", "triangle"], setError.Candidates!);
    }

    [Theory]
    [InlineData(-5)]
    [InlineData(100001)]
    public void AdjustOutOfRange_IsInvalidArgs(int adjust)
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "roundRect"), ("adjust", JsonValue.Create(adjust))))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("100000", error.Suggestion!, StringComparison.Ordinal);
    }

    [Fact]
    public void AdjustValueShapeMismatch_IsInvalidArgs()
    {
        Create();

        // An {adj1, adj2} object on a single-guide preset.
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "roundRect"), ("adjust", TestEnv.Props(("adj1", JsonValue.Create(16667))))))]),
            ErrorCodes.InvalidArgs);

        // A bare number on arrow (which needs the object form).
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "arrow"), ("adjust", JsonValue.Create(60000))))]),
            ErrorCodes.InvalidArgs);

        // An empty object on arrow (at least one of adj1/adj2 required).
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "arrow"), ("adjust", TestEnv.Props())))]),
            ErrorCodes.InvalidArgs);

        // An unknown guide name on arrow.
        var unknown = TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "arrow"), ("adjust", TestEnv.Props(("adj3", JsonValue.Create(1))))))]),
            ErrorCodes.InvalidArgs);
        Assert.Equal(["adj1", "adj2"], unknown.Candidates!);
    }

    [Fact]
    public void AdjustedShapes_RenderSmokeAndValidate()
    {
        Create();
        Edit(
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "roundRect"), ("adjust", JsonValue.Create(50000)), ("fill", "3366CC"))),
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "arrow"), ("adjust", TestEnv.Props(("adj1", JsonValue.Create(60000)))))),
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("shape", "triangle"), ("adjust", JsonValue.Create(75000)))));

        // Honest note: the SVG renderer approximates every roundRect with the
        // default corner radius (and arrow/triangle with default handles) — it
        // does not read a:avLst, so an adjusted deck renders like a default one.
        // The OOXML asserts above plus the validator carry the correctness proof.
        var svg = RenderSvg();
        Assert.Contains("rx=", svg, StringComparison.Ordinal);
        Assert.Contains("<polygon ", svg, StringComparison.Ordinal);
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
