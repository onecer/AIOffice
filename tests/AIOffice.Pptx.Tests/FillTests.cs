using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// 1.8 visual fills: the <c>gradient</c> and <c>image</c> props on shapes and on
/// slide/master/layout backgrounds. Every edit round-trips through the raw OOXML
/// (a:gradFill / a:blipFill) and reopens validator-clean; a render sanity check
/// confirms the gradient colour reaches the SVG.
/// </summary>
public sealed class FillTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() =>
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Cover"))));

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private string WriteImage(string name, byte[] bytes)
    {
        File.WriteAllBytes(_ws.PathOf(name), bytes);
        return name;
    }

    /// <summary>The fill children a shape's spPr may carry (exactly one should survive a fill set).</summary>
    private static bool IsFill(DocumentFormat.OpenXml.OpenXmlElement e) =>
        e is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill;

    private static JsonObject LinearGradient(int angle = 90) => new()
    {
        ["type"] = "linear",
        ["angle"] = angle,
        ["stops"] = new JsonArray(
            new JsonObject { ["color"] = "0EA5E9", ["at"] = 0 },
            new JsonObject { ["color"] = "6366F1", ["at"] = 100 }),
    };

    // ---- gradient on a shape -------------------------------------------------

    [Fact]
    public void SetShapeGradient_WritesGradFill_RoundTrips()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "10cm"), ("h", "6cm"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("gradient", LinearGradient(90)))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var shape = doc.PresentationPart!.SlideParts.Single().Slide!
                .Descendants<P.Shape>().Single(s => s.ShapeProperties?.GetFirstChild<A.GradientFill>() is not null);
            var gradient = shape.ShapeProperties!.GetFirstChild<A.GradientFill>()!;
            var stops = gradient.GradientStopList!.Elements<A.GradientStop>().ToList();
            Assert.Equal(2, stops.Count);
            Assert.Equal(0, stops[0].Position!.Value);
            Assert.Equal(100000, stops[1].Position!.Value);
            Assert.Equal("0EA5E9", stops[0].RgbColorModelHex!.Val!.Value);
            Assert.Equal("6366F1", stops[1].RgbColorModelHex!.Val!.Value);
            Assert.Equal(90 * 60000, gradient.GetFirstChild<A.LinearGradientFill>()!.Angle!.Value);

            // Exactly one fill child survives (no leftover solid fill).
            Assert.Single(shape.ShapeProperties!.ChildElements, IsFill);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RadialGradient_WritesPathFill()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "ellipse"), ("x", "2cm"), ("y", "2cm"), ("w", "8cm"), ("h", "8cm"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        var gradient = new JsonObject
        {
            ["type"] = "radial",
            ["stops"] = new JsonArray(
                new JsonObject { ["color"] = "FFFFFF", ["at"] = 0 },
                new JsonObject { ["color"] = "0F172A", ["at"] = 100 }),
        };
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("gradient", gradient))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var shape = doc.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<P.Shape>().Single(s => s.ShapeProperties?.GetFirstChild<A.GradientFill>() is not null);
        var grad = shape.ShapeProperties!.GetFirstChild<A.GradientFill>()!;
        Assert.NotNull(grad.GetFirstChild<A.PathGradientFill>());
        Assert.Equal(A.PathShadeValues.Circle, grad.GetFirstChild<A.PathGradientFill>()!.Path!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void GradientReplacesPriorSolidFill()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("fill", "FF0000"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("gradient", LinearGradient()))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var shape = doc.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<P.Shape>().Single(s => s.ShapeProperties?.GetFirstChild<A.GradientFill>() is not null);
        Assert.Null(shape.ShapeProperties!.GetFirstChild<A.SolidFill>());
        Assert.Single(shape.ShapeProperties!.ChildElements, IsFill);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddShapeWithGradient_GradientWinsOverFill()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("fill", "FF0000"), ("gradient", LinearGradient()))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var shape = doc.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<P.Shape>().Single(s => s.ShapeProperties?.GetFirstChild<A.GradientFill>() is not null);
        // gradient wins: no solid fill, exactly one fill child.
        Assert.Null(shape.ShapeProperties!.GetFirstChild<A.SolidFill>());
        Assert.Single(shape.ShapeProperties!.ChildElements, IsFill);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ShapeGradient_RendersStartColorIntoSvg()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "10cm"), ("h", "6cm"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("gradient", LinearGradient()))));

        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();
        // The shape outline paints the gradient's start stop (renderer keeps the raw hex case).
        Assert.Contains("fill=\"#0EA5E9\"", svg, StringComparison.Ordinal);
    }

    // ---- gradient on a background --------------------------------------------

    [Fact]
    public void SlideBackgroundGradient_WritesBgGradFill_AndRenders()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("gradient", LinearGradient(45)))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slideData = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!;
            var gradient = slideData.Background!.BackgroundProperties!.GetFirstChild<A.GradientFill>();
            Assert.NotNull(gradient);
            Assert.Equal(2, gradient!.GradientStopList!.Elements<A.GradientStop>().Count());
        }

        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();
        Assert.Contains("fill=\"#0ea5e9\"", svg, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MasterBackgroundGradient_WritesBgGradFill()
    {
        Create();
        Edit(TestEnv.Op("set", "/master[1]", props: TestEnv.Props(("gradient", LinearGradient()))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var slideData = doc.PresentationPart!.SlideMasterParts.Single().SlideMaster!.CommonSlideData!;
        Assert.NotNull(slideData.Background!.BackgroundProperties!.GetFirstChild<A.GradientFill>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void LayoutBackgroundGradient_WritesBgGradFill()
    {
        Create();
        Edit(TestEnv.Op("set", "/master[1]/layout[1]", props: TestEnv.Props(("gradient", LinearGradient()))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var layout = doc.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts.Single();
        var slideData = layout.SlideLayout!.CommonSlideData!;
        Assert.NotNull(slideData.Background!.BackgroundProperties!.GetFirstChild<A.GradientFill>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- image fill on a shape -----------------------------------------------

    [Fact]
    public void SetShapeImageFill_EmbedsBlipFill_StretchByDefault()
    {
        Create();
        WriteImage("banner.png", TestImages.Png(64, 32));
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "10cm"), ("h", "6cm"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("image", "banner.png"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var slidePart = doc.PresentationPart!.SlideParts.Single();
        var imagePart = Assert.Single(slidePart.ImageParts);
        var shape = slidePart.Slide!.Descendants<P.Shape>()
            .Single(s => s.ShapeProperties?.GetFirstChild<A.BlipFill>() is not null);
        var blipFill = shape.ShapeProperties!.GetFirstChild<A.BlipFill>()!;
        Assert.Equal(slidePart.GetIdOfPart(imagePart), blipFill.Blip!.Embed!.Value);
        Assert.NotNull(blipFill.GetFirstChild<A.Stretch>());
        Assert.Single(shape.ShapeProperties!.ChildElements, IsFill);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ShapeImageFill_TileMode_WritesTile()
    {
        Create();
        WriteImage("tile.jpg", TestImages.Jpeg(48, 48));
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "10cm"), ("h", "6cm"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        var image = new JsonObject { ["src"] = "tile.jpg", ["mode"] = "tile" };
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("image", image))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var shape = doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>()
            .Single(s => s.ShapeProperties?.GetFirstChild<A.BlipFill>() is not null);
        var blipFill = shape.ShapeProperties!.GetFirstChild<A.BlipFill>()!;
        Assert.NotNull(blipFill.GetFirstChild<A.Tile>());
        Assert.Null(blipFill.GetFirstChild<A.Stretch>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ShapeImageFill_WithTint_WritesDuotone()
    {
        Create();
        WriteImage("banner.png", TestImages.Png(40, 20));
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "10cm"), ("h", "6cm"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        var image = new JsonObject { ["src"] = "banner.png", ["tint"] = "1E40AF" };
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("image", image))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var blip = doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.BlipFill>().Single().Blip!;
        var duotone = blip.GetFirstChild<A.Duotone>();
        Assert.NotNull(duotone);
        Assert.Contains(duotone!.Elements<A.RgbColorModelHex>(), c => c.Val!.Value == "1E40AF");
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddShapeWithImageFill_EmbedsAtCreation()
    {
        Create();
        WriteImage("banner.png", TestImages.Png(50, 25));
        Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "10cm"), ("h", "6cm"),
            ("image", "banner.png"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var slidePart = doc.PresentationPart!.SlideParts.Single();
        Assert.Single(slidePart.ImageParts);
        Assert.Single(slidePart.Slide!.Descendants<A.BlipFill>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MasterBackgroundImage_EmbedsOnMasterPart()
    {
        Create();
        WriteImage("bg.png", TestImages.Png(80, 45));
        Edit(TestEnv.Op("set", "/master[1]", props: TestEnv.Props(("image", "bg.png"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var masterPart = doc.PresentationPart!.SlideMasterParts.Single();
        Assert.Single(masterPart.ImageParts);
        var slideData = masterPart.SlideMaster!.CommonSlideData!;
        Assert.NotNull(slideData.Background!.BackgroundProperties!.GetFirstChild<A.BlipFill>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- validation / errors -------------------------------------------------

    [Fact]
    public void ImageFill_OutsideSandbox_IsDenied()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("shape", "rect"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("image", "../escape.png")))]);
        Assert.False(envelope.IsOk);
    }

    [Fact]
    public void GradientWithOneStop_IsInvalidArgs()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("shape", "rect"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        var gradient = new JsonObject
        {
            ["type"] = "linear",
            ["stops"] = new JsonArray(new JsonObject { ["color"] = "0EA5E9", ["at"] = 0 }),
        };
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("gradient", gradient)))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void GradientStopOutOfRange_IsInvalidArgs()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("shape", "rect"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        var gradient = new JsonObject
        {
            ["type"] = "linear",
            ["stops"] = new JsonArray(
                new JsonObject { ["color"] = "0EA5E9", ["at"] = 0 },
                new JsonObject { ["color"] = "6366F1", ["at"] = 150 }),
        };
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("gradient", gradient)))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void UnknownGradientType_IsUnsupportedFeature()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("shape", "rect"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        var gradient = new JsonObject
        {
            ["type"] = "conic",
            ["stops"] = new JsonArray(
                new JsonObject { ["color"] = "0EA5E9", ["at"] = 0 },
                new JsonObject { ["color"] = "6366F1", ["at"] = 100 }),
        };
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("gradient", gradient)))]);
        TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
    }

    // ---- read-side: get projects the FULL shape fill object (1.19) ------------

    private JsonObject Get(string path) => TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));

    private string AddRect()
    {
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "10cm"), ("h", "6cm"))));
        return added["results"]![0]!["target"]!.GetValue<string>();
    }

    [Fact]
    public void GetShape_SolidFill_StaysBareHexStringByteIdentical()
    {
        // BYTE-STABLE: a:solidFill projects Fill as the bare RRGGBB hex string FillHex returns today,
        // NOT an object and with no new field — identical to the pre-1.19 get output.
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("fill", "4472C4"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        var fill = Get(path)["fill"];
        var value = Assert.IsType<JsonValue>(fill, exactMatch: false);
        Assert.Equal("4472C4", value.GetValue<string>());
    }

    [Fact]
    public void GetShape_NoFill_ProjectsNull()
    {
        // BYTE-STABLE: a shape with no explicit fill projects Fill=null (omitted) exactly as today.
        Create();
        var path = AddRect();

        Assert.Null(Get(path)["fill"]);
    }

    [Fact]
    public void GetShape_GradientFill_ProjectsTheFullObjectAndRoundTrips()
    {
        Create();
        var path = AddRect();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("gradient", LinearGradient(45)))));

        var fill = Assert.IsType<JsonObject>(Get(path)["fill"]);
        Assert.Equal("linear", fill["type"]!.GetValue<string>());
        // sub-0.1deg drift tolerance (a:lin@ang ÷ 60000 → degrees), mirroring v1.18's background read-back.
        Assert.Equal(45.0, fill["angle"]!.GetValue<double>(), 1);
        var stops = Assert.IsType<JsonArray>(fill["stops"]);
        Assert.Equal(2, stops.Count);
        Assert.Equal("0EA5E9", stops[0]!["color"]!.GetValue<string>());
        Assert.Equal(0.0, stops[0]!["at"]!.GetValue<double>(), 1);
        Assert.Equal("6366F1", stops[1]!["color"]!.GetValue<string>());
        Assert.Equal(100.0, stops[1]!["at"]!.GetValue<double>(), 1);

        // Re-feed the projected object back into set (it is the shape the 'gradient' key carries on
        // write) → byte-identical a:gradFill.
        string gradXmlBefore;
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            gradXmlBefore = doc.PresentationPart!.SlideParts.Single().Slide!
                .Descendants<P.Shape>().Single(s => s.ShapeProperties?.GetFirstChild<A.GradientFill>() is not null)
                .ShapeProperties!.GetFirstChild<A.GradientFill>()!.OuterXml;
        }

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("gradient", fill.DeepClone()))));
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var gradXmlAfter = doc.PresentationPart!.SlideParts.Single().Slide!
                .Descendants<P.Shape>().Single(s => s.ShapeProperties?.GetFirstChild<A.GradientFill>() is not null)
                .ShapeProperties!.GetFirstChild<A.GradientFill>()!.OuterXml;
            Assert.Equal(gradXmlBefore, gradXmlAfter);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void GetShape_RadialGradientFill_OmitsAngleKey()
    {
        Create();
        var path = AddRect();
        var gradient = new JsonObject
        {
            ["type"] = "radial",
            ["stops"] = new JsonArray(
                new JsonObject { ["color"] = "FFFFFF", ["at"] = 0 },
                new JsonObject { ["color"] = "0F172A", ["at"] = 100 }),
        };
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("gradient", gradient))));

        var fill = Assert.IsType<JsonObject>(Get(path)["fill"]);
        Assert.Equal("radial", fill["type"]!.GetValue<string>());
        Assert.False(fill.ContainsKey("angle"), "radial writes a:path, not a:lin — there is NO angle key");
        Assert.Equal(2, Assert.IsType<JsonArray>(fill["stops"]).Count);
    }

    [Fact]
    public void GetShape_ImageFill_ProjectsSrcModeTintAndRoundTrips()
    {
        Create();
        WriteImage("banner.png", TestImages.Png(40, 20));
        var path = AddRect();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("image", new JsonObject { ["src"] = "banner.png", ["mode"] = "tile", ["tint"] = "1E40AF" }))));

        var fill = Assert.IsType<JsonObject>(Get(path)["fill"]);
        // src is the embedded media-part filename (the original caller path is not stored in OOXML).
        Assert.EndsWith(".png", fill["src"]!.GetValue<string>(), StringComparison.Ordinal);
        Assert.Equal("tile", fill["mode"]!.GetValue<string>());
        Assert.Equal("1E40AF", fill["tint"]!.GetValue<string>());

        // Re-feed projected object back into set (the shape the 'image' key carries) → valid a:blipFill.
        var reSrc = fill["src"]!.GetValue<string>();
        File.WriteAllBytes(_ws.PathOf(reSrc), TestImages.Png(8, 8));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("image", fill.DeepClone()))));
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Single(doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<A.BlipFill>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void GetShape_ImageFillWithoutTint_OmitsTheTintKey()
    {
        Create();
        WriteImage("banner.png", TestImages.Png(40, 20));
        var path = AddRect();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("image", "banner.png"))));

        var fill = Assert.IsType<JsonObject>(Get(path)["fill"]);
        Assert.EndsWith(".png", fill["src"]!.GetValue<string>(), StringComparison.Ordinal);
        Assert.Equal("stretch", fill["mode"]!.GetValue<string>());
        Assert.False(fill.ContainsKey("tint"), "tint is OMITTED when no a:duotone is present");
    }

    // ---- coverage: same Fill shape on a group child + master/layout shape ------

    [Fact]
    public void GetGroupChild_GradientFill_ProjectsTheSameObjectShape()
    {
        Create();
        var path = AddRect();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("gradient", LinearGradient(90)))));
        var childId = uint.Parse(
            path.Split("@id=")[1].TrimEnd(']'), System.Globalization.CultureInfo.InvariantCulture);
        var otherId = uint.Parse(
            AddRect().Split("@id=")[1].TrimEnd(']'), System.Globalization.CultureInfo.InvariantCulture);

        var grouped = Edit(TestEnv.Op("add", "/slide[1]", type: "group",
            props: TestEnv.Props(("shapes", new JsonArray("@" + childId, "@" + otherId)))));
        var groupPath = grouped["results"]![0]!["target"]!.GetValue<string>();

        var fill = Assert.IsType<JsonObject>(Get(groupPath + "/shape[@id=" + childId + "]")["fill"]);
        Assert.Equal("linear", fill["type"]!.GetValue<string>());
        Assert.Equal(2, Assert.IsType<JsonArray>(fill["stops"]).Count);
    }

    [Fact]
    public void GetGroupChild_ImageFill_ResolvesSrcOffTheSlidePart()
    {
        Create();
        WriteImage("banner.png", TestImages.Png(40, 20));
        var path = AddRect();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("image", "banner.png"))));
        var childId = uint.Parse(
            path.Split("@id=")[1].TrimEnd(']'), System.Globalization.CultureInfo.InvariantCulture);
        var otherId = uint.Parse(
            AddRect().Split("@id=")[1].TrimEnd(']'), System.Globalization.CultureInfo.InvariantCulture);

        var grouped = Edit(TestEnv.Op("add", "/slide[1]", type: "group",
            props: TestEnv.Props(("shapes", new JsonArray("@" + childId, "@" + otherId)))));
        var groupPath = grouped["results"]![0]!["target"]!.GetValue<string>();

        // The blipFill relId resolves off the slide part (proves GroupDetail passes the right part).
        var fill = Assert.IsType<JsonObject>(Get(groupPath + "/shape[@id=" + childId + "]")["fill"]);
        Assert.EndsWith(".png", fill["src"]!.GetValue<string>(), StringComparison.Ordinal);
        Assert.Equal("stretch", fill["mode"]!.GetValue<string>());
    }

    [Fact]
    public void GetLayoutShape_GradientFill_ProjectsTheSameObjectShape()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/master[1]/layout[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "10cm"), ("h", "6cm"))));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("gradient", LinearGradient(90)))));

        var fill = Assert.IsType<JsonObject>(Get(shapePath)["fill"]);
        Assert.Equal("linear", fill["type"]!.GetValue<string>());
        Assert.Equal(2, Assert.IsType<JsonArray>(fill["stops"]).Count);
    }

    [Fact]
    public void GetLayoutShape_ImageFill_ResolvesSrcOffTheLayoutPart()
    {
        Create();
        WriteImage("banner.png", TestImages.Png(40, 20));
        var added = Edit(TestEnv.Op("add", "/master[1]/layout[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", "2cm"), ("y", "2cm"), ("w", "10cm"), ("h", "6cm"))));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("image", "banner.png"))));

        // The blipFill relId resolves off the LAYOUT part (proves the part-resolution generalization).
        var fill = Assert.IsType<JsonObject>(Get(shapePath)["fill"]);
        Assert.EndsWith(".png", fill["src"]!.GetValue<string>(), StringComparison.Ordinal);
        Assert.Equal("stretch", fill["mode"]!.GetValue<string>());
    }

    // ---- negatives: no area fill / pattFill / empty spPr → null, no null-ref ---

    [Fact]
    public void GetPicture_NoAreaFill_ProjectsNullWithoutNullRef()
    {
        Create();
        WriteImage("banner.png", TestImages.Png(40, 20));
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(
            ("src", "banner.png"), ("x", "2cm"), ("y", "2cm"), ("w", "6cm"), ("h", "3cm"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        // A P.Picture has a blipFill in its spPr image content, not an *area* fill → Fill=null.
        Assert.Null(Get(path)["fill"]);
    }

    [Fact]
    public void GetConnector_NoFill_ProjectsNull()
    {
        Create();
        var a = AddRect();
        var b = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", "14cm"), ("y", "2cm"), ("w", "4cm"), ("h", "3cm"))));
        var bPath = b["results"]![0]!["target"]!.GetValue<string>();
        var aId = uint.Parse(a.Split("@id=")[1].TrimEnd(']'), System.Globalization.CultureInfo.InvariantCulture);
        var bId = uint.Parse(bPath.Split("@id=")[1].TrimEnd(']'), System.Globalization.CultureInfo.InvariantCulture);

        var connected = Edit(TestEnv.Op("add", "/slide[1]", type: "connector", props: TestEnv.Props(
            ("from", JsonValue.Create("@" + aId)), ("to", JsonValue.Create("@" + bId)))));
        var connPath = connected["results"]![0]!["target"]!.GetValue<string>();

        // A P.ConnectionShape carries a line, not an area fill → Fill=null, no null-ref.
        Assert.Null(Get(connPath)["fill"]);
    }

    [Fact]
    public void GetShape_PatternFill_ProjectsNull()
    {
        // pattFill is neither solid/gradient/blip → FillDetail returns null (no new projection).
        Create();
        var path = AddRect();
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true))
        {
            var shape = doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>().Last();
            shape.ShapeProperties!.Append(new A.PatternFill { Preset = A.PresetPatternValues.Cross });
            doc.PresentationPart!.SlideParts.Single().Slide!.Save();
        }

        Assert.Null(Get(path)["fill"]);
    }

    // ---- selector match semantics stay hex-only (FillHex at :456) -------------

    [Fact]
    public void Selector_OverGradientShape_StaysHexOnly_NoMatch()
    {
        // A selector over fill still uses FillHex (non-solid → null): a gradient shape does not match
        // any =RRGGBB predicate — byte-identical match semantics, unchanged by the FillDetail flips.
        Create();
        var path = AddRect();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("gradient", LinearGradient(90)))));

        // Selector filtering stays hex-only: the gradient's start color is NOT a solid fill → 0 matches.
        var data = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape[fill=0EA5E9]"))));
        Assert.Equal(0, data["count"]!.GetValue<int>());
    }
}
