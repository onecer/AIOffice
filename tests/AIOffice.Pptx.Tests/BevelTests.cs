using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>v1.25 shape bevel / 3-D: a:sp3d/a:bevelT, the first 3-D property, outside a:effectLst.</summary>
public sealed class BevelTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() => TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    /// <summary>Adds a textbox and returns its shape path.</summary>
    private string AddShape()
    {
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "textbox", props: TestEnv.Props(
            ("text", "Bevel"), ("x", "2cm"), ("y", "2cm"), ("w", "8cm"), ("h", "3cm"))));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private P.Shape SingleShape()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        return doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>()
            .First(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("TextBox", StringComparison.Ordinal) == true);
    }

    // ---- (1) byte-stable legacy: a deck edited without bevel gains no a:sp3d ----

    [Fact]
    public void ShapeWithoutBevel_HasNoShape3D_AndNoBevelKey_ByteStable()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "404040"))));

        // No a:sp3d appears; the existing effect switch arm is untouched.
        Assert.Null(SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>());

        var effects = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!;
        Assert.Equal("404040", effects["shadow"]!.GetValue<string>());
        Assert.False(effects.AsObject().ContainsKey("bevel"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (2) round-trip string ----

    [Fact]
    public void SetBevelPreset_WritesBevelTop_AndReadsBackAsString()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", "relaxedInset"))));

        var sp3d = SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>()!;
        var bevelTop = sp3d.GetFirstChild<A.BevelTop>()!;
        Assert.Equal(A.BevelPresetValues.RelaxedInset, bevelTop.Preset!.Value);
        Assert.Equal(76_200L, bevelTop.Width!.Value);  // 6pt default
        Assert.Equal(76_200L, bevelTop.Height!.Value);
        Assert.Null(sp3d.ExtrusionHeight);
        Assert.Null(sp3d.GetFirstChild<A.ExtrusionColor>());

        var effects = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!;
        Assert.Equal("relaxedInset", effects["bevel"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (3) round-trip object ----

    [Fact]
    public void SetBevelObject_WritesSizeDepthColor_AndRoundTripsFieldForField()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject
            {
                ["preset"] = "circle",
                ["width"] = "8pt",
                ["height"] = "8pt",
                ["depth"] = "4pt",
                ["depthColor"] = "C00000",
            },
        }));

        var sp3d = SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>()!;
        var bevelTop = sp3d.GetFirstChild<A.BevelTop>()!;
        Assert.Equal(A.BevelPresetValues.Circle, bevelTop.Preset!.Value);
        Assert.Equal(101_600L, bevelTop.Width!.Value);   // 8pt
        Assert.Equal(101_600L, bevelTop.Height!.Value);  // 8pt
        Assert.Equal(50_800L, sp3d.ExtrusionHeight!.Value); // 4pt
        Assert.Equal("C00000", sp3d.GetFirstChild<A.ExtrusionColor>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);

        var bevel = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!["bevel"]!;
        var obj = Assert.IsType<JsonObject>(bevel);
        Assert.Equal("circle", obj["preset"]!.GetValue<string>());
        Assert.Equal("8pt", obj["width"]!.GetValue<string>());
        Assert.Equal("8pt", obj["height"]!.GetValue<string>());
        Assert.Equal("4pt", obj["depth"]!.GetValue<string>());
        Assert.Equal("C00000", obj["depthColor"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (4) null-omit ----

    [Fact]
    public void ShapeWithoutShape3D_ProjectsNoBevelKey()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("glow", "00FF00"))));

        var effects = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!;
        Assert.Equal("00FF00", effects["glow"]!.GetValue<string>());
        Assert.False(effects.AsObject().ContainsKey("bevel"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (5) child order ----

    [Fact]
    public void Bevel_SitsAfterLineAndEffectList_BevelTopBeforeExtrusionColor()
    {
        Create();
        var path = AddShape();
        // Set an outline (a:ln) and effects (a:effectLst) first, then the bevel object (a:sp3d).
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("outline", "0000FF"), ("glow", "00FF00"), ("shadow", "303030"))));
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject { ["preset"] = "circle", ["depthColor"] = "C00000" },
        }));

        var properties = SingleShape().ShapeProperties!;
        var order = properties.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(order.IndexOf("ln") < order.IndexOf("sp3d"));
        Assert.True(order.IndexOf("effectLst") < order.IndexOf("sp3d"));

        var sp3d = properties.GetFirstChild<A.Shape3DType>()!;
        var sp3dOrder = sp3d.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(sp3dOrder.IndexOf("bevelT") < sp3dOrder.IndexOf("extrusionClr"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Bevel_SetBeforeEffects_KeepsEffectListBeforeShape3D()
    {
        // Order-independence: bevel first, then an effect; a:effectLst must still precede a:sp3d.
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", "circle"))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "303030"))));

        var order = SingleShape().ShapeProperties!.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(order.IndexOf("effectLst") < order.IndexOf("sp3d"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("relaxedInset")]
    [InlineData("artDeco")]
    public void Bevel_PresetOnly_IsValid(string preset)
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", preset))));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Bevel_DepthOnlyObject_IsValid()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject { ["preset"] = "slope", ["depth"] = "3pt" },
        }));
        var sp3d = SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>()!;
        Assert.Equal(38_100L, sp3d.ExtrusionHeight!.Value); // 3pt

        var bevel = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!["bevel"]!;
        var obj = Assert.IsType<JsonObject>(bevel);
        Assert.Equal("slope", obj["preset"]!.GetValue<string>());
        Assert.Equal("3pt", obj["depth"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (6) negatives ----

    [Fact]
    public void Bevel_UnknownPreset_IsInvalidArgsWithTwelveCandidates()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("bevel", "banana")))]);

        var error = TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
        Assert.NotNull(error.Candidates);
        Assert.Equal(12, error.Candidates!.Count);
        Assert.Contains("relaxedInset", error.Candidates!);
        Assert.Contains("artDeco", error.Candidates!);
    }

    [Fact]
    public void Bevel_ObjectUnknownKey_IsInvalidArgsWithObjectCandidates()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: new JsonObject
            {
                ["bevel"] = new JsonObject { ["preset"] = "circle", ["foo"] = "bar" },
            })]);

        var error = TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
        Assert.NotNull(error.Candidates);
        // The 5 v1.25 keys come first (stable order) followed by the 4 v1.26 keys.
        Assert.Equal(
            new[] { "preset", "width", "height", "depth", "depthColor", "bevelBottom", "contour", "material", "z" },
            error.Candidates!);
    }

    [Fact]
    public void Bevel_OnGroup_IsUnsupportedFeature()
    {
        Create();
        var a = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", JsonValue.Create(2)), ("y", JsonValue.Create(2)))));
        var aId = a["results"]![0]!["target"]!.GetValue<string>().Split("@id=")[1].TrimEnd(']');
        var b = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("shape", "rect"), ("x", JsonValue.Create(10)), ("y", JsonValue.Create(2)))));
        var bId = b["results"]![0]!["target"]!.GetValue<string>().Split("@id=")[1].TrimEnd(']');
        var grouped = Edit(TestEnv.Op("add", "/slide[1]", type: "group", props: TestEnv.Props(
            ("shapes", new JsonArray("@" + aId, "@" + bId)))));
        var groupPath = grouped["results"]![0]!["target"]!.GetValue<string>();

        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", groupPath, props: TestEnv.Props(("bevel", "circle")))]);

        TestEnv.AssertFail(result, ErrorCodes.UnsupportedFeature);
    }

    // ---- (7) idempotent + clear ----

    [Fact]
    public void SetBevelTwice_ReplacesInPlace()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", "circle"))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", "slope"))));

        var properties = SingleShape().ShapeProperties!;
        var sp3d = Assert.Single(properties.Elements<A.Shape3DType>());
        Assert.Equal(A.BevelPresetValues.Slope, sp3d.GetFirstChild<A.BevelTop>()!.Preset!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetBevelFalse_RemovesShape3D_LeavesShapePropertiesOtherwiseUnchanged()
    {
        Create();
        var path = AddShape();
        // An outline gives spPr a child that must survive the bevel clear.
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("outline", "0000FF"), ("bevel", "circle"))));
        Assert.NotNull(SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>());

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", false))));

        var properties = SingleShape().ShapeProperties!;
        Assert.Null(properties.GetFirstChild<A.Shape3DType>());
        Assert.NotNull(properties.GetFirstChild<A.Outline>()); // outline untouched

        var effects = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!;
        Assert.False(effects.AsObject().ContainsKey("bevel"));
        Assert.Equal("0000FF", effects["outline"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Bevel_OnPicture_RoundTrips()
    {
        Create();
        File.WriteAllBytes(Path.Combine(_ws.Dir, "logo.png"), TestImages.Png(40, 30));
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", "logo.png"))));
        var picPath = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", picPath, props: TestEnv.Props(("bevel", "convex"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", picPath))));
        Assert.Equal("convex", detail["effects"]!["bevel"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ======================= v1.26 =======================

    private string Sp3dXml() => SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>()!.OuterXml;

    private JsonNode BevelOf(string path) =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!["bevel"]!.DeepClone()!;

    /// <summary>Every bevel preset token paired with its a:bevelT @prst enum value.</summary>
    public static IEnumerable<object[]> AllBevelPresets() => new[]
    {
        new object[] { "relaxedInset", A.BevelPresetValues.RelaxedInset },
        new object[] { "circle", A.BevelPresetValues.Circle },
        new object[] { "slope", A.BevelPresetValues.Slope },
        new object[] { "cross", A.BevelPresetValues.Cross },
        new object[] { "angle", A.BevelPresetValues.Angle },
        new object[] { "softRound", A.BevelPresetValues.SoftRound },
        new object[] { "convex", A.BevelPresetValues.Convex },
        new object[] { "coolSlant", A.BevelPresetValues.CoolSlant },
        new object[] { "divot", A.BevelPresetValues.Divot },
        new object[] { "riblet", A.BevelPresetValues.Riblet },
        new object[] { "hardEdge", A.BevelPresetValues.HardEdge },
        new object[] { "artDeco", A.BevelPresetValues.ArtDeco },
    };

    // ---- (ACCEPT 1) v1.25 byte-identical ----

    [Theory]
    [MemberData(nameof(AllBevelPresets))]
    public void BarePreset_IsByteIdenticalToV125(string token, A.BevelPresetValues preset)
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", token))));

        // v1.25 wrote exactly: <a:sp3d><a:bevelT prst=... w="76200" h="76200"/></a:sp3d>.
        var expected = new A.Shape3DType(new A.BevelTop { Preset = preset, Width = 76_200L, Height = 76_200L }).OuterXml;
        Assert.Equal(expected, Sp3dXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ObjectSizeDepthColor_IsByteIdenticalToV125()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject
            {
                ["preset"] = "circle",
                ["width"] = "8pt",
                ["height"] = "8pt",
                ["depth"] = "4pt",
                ["depthColor"] = "C00000",
            },
        }));

        // Reconstruct exactly what v1.25 BuildBevel produced for this input.
        var expected = new A.Shape3DType(new A.BevelTop
        {
            Preset = A.BevelPresetValues.Circle,
            Width = 101_600L,
            Height = 101_600L,
        })
        { ExtrusionHeight = 50_800L };
        expected.Append(new A.ExtrusionColor(new A.RgbColorModelHex { Val = "C00000" }));

        Assert.Equal(expected.OuterXml, Sp3dXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void DepthOnly_IsByteIdenticalToV125()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject { ["preset"] = "slope", ["depth"] = "3pt" },
        }));

        var expected = new A.Shape3DType(new A.BevelTop
        {
            Preset = A.BevelPresetValues.Slope,
            Width = 76_200L,
            Height = 76_200L,
        })
        { ExtrusionHeight = 38_100L };

        Assert.Equal(expected.OuterXml, Sp3dXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData(false)]
    [InlineData("")]
    public void ClearForms_RemoveShape3D(object clearValue)
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", "circle"))));
        Assert.NotNull(SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>());

        var value = clearValue is bool b ? JsonValue.Create(b) : JsonValue.Create((string)clearValue);
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", value))));

        Assert.Null(SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (ACCEPT 2) round-trip each new bevel field ----

    [Fact]
    public void BevelBottom_BareAndObject_RoundTripIdempotent()
    {
        Create();
        var path = AddShape();

        // Bare preset => a:bevelB at the 6pt default; reads back as the bare token.
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject { ["preset"] = "circle", ["bevelBottom"] = "slope" },
        }));
        var sp3d = SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>()!;
        var bevelB = sp3d.GetFirstChild<A.BevelBottom>()!;
        Assert.Equal(A.BevelPresetValues.Slope, bevelB.Preset!.Value);
        Assert.Equal(76_200L, bevelB.Width!.Value);
        Assert.Equal(76_200L, bevelB.Height!.Value);
        Assert.Equal("slope", BevelOf(path)["bevelBottom"]!.GetValue<string>());

        var xml1 = Sp3dXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["bevel"] = BevelOf(path) }));
        Assert.Equal(xml1, Sp3dXml());
        TestEnv.AssertValid(_ws, "deck.pptx");

        // Object form with a non-default width => reads back as {preset, width, height}.
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject
            {
                ["preset"] = "circle",
                ["bevelBottom"] = new JsonObject { ["preset"] = "angle", ["width"] = "9pt" },
            },
        }));
        var bb = Assert.IsType<JsonObject>(BevelOf(path)["bevelBottom"]);
        Assert.Equal("angle", bb["preset"]!.GetValue<string>());
        Assert.Equal("9pt", bb["width"]!.GetValue<string>());

        var xml2 = Sp3dXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["bevel"] = BevelOf(path) }));
        Assert.Equal(xml2, Sp3dXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Contour_ColorOnlyAndWithWidth_RoundTripIdempotent()
    {
        Create();
        var path = AddShape();

        // Color-only contour defaults to 1pt (no @contourW=0 trap); width omitted on read.
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject { ["preset"] = "circle", ["contour"] = new JsonObject { ["color"] = "00B050" } },
        }));
        var sp3d = SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>()!;
        Assert.Equal(12_700L, sp3d.ContourWidth!.Value);
        Assert.Equal("00B050", sp3d.GetFirstChild<A.ContourColor>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);

        var contour = Assert.IsType<JsonObject>(BevelOf(path)["contour"]);
        Assert.Equal("00B050", contour["color"]!.GetValue<string>());
        Assert.False(contour.ContainsKey("width"));

        var xml1 = Sp3dXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["bevel"] = BevelOf(path) }));
        Assert.Equal(xml1, Sp3dXml());

        // Explicit width round-trips.
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject
            {
                ["preset"] = "circle",
                ["contour"] = new JsonObject { ["color"] = "0000FF", ["width"] = "2.5pt" },
            },
        }));
        var contour2 = Assert.IsType<JsonObject>(BevelOf(path)["contour"]);
        Assert.Equal("0000FF", contour2["color"]!.GetValue<string>());
        Assert.Equal("2.5pt", contour2["width"]!.GetValue<string>());

        var xml2 = Sp3dXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["bevel"] = BevelOf(path) }));
        Assert.Equal(xml2, Sp3dXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    public static IEnumerable<object[]> AllMaterials() => new[]
    {
        new object[] { "matte", A.PresetMaterialTypeValues.Matte },
        new object[] { "warmMatte", A.PresetMaterialTypeValues.WarmMatte },
        new object[] { "metal", A.PresetMaterialTypeValues.Metal },
        new object[] { "plastic", A.PresetMaterialTypeValues.Plastic },
        new object[] { "powder", A.PresetMaterialTypeValues.Powder },
        new object[] { "translucentPowder", A.PresetMaterialTypeValues.TranslucentPowder },
        new object[] { "clear", A.PresetMaterialTypeValues.Clear },
        new object[] { "flat", A.PresetMaterialTypeValues.Flat },
        new object[] { "dkEdge", A.PresetMaterialTypeValues.DarkEdge },
        new object[] { "softEdge", A.PresetMaterialTypeValues.SoftEdge },
        new object[] { "softmetal", A.PresetMaterialTypeValues.SoftMetal },
    };

    [Theory]
    [MemberData(nameof(AllMaterials))]
    public void Material_EachOfEleven_RoundTripIdempotent(string token, A.PresetMaterialTypeValues expected)
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject { ["preset"] = "circle", ["material"] = token },
        }));

        var sp3d = SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>()!;
        Assert.Equal(expected, sp3d.PresetMaterial!.Value);
        Assert.Equal(token, BevelOf(path)["material"]!.GetValue<string>());

        var xml1 = Sp3dXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["bevel"] = BevelOf(path) }));
        Assert.Equal(xml1, Sp3dXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Z_RoundTripIdempotent()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject { ["preset"] = "circle", ["z"] = "6pt" },
        }));

        var sp3d = SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>()!;
        Assert.Equal(76_200L, sp3d.Z!.Value);
        Assert.Equal("6pt", BevelOf(path)["z"]!.GetValue<string>());

        var xml1 = Sp3dXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["bevel"] = BevelOf(path) }));
        Assert.Equal(xml1, Sp3dXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (ACCEPT 4) child order across all bevel children ----

    [Fact]
    public void AllBevelChildren_AreInSchemaOrder_ValidatorClean()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject
            {
                ["preset"] = "circle",
                ["bevelBottom"] = "angle",
                ["depth"] = "4pt",
                ["depthColor"] = "C00000",
                ["contour"] = new JsonObject { ["color"] = "0000FF", ["width"] = "1.5pt" },
                ["material"] = "metal",
                ["z"] = "3pt",
            },
        }));

        var sp3d = SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>()!;
        var order = sp3d.ChildElements.Select(c => c.LocalName).ToList();
        Assert.Equal(new[] { "bevelT", "bevelB", "extrusionClr", "contourClr" }, order);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (ACCEPT 5) negatives ----

    [Fact]
    public void Material_UnknownToken_IsInvalidArgsWithElevenCandidates()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: new JsonObject
            {
                ["bevel"] = new JsonObject { ["preset"] = "circle", ["material"] = "chrome" },
            })]);

        var error = TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
        Assert.Equal(11, error.Candidates!.Count);
        Assert.Contains("warmMatte", error.Candidates!);
        Assert.Contains("softmetal", error.Candidates!);
    }

    [Theory]
    [InlineData("legacyMatte")]
    [InlineData("legacyPlastic")]
    [InlineData("legacyMetal")]
    [InlineData("legacyWireframe")]
    public void Material_LegacyToken_IsRejected(string legacy)
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: new JsonObject
            {
                ["bevel"] = new JsonObject { ["preset"] = "circle", ["material"] = legacy },
            })]);

        var error = TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
        Assert.Equal(11, error.Candidates!.Count);
        Assert.DoesNotContain(legacy, error.Candidates!);
    }

    [Fact]
    public void NegativeZ_IsInvalidArgs()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: new JsonObject
            {
                ["bevel"] = new JsonObject { ["preset"] = "circle", ["z"] = "-3pt" },
            })]);

        TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void ContourWithoutColor_IsInvalidArgs()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: new JsonObject
            {
                ["bevel"] = new JsonObject { ["preset"] = "circle", ["contour"] = new JsonObject { ["width"] = "2pt" } },
            })]);

        TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
    }

    // ---- (ACCEPT 6) byte-stable read: absent material reads null ----

    [Fact]
    public void BevelWithoutMaterial_ReadsNoMaterialKey()
    {
        Create();
        var path = AddShape();
        // A depth-only bevel projects an object, but with NO material key (absent @prstMaterial != warmMatte).
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["bevel"] = new JsonObject { ["preset"] = "circle", ["depth"] = "4pt" },
        }));

        var bevel = Assert.IsType<JsonObject>(BevelOf(path));
        Assert.False(bevel.ContainsKey("material"));
        Assert.False(bevel.ContainsKey("bevelBottom"));
        Assert.False(bevel.ContainsKey("contour"));
        Assert.False(bevel.ContainsKey("z"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
