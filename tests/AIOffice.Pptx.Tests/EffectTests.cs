using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>v1.1 shape visual effects: shadow/glow/reflection (a:effectLst) and outline (a:ln).</summary>
public sealed class EffectTests : IDisposable
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
            ("text", "Effects"), ("x", "2cm"), ("y", "2cm"), ("w", "8cm"), ("h", "3cm"))));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private P.Shape SingleShape()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        return doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>()
            .First(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("TextBox", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void SetShadow_WritesOuterShadowAndReflects()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "404040"))));

        var shadow = SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!.GetFirstChild<A.OuterShadow>()!;
        Assert.Equal("404040", shadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("404040", detail["effects"]!["shadow"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetGlow_WritesGlowAndReflects()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("glow", "FFD700"))));

        var glow = SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!.GetFirstChild<A.Glow>()!;
        Assert.Equal("FFD700", glow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("FFD700", detail["effects"]!["glow"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetReflection_WritesReflectionAndReflects()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("reflection", true))));

        Assert.NotNull(SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!.GetFirstChild<A.Reflection>());

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.True(detail["effects"]!["reflection"]!.GetValue<bool>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetOutline_WritesLineAndReflects()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("outline", "FF0000"))));

        var outline = SingleShape().ShapeProperties!.GetFirstChild<A.Outline>()!;
        Assert.Equal("FF0000", outline.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("FF0000", detail["effects"]!["outline"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetAllEffectsTogether_WritesEffectListInSchemaOrder()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("glow", "00FF00"), ("shadow", "303030"), ("reflection", true), ("outline", "0000FF"))));

        var shape = SingleShape();
        var effectList = shape.ShapeProperties!.GetFirstChild<A.EffectList>()!;
        // a:effectLst schema order: glow precedes outerShdw precedes reflection.
        var order = effectList.ChildElements.Select(e => e.LocalName).ToList();
        Assert.True(order.IndexOf("glow") < order.IndexOf("outerShdw"));
        Assert.True(order.IndexOf("outerShdw") < order.IndexOf("reflection"));
        Assert.NotNull(shape.ShapeProperties.GetFirstChild<A.Outline>());

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        var effects = detail["effects"]!;
        Assert.Equal("00FF00", effects["glow"]!.GetValue<string>());
        Assert.Equal("303030", effects["shadow"]!.GetValue<string>());
        Assert.True(effects["reflection"]!.GetValue<bool>());
        Assert.Equal("0000FF", effects["outline"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetShadowFalse_ClearsTheShadow()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "404040"))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", false))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Null(detail["effects"]);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetShadow_IsIdempotent()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "404040"))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "808080"))));

        // A second shadow set replaces the first, not stacks a duplicate.
        var effectList = SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!;
        Assert.Single(effectList.Elements<A.OuterShadow>());
        Assert.Equal("808080", effectList.GetFirstChild<A.OuterShadow>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- outline widening (v1.15.0): {color?, width?, dash?, compound?} object form ----

    [Fact]
    public void SetOutline_BareString_IsByteStable_AndReadsBackAsString()
    {
        // The legacy bare-string form must be UNCHANGED: a:ln @w=12700, a:solidFill, no prstDash, no @cmpd.
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("outline", "FF0000"))));

        var outline = SingleShape().ShapeProperties!.GetFirstChild<A.Outline>()!;
        Assert.Equal(12_700, outline.Width!.Value);
        Assert.Equal("FF0000", outline.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        Assert.Null(outline.GetFirstChild<A.PresetDash>());
        Assert.Null(outline.CompoundLineType);

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        // Bare-string set round-trips to the bare hex STRING (a JsonValue, not an object).
        Assert.Equal("FF0000", detail["effects"]!["outline"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetOutlineFalse_ClearsTheLineEntirely()
    {
        // Byte-stable legacy clear form: outline:false removes a:ln (the only clear form on the
        // baseline — null/"" keep their legacy meaning, asserted separately below).
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("outline", "FF0000"))));
        Assert.NotNull(SingleShape().ShapeProperties!.GetFirstChild<A.Outline>());

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("outline", false))));

        Assert.Null(SingleShape().ShapeProperties!.GetFirstChild<A.Outline>());
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Null(detail["effects"]);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetOutlineNull_KeepsLegacyDefaultBlackLine_ByteStable()
    {
        // BYTE-STABLE guard: on the baseline, outline:null writes the default 1pt black line (it is
        // NOT a clear form). The object branch must not change this — null is not a JsonObject.
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["outline"] = null }));

        var outline = SingleShape().ShapeProperties!.GetFirstChild<A.Outline>()!;
        Assert.Equal(12_700, outline.Width!.Value);
        Assert.Equal("000000", outline.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        Assert.Null(outline.GetFirstChild<A.PresetDash>());
        Assert.Null(outline.CompoundLineType);

        // get projects the default line as the bare hex string (legacy shape).
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("000000", detail["effects"]!["outline"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetOutline_ObjectWithWidth_WritesEmuWidth()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["outline"] = new JsonObject { ["color"] = "FF0000", ["width"] = "2pt" },
        }));

        var outline = SingleShape().ShapeProperties!.GetFirstChild<A.Outline>()!;
        Assert.Equal(25_400, outline.Width!.Value); // 2pt = 25400 EMU
        Assert.Equal("FF0000", outline.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetOutline_ObjectWithDash_WritesPresetDash()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["outline"] = new JsonObject { ["color"] = "00FF00", ["dash"] = "dashDot" },
        }));

        var outline = SingleShape().ShapeProperties!.GetFirstChild<A.Outline>()!;
        // a:solidFill precedes a:prstDash in the a:ln child order.
        var children = outline.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(children.IndexOf("solidFill") < children.IndexOf("prstDash"));
        Assert.Equal(A.PresetLineDashValues.DashDot, outline.GetFirstChild<A.PresetDash>()!.Val!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetOutline_ObjectWithCompound_WritesCompoundAttribute()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["outline"] = new JsonObject { ["color"] = "0000FF", ["compound"] = "double" },
        }));

        var outline = SingleShape().ShapeProperties!.GetFirstChild<A.Outline>()!;
        Assert.Equal(A.CompoundLineValues.Double, outline.CompoundLineType!.Value); // a:ln @cmpd=dbl
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetOutline_FullObject_RoundTripsAsObject_And_LegacyString_RoundTripsAsString()
    {
        Create();

        // (a) Full object set -> get returns the full {color, width, dash, compound} object.
        var pathA = AddShape();
        Edit(TestEnv.Op("set", pathA, props: new JsonObject
        {
            ["outline"] = new JsonObject
            {
                ["color"] = "112233",
                ["width"] = "2pt",
                ["dash"] = "dashDot",
                ["compound"] = "double",
            },
        }));

        var outlineA = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", pathA))))["effects"]!["outline"]!;
        var obj = Assert.IsType<JsonObject>(outlineA);
        Assert.Equal("112233", obj["color"]!.GetValue<string>());
        Assert.Equal("25400emu", obj["width"]!.GetValue<string>()); // lossless 2pt
        Assert.Equal("dashDot", obj["dash"]!.GetValue<string>());
        Assert.Equal("double", obj["compound"]!.GetValue<string>());

        // The projected width round-trips byte-exactly: re-feeding it yields the same object/EMU.
        Edit(TestEnv.Op("set", pathA, props: new JsonObject
        {
            ["outline"] = new JsonObject { ["color"] = "112233", ["width"] = obj["width"]!.DeepClone() },
        }));
        var reread = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", pathA))))["effects"]!["outline"]!;
        Assert.Equal("25400emu", Assert.IsType<JsonObject>(reread)["width"]!.GetValue<string>());

        // (b) Independent shape: legacy string set -> get returns the bare string.
        var pathB = AddShape();
        Edit(TestEnv.Op("set", pathB, props: TestEnv.Props(("outline", "445566"))));
        var outlineB = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", pathB))))["effects"]!["outline"]!;
        Assert.Equal("445566", outlineB.GetValue<string>());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("dash", "squiggle")]
    [InlineData("compound", "quad")]
    [InlineData("bogus", "anything")]
    public void SetOutline_BadObjectToken_IsInvalidArgsWithCandidates(string key, string value)
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"), new[]
        {
            TestEnv.Op("set", path, props: new JsonObject
            {
                ["outline"] = new JsonObject { ["color"] = "FF0000", [key] = value },
            }),
        });

        var error = TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
        Assert.NotNull(error.Candidates);
        Assert.NotEmpty(error.Candidates!);
    }

    [Fact]
    public void SetOutline_BadWidth_IsInvalidArgs()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"), new[]
        {
            TestEnv.Op("set", path, props: new JsonObject
            {
                ["outline"] = new JsonObject { ["color"] = "FF0000", ["width"] = "wide" },
            }),
        });

        TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Effects_OnPicture_AlsoWork()
    {
        Create();
        File.WriteAllBytes(Path.Combine(_ws.Dir, "logo.png"), TestImages.Png(40, 30));
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", "logo.png"))));
        var picPath = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", picPath, props: TestEnv.Props(("shadow", "202020"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", picPath))));
        Assert.Equal("202020", detail["effects"]!["shadow"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- soft edge (v1.23.0): a:softEdge, the trailing effect in a:effectLst ----

    [Fact]
    public void SetSoftEdgeTrue_WritesDefaultRadius_AndSurvivesReload()
    {
        // true -> a:softEdge @rad=31750 (PowerPoint's 2.5pt default), asserted after save+reload.
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("softEdge", true))));

        var softEdge = SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!.GetFirstChild<A.SoftEdge>()!;
        Assert.Equal(31_750L, softEdge.Radius!.Value);

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("2.5pt", detail["effects"]!["softEdge"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetSoftEdgeSize_WritesRadius_AndReadsBackAlongsideOtherEffects()
    {
        // '5pt' -> @rad=63500; get returns softEdge:'5pt' next to shadow/glow/reflection.
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("glow", "00FF00"), ("shadow", "303030"), ("reflection", true), ("softEdge", "5pt"))));

        var softEdge = SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!.GetFirstChild<A.SoftEdge>()!;
        Assert.Equal(63_500L, softEdge.Radius!.Value); // 5pt * 12700 EMU/pt

        var effects = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!;
        Assert.Equal("5pt", effects["softEdge"]!.GetValue<string>());
        Assert.Equal("00FF00", effects["glow"]!.GetValue<string>());
        Assert.Equal("303030", effects["shadow"]!.GetValue<string>());
        Assert.True(effects["reflection"]!.GetValue<bool>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData(false)]
    [InlineData("")]
    public void SetSoftEdgeFalseOrEmpty_ClearsSoftEdge_AndDropsEmptyEffectList(object clearValue)
    {
        // false or '' removes a:softEdge and drops the now-empty a:effectLst.
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("softEdge", true))));
        Assert.NotNull(SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>());

        var clear = clearValue is bool b ? JsonValue.Create(b) : JsonValue.Create((string)clearValue);
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["softEdge"] = clear }));

        Assert.Null(SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()); // empty list dropped
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Null(detail["effects"]);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetSoftEdge_TrailsReflection_InSchemaChildOrder()
    {
        // softEdge coexists with shadow/glow/reflection; a:softEdge must be the last child.
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("glow", "00FF00"), ("shadow", "303030"), ("reflection", true), ("softEdge", true))));

        var effectList = SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!;
        var order = effectList.ChildElements.Select(e => e.LocalName).ToList();
        Assert.True(order.IndexOf("glow") < order.IndexOf("outerShdw"));
        Assert.True(order.IndexOf("outerShdw") < order.IndexOf("reflection"));
        Assert.True(order.IndexOf("reflection") < order.IndexOf("softEdge"));
        Assert.Equal("softEdge", order[^1]); // softEdge is the trailing anchor
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SoftEdge_OnPicture_RoundTrips()
    {
        Create();
        File.WriteAllBytes(Path.Combine(_ws.Dir, "logo.png"), TestImages.Png(40, 30));
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", "logo.png"))));
        var picPath = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", picPath, props: TestEnv.Props(("softEdge", "5pt"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", picPath))));
        Assert.Equal("5pt", detail["effects"]!["softEdge"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ShapeWithoutSoftEdge_ProjectsNoSoftEdgeKey_ByteStable()
    {
        // A 1.22 shape with only a shadow gains no softEdge key (the null key drops).
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "404040"))));

        var effects = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!;
        Assert.Equal("404040", effects["shadow"]!.GetValue<string>());
        Assert.Null(effects["softEdge"]); // absent, not present-and-null
        Assert.Null(SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!.GetFirstChild<A.SoftEdge>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SoftEdge_OnGroup_IsUnsupportedFeature()
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
            [TestEnv.Op("set", groupPath, props: TestEnv.Props(("softEdge", true)))]);

        TestEnv.AssertFail(result, ErrorCodes.UnsupportedFeature);
    }

    [Fact]
    public void SoftEdge_NonSizeNonBoolValue_IsInvalidArgs()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("softEdge", "banana")))]);

        TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
    }
}
