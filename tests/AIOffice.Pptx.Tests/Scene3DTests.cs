using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>v1.26 shape scene (a:scene3d): the camera + optional light rig, outside a:effectLst and before a:sp3d.</summary>
public sealed class Scene3DTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() => TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private string AddShape()
    {
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "textbox", props: TestEnv.Props(
            ("text", "Scene"), ("x", "2cm"), ("y", "2cm"), ("w", "8cm"), ("h", "3cm"))));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private P.Shape SingleShape()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        return doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>()
            .First(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("TextBox", StringComparison.Ordinal) == true);
    }

    private string SceneXml() => SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>()!.OuterXml;

    private JsonNode SceneOf(string path) =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!["scene3d"]!.DeepClone()!;

    // ---- (ACCEPT 3) round-trip bare camera string ----

    [Fact]
    public void BareCameraString_WritesCameraOnlyScene_RoundTrips()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("scene3d", "perspectiveFront"))));

        var scene = SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>()!;
        Assert.Equal(A.PresetCameraValues.PerspectiveFront, scene.GetFirstChild<A.Camera>()!.Preset!.Value);
        // The schema requires a light rig, so a camera-only input synthesizes the three-point/Top default.
        var lightRig = scene.GetFirstChild<A.LightRig>()!;
        Assert.Equal(A.LightRigValues.ThreePoints, lightRig.Rig!.Value);
        Assert.Equal(A.LightRigDirectionValues.Top, lightRig.Direction!.Value);

        // ...which the read side collapses, so the bare camera string round-trips.
        Assert.Equal("perspectiveFront", SceneOf(path).GetValue<string>());

        var xml1 = SceneXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["scene3d"] = SceneOf(path) }));
        Assert.Equal(xml1, SceneXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void CameraWithRotation_RoundTripsAsObject()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["scene3d"] = new JsonObject
            {
                ["camera"] = new JsonObject
                {
                    ["preset"] = "perspectiveRelaxedModerately",
                    ["rotation"] = new JsonObject { ["lat"] = 20, ["lon"] = 30, ["rev"] = 0 },
                },
            },
        }));

        var camera = SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>()!.GetFirstChild<A.Camera>()!;
        var rot = camera.GetFirstChild<A.Rotation>()!;
        Assert.Equal(20 * 60_000, rot.Latitude!.Value);
        Assert.Equal(30 * 60_000, rot.Longitude!.Value);
        Assert.Equal(0, rot.Revolution!.Value);

        var scene = Assert.IsType<JsonObject>(SceneOf(path));
        var cam = Assert.IsType<JsonObject>(scene["camera"]);
        Assert.Equal("perspectiveRelaxedModerately", cam["preset"]!.GetValue<string>());
        Assert.Equal(20, cam["rotation"]!["lat"]!.GetValue<int>());
        Assert.Equal(30, cam["rotation"]!["lon"]!.GetValue<int>());

        var xml1 = SceneXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["scene3d"] = SceneOf(path) }));
        Assert.Equal(xml1, SceneXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void LightRigBareString_SynthesizesTopDirection_RoundTripsToBareString()
    {
        Create();
        var path = AddShape();
        // "balanced" (not the synthesized three-point default) proves a bare rig string round-trips to the bare
        // form; the three-point/Top rig is reserved as the camera-only sentinel and collapses on read.
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["scene3d"] = new JsonObject { ["camera"] = "orthographicFront", ["lightRig"] = "balanced" },
        }));

        var lightRig = SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>()!.GetFirstChild<A.LightRig>()!;
        Assert.Equal(A.LightRigValues.Balanced, lightRig.Rig!.Value);
        Assert.Equal(A.LightRigDirectionValues.Top, lightRig.Direction!.Value);

        // Camera reads bare, but the presence of an explicit light rig forces the object form; the rig itself
        // reads back as the bare string (dir == 't', no rotation).
        var scene = Assert.IsType<JsonObject>(SceneOf(path));
        Assert.Equal("orthographicFront", scene["camera"]!.GetValue<string>());
        Assert.Equal("balanced", scene["lightRig"]!.GetValue<string>());

        var xml1 = SceneXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["scene3d"] = SceneOf(path) }));
        Assert.Equal(xml1, SceneXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void LightRigObject_WithDirAndRotation_RoundTrips()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["scene3d"] = new JsonObject
            {
                ["camera"] = "isometricTopUp",
                ["lightRig"] = new JsonObject
                {
                    ["rig"] = "brightRoom",
                    ["dir"] = "tl",
                    ["rotation"] = new JsonObject { ["lat"] = 10, ["lon"] = 350, ["rev"] = 45 },
                },
            },
        }));

        var lightRig = SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>()!.GetFirstChild<A.LightRig>()!;
        Assert.Equal(A.LightRigValues.BrightRoom, lightRig.Rig!.Value);
        Assert.Equal(A.LightRigDirectionValues.TopLeft, lightRig.Direction!.Value);
        var rot = lightRig.GetFirstChild<A.Rotation>()!;
        Assert.Equal(350 * 60_000, rot.Longitude!.Value);
        Assert.Equal(45 * 60_000, rot.Revolution!.Value);

        var scene = Assert.IsType<JsonObject>(SceneOf(path));
        var rig = Assert.IsType<JsonObject>(scene["lightRig"]);
        Assert.Equal("brightRoom", rig["rig"]!.GetValue<string>());
        Assert.Equal("tl", rig["dir"]!.GetValue<string>());
        Assert.Equal(45, rig["rotation"]!["rev"]!.GetValue<int>());

        var xml1 = SceneXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["scene3d"] = SceneOf(path) }));
        Assert.Equal(xml1, SceneXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- camera / rig / dir coverage ----

    // The a:*Values structs have no member-name ToString, so these verify the mapping via clean round-trips
    // (set token -> valid XML -> read the same token back), which exercises both directions of the map.
    [Theory]
    [InlineData("orthographicFront")]
    [InlineData("isometricTopUp")]
    [InlineData("isometricBottomDown")]
    [InlineData("isometricOffAxis1Left")]
    [InlineData("isometricOffAxis3Bottom")]
    [InlineData("isometricOffAxis4Bottom")]
    [InlineData("obliqueTopLeft")]
    [InlineData("obliqueBottomRight")]
    [InlineData("perspectiveAbove")]
    [InlineData("perspectiveHeroicExtremeLeftFacing")]
    [InlineData("perspectiveRelaxedModerately")]
    public void Camera_SpreadOfPresets_RoundTrips(string token)
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("scene3d", token))));

        Assert.NotNull(SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>()!.GetFirstChild<A.Camera>()!.Preset);
        Assert.Equal(token, SceneOf(path).GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // "threePt"/Top is the camera-only sentinel (collapses on read), so it is exercised separately below.
    [Theory]
    [InlineData("twoPt")]
    [InlineData("balanced")]
    [InlineData("harsh")]
    [InlineData("sunrise")]
    [InlineData("freezing")]
    [InlineData("glow")]
    [InlineData("brightRoom")]
    public void LightRig_SpreadOfRigs_RoundTrips(string token)
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["scene3d"] = new JsonObject { ["camera"] = "orthographicFront", ["lightRig"] = token },
        }));

        Assert.NotNull(SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>()!.GetFirstChild<A.LightRig>()!.Rig);
        Assert.Equal(token, SceneOf(path)["lightRig"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("tl")]
    [InlineData("t")]
    [InlineData("tr")]
    [InlineData("l")]
    [InlineData("r")]
    [InlineData("bl")]
    [InlineData("b")]
    [InlineData("br")]
    public void LightRig_EveryDirection_RoundTrips(string token)
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["scene3d"] = new JsonObject
            {
                ["camera"] = "orthographicFront",
                ["lightRig"] = new JsonObject { ["rig"] = "balanced", ["dir"] = token },
            },
        }));

        Assert.NotNull(SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>()!.GetFirstChild<A.LightRig>()!.Direction);

        // 't' reads as the bare rig string; every other dir yields the object with the token echoed back.
        var rig = SceneOf(path)["lightRig"]!;
        if (token == "t")
        {
            Assert.Equal("balanced", rig.GetValue<string>());
        }
        else
        {
            Assert.Equal(token, rig["dir"]!.GetValue<string>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ExplicitThreePointTopRig_CollapsesToBareCameraOnRead()
    {
        // The three-point/Top rig is indistinguishable from the synthesized camera-only default, so an explicit
        // threePt (dir defaults to Top) reads back as the bare camera string. The XML still honors the rig, so
        // re-setting the bare camera reproduces the identical scene.
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["scene3d"] = new JsonObject { ["camera"] = "perspectiveFront", ["lightRig"] = "threePt" },
        }));

        Assert.Equal("perspectiveFront", SceneOf(path).GetValue<string>());

        var xml1 = SceneXml();
        Edit(TestEnv.Op("set", path, props: new JsonObject { ["scene3d"] = SceneOf(path) }));
        Assert.Equal(xml1, SceneXml());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (ACCEPT 4) child order + InsertEffect regression ----

    [Fact]
    public void CombinedShape_LnEffectListSceneSp3d_AreInSchemaOrder_ValidatorClean()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("outline", "0000FF"), ("glow", "00FF00"))));
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["scene3d"] = new JsonObject { ["camera"] = "perspectiveFront", ["lightRig"] = "threePt" },
        }));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", "circle"))));

        var order = SingleShape().ShapeProperties!.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(order.IndexOf("ln") < order.IndexOf("effectLst"));
        Assert.True(order.IndexOf("effectLst") < order.IndexOf("scene3d"));
        Assert.True(order.IndexOf("scene3d") < order.IndexOf("sp3d"));

        var scene = SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>()!;
        Assert.Equal(new[] { "camera", "lightRig" }, scene.ChildElements.Select(c => c.LocalName).ToList());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetScene3DThenShadow_KeepsEffectListBeforeScene3D_ValidatorClean()
    {
        // The InsertEffect fix: with an a:scene3d present, a newly-created a:effectLst must anchor BEFORE it.
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("scene3d", "perspectiveFront"))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "303030"))));

        var order = SingleShape().ShapeProperties!.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(order.IndexOf("effectLst") < order.IndexOf("scene3d"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- idempotent + clear ----

    [Fact]
    public void SetScene3DTwice_ReplacesInPlace()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("scene3d", "perspectiveFront"))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("scene3d", "isometricTopUp"))));

        var scene = Assert.Single(SingleShape().ShapeProperties!.Elements<A.Scene3DType>());
        Assert.Equal(A.PresetCameraValues.IsometricTopUp, scene.GetFirstChild<A.Camera>()!.Preset!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData(false)]
    [InlineData("")]
    public void ClearForms_RemoveScene3D(object clearValue)
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("scene3d", "perspectiveFront"))));
        Assert.NotNull(SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>());

        var value = clearValue is bool b ? JsonValue.Create(b) : JsonValue.Create((string)clearValue);
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("scene3d", value))));

        Assert.Null(SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Scene3DAndBevel_AreOrthogonal_NoAutoEmission()
    {
        Create();
        var path = AddShape();

        // A light rig alone must not synthesize an a:sp3d bevel.
        Edit(TestEnv.Op("set", path, props: new JsonObject
        {
            ["scene3d"] = new JsonObject { ["camera"] = "orthographicFront", ["lightRig"] = "threePt" },
        }));
        Assert.Null(SingleShape().ShapeProperties!.GetFirstChild<A.Shape3DType>());

        // A bevel alone must not synthesize an a:scene3d.
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("scene3d", false))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("bevel", "circle"))));
        Assert.Null(SingleShape().ShapeProperties!.GetFirstChild<A.Scene3DType>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (ACCEPT 6) byte-stable read ----

    [Fact]
    public void ShapeWithoutScene3D_ProjectsNoScene3DKey()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("glow", "00FF00"))));

        var effects = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["effects"]!;
        Assert.False(effects.AsObject().ContainsKey("scene3d"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- (ACCEPT 5) negatives ----

    [Fact]
    public void Scene3DObjectWithoutCamera_IsInvalidArgs()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: new JsonObject
            {
                ["scene3d"] = new JsonObject { ["lightRig"] = "threePt" },
            })]);

        TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void UnknownCameraPreset_IsInvalidArgsWith44Candidates()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("scene3d", "fisheye")))]);

        var error = TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
        Assert.Equal(44, error.Candidates!.Count);
        Assert.Contains("perspectiveFront", error.Candidates!);
    }

    [Theory]
    [InlineData("legacyObliqueTopLeft")]
    [InlineData("legacyPerspectiveFront")]
    public void LegacyCameraPreset_IsRejected(string legacy)
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("scene3d", legacy)))]);

        var error = TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
        Assert.DoesNotContain(legacy, error.Candidates!);
    }

    [Fact]
    public void UnknownLightRig_IsInvalidArgsWith15Candidates()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: new JsonObject
            {
                ["scene3d"] = new JsonObject { ["camera"] = "orthographicFront", ["lightRig"] = "studio" },
            })]);

        var error = TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
        Assert.Equal(15, error.Candidates!.Count);
    }

    [Theory]
    [InlineData("legacyFlat1")]
    [InlineData("legacyHarsh4")]
    [InlineData("legacyNormal2")]
    public void LegacyLightRig_IsRejected(string legacy)
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: new JsonObject
            {
                ["scene3d"] = new JsonObject { ["camera"] = "orthographicFront", ["lightRig"] = legacy },
            })]);

        TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void UnknownLightDirection_IsInvalidArgs()
    {
        Create();
        var path = AddShape();
        var result = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: new JsonObject
            {
                ["scene3d"] = new JsonObject
                {
                    ["camera"] = "orthographicFront",
                    ["lightRig"] = new JsonObject { ["rig"] = "balanced", ["dir"] = "north" },
                },
            })]);

        TestEnv.AssertFail(result, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Scene3DOnGroup_IsUnsupportedFeature()
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
            [TestEnv.Op("set", groupPath, props: TestEnv.Props(("scene3d", "perspectiveFront")))]);

        TestEnv.AssertFail(result, ErrorCodes.UnsupportedFeature);
    }
}
