using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Shape text/visual effects: shadow (a:outerShdw), glow (a:glow), reflection
/// (a:reflection) — all children of the shape's a:effectLst — outline (a:ln)
/// and the first 3-D property, bevel (a:sp3d/a:bevelT), both on the shape's spPr
/// but outside a:effectLst. Each setter accepts a color hex (the effect's
/// color, where it has one) or the boolean true/false (a sensible default, or
/// clearing the effect); get reflects what is set.
/// </summary>
internal static class PptxEffects
{
    /// <summary>Default glow/shadow accent when the caller passes <c>true</c> rather than a color.</summary>
    private const string DefaultEffectColor = "808080";

    // Effect geometry in EMU (PowerPoint's own defaults for these effects).
    private const long ShadowBlur = 40_000L;
    private const long ShadowDistance = 38_100L;
    private const int ShadowDirection = 2_700_000; // 45 degrees down-right (60000ths)
    private const long GlowRadius = 63_500L;
    private const long SoftEdgeRadius = 31_750L; // PowerPoint's 2.5pt default (1pt = 12700 EMU)
    private const double EmuPerPoint = 12_700.0;

    // Inner-shadow geometry (PowerPoint's own defaults) and its color alpha.
    private const long InnerShadowBlur = 63_500L;       // 5pt
    private const long InnerShadowDistance = 50_800L;   // 4pt
    private const int InnerShadowDirection = 2_700_000; // 45 degrees down-right (60000ths)
    private const int InnerShadowAlpha = 50_000;        // 50% opacity (a:alpha)
    private const double DirUnitsPerDegree = 60_000.0;  // OOXML angles are 60000ths of a degree

    // ---- shadow -------------------------------------------------------------

    /// <summary>Sets (or clears, on false) an outer shadow; a color value tints it.</summary>
    public static void SetShadow(ShapeView view, JsonNode? value)
    {
        if (IsClear(value, out var hex, DefaultEffectColor))
        {
            RemoveEffect<A.OuterShadow>(view);
            return;
        }

        var shadow = new A.OuterShadow(
            new A.RgbColorModelHex(new A.Alpha { Val = 40000 }) { Val = hex })
        {
            BlurRadius = ShadowBlur,
            Distance = ShadowDistance,
            Direction = ShadowDirection,
            RotateWithShape = false,
        };

        // outerShdw follows glow but precedes reflection in a:effectLst.
        InsertEffect(view, shadow, after: ["glow", "innerShdw"], before: ["reflection", "softEdge"]);
    }

    // ---- glow ---------------------------------------------------------------

    /// <summary>Sets (or clears, on false) a glow; a color value tints it.</summary>
    public static void SetGlow(ShapeView view, JsonNode? value)
    {
        if (IsClear(value, out var hex, "FFD966"))
        {
            RemoveEffect<A.Glow>(view);
            return;
        }

        var glow = new A.Glow(new A.RgbColorModelHex(new A.Alpha { Val = 60000 }) { Val = hex })
        {
            Radius = GlowRadius,
        };

        // glow is the first of the effects aioffice writes (precedes outerShdw).
        InsertEffect(view, glow, after: [], before: ["innerShdw", "outerShdw", "reflection", "softEdge"]);
    }

    // ---- reflection ---------------------------------------------------------

    /// <summary>Sets (or clears, on false) a reflection. Reflection carries no color of its own.</summary>
    public static void SetReflection(ShapeView view, JsonNode? value)
    {
        if (IsFalse(value))
        {
            RemoveEffect<A.Reflection>(view);
            return;
        }

        var reflection = new A.Reflection
        {
            BlurRadius = 6_350L,
            StartOpacity = 52000,
            StartPosition = 0,
            EndAlpha = 300,
            EndPosition = 35000,
            Distance = 0L,
            Direction = 5_400_000,
            FadeDirection = 5_400_000,
            RotateWithShape = false,
        };

        // reflection is last among the effects aioffice writes.
        InsertEffect(view, reflection, after: ["glow", "innerShdw", "outerShdw"], before: ["softEdge"]);
    }

    // ---- soft edge ----------------------------------------------------------

    /// <summary>
    /// Sets (or clears, on false/"") a soft edge (a:softEdge). <c>true</c> uses PowerPoint's
    /// 2.5pt default radius; a size string like "5pt" tunes it. Soft edge carries no color.
    /// </summary>
    public static void SetSoftEdge(ShapeView view, JsonNode? value)
    {
        if (IsFalse(value) || IsEmptyString(value))
        {
            RemoveEffect<A.SoftEdge>(view);
            return;
        }

        var softEdge = new A.SoftEdge { Radius = SoftEdgeRadiusOf(value) };

        // softEdge is the trailing effect in a:effectLst (after reflection).
        InsertEffect(view, softEdge, after: ["glow", "innerShdw", "outerShdw", "reflection"], before: []);
    }

    /// <summary>The soft-edge radius in EMU: the 2.5pt default for true/null, else the parsed size.</summary>
    private static long SoftEdgeRadiusOf(JsonNode? value)
    {
        if (value is null)
        {
            return SoftEdgeRadius; // "on, default radius"
        }

        if (value is JsonValue jv)
        {
            if (jv.TryGetValue<bool>(out _))
            {
                return SoftEdgeRadius; // true (false handled by the caller)
            }

            if (jv.TryGetValue<string>(out var text) &&
                string.Equals(text.Trim(), "true", StringComparison.OrdinalIgnoreCase))
            {
                return SoftEdgeRadius;
            }
        }

        return Units.ParseLengthEmu("softEdge", value);
    }

    // ---- inner shadow -------------------------------------------------------

    /// <summary>
    /// Sets (or clears, on false/"") an inner shadow (a:innerShdw). A bare color hex/name tints it
    /// (default black); <c>true</c> uses PowerPoint's default geometry; a {color?, blur?, dist?, dir?}
    /// object tunes blur radius, distance and direction. Inner shadow has no rotate-with-shape.
    /// </summary>
    public static void SetInnerShadow(ShapeView view, JsonNode? value)
    {
        if (IsFalse(value) || IsEmptyString(value))
        {
            RemoveEffect<A.InnerShadow>(view);
            return;
        }

        var inner = value is JsonObject obj
            ? BuildInnerShadow(obj)
            : new A.InnerShadow(
                new A.RgbColorModelHex(new A.Alpha { Val = InnerShadowAlpha }) { Val = ColorOrDefault(value, "000000") })
            {
                BlurRadius = InnerShadowBlur,
                Distance = InnerShadowDistance,
                Direction = InnerShadowDirection,
            };

        // innerShdw follows glow but precedes outerShdw in a:effectLst.
        InsertEffect(view, inner, after: ["glow"], before: ["outerShdw", "reflection", "softEdge"]);
    }

    /// <summary>The keys the inner-shadow object form accepts.</summary>
    private static readonly IReadOnlyList<string> InnerShadowKeys = ["color", "blur", "dist", "dir"];

    /// <summary>Builds an a:innerShdw from the {color?, blur?, dist?, dir?} object form.</summary>
    private static A.InnerShadow BuildInnerShadow(JsonObject obj)
    {
        foreach (var (key, _) in obj)
        {
            if (!InnerShadowKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown inner shadow prop '{key}'.",
                    "inner shadow object props: color, blur, dist, dir.",
                    candidates: InnerShadowKeys);
            }
        }

        var hex = obj.TryGetPropertyValue("color", out var colorNode)
            ? Units.ParseColorHex("color", colorNode)
            : "000000";
        return new A.InnerShadow(
            new A.RgbColorModelHex(new A.Alpha { Val = InnerShadowAlpha }) { Val = hex })
        {
            BlurRadius = obj.TryGetPropertyValue("blur", out var blurNode)
                ? Units.ParseLengthEmu("blur", blurNode)
                : InnerShadowBlur,
            Distance = obj.TryGetPropertyValue("dist", out var distNode)
                ? Units.ParseLengthEmu("dist", distNode)
                : InnerShadowDistance,
            Direction = obj.TryGetPropertyValue("dir", out var dirNode)
                ? ParseDirection(dirNode)
                : InnerShadowDirection,
        };
    }

    /// <summary>Parses a direction in degrees to OOXML's 60000ths-of-a-degree, normalized into [0,360).</summary>
    private static int ParseDirection(JsonNode? node)
    {
        if (node is not JsonValue value || !Units.TryNumber(value, out var degrees))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"The inner shadow 'dir' is not a number: {node?.ToJsonString() ?? "null"}.",
                "Pass dir in degrees, e.g. 45, 135, 315.");
        }

        var normalized = ((degrees % 360) + 360) % 360;
        return (int)Math.Round(normalized * DirUnitsPerDegree);
    }

    // ---- bevel / 3-D --------------------------------------------------------

    /// <summary>Default bevel width/height when only a preset is given (PowerPoint's 6pt).</summary>
    private const long DefaultBevelSize = 76_200L; // 6pt

    /// <summary>Default contour width when a color-only contour is given (PowerPoint's 1pt) — so it never renders invisibly.</summary>
    private const long DefaultContourWidth = 12_700L; // 1pt

    /// <summary>
    /// Sets (or clears, on false/"") a shape bevel — the first 3-D property, a:sp3d/a:bevelT, which
    /// lives OUTSIDE a:effectLst. A bare preset string (e.g. "circle") writes the preset at the 6pt
    /// default; a {preset?, width?, height?, depth?, depthColor?} object tunes the bevel size, the
    /// extrusion depth (a:sp3d/@extrusionH) and its color (a:extrusionClr). bevelB/contour/material/
    /// scene3d are out of scope for this slot.
    /// </summary>
    public static void SetBevel(ShapeView view, JsonNode? value)
    {
        if (IsFalse(value) || IsEmptyString(value))
        {
            RequireShapeProperties(view).GetFirstChild<A.Shape3DType>()?.Remove();
            return;
        }

        var sp3d = value is JsonObject obj
            ? BuildBevel(obj)
            : new A.Shape3DType(new A.BevelTop { Preset = ParseBevelPreset(value), Width = DefaultBevelSize, Height = DefaultBevelSize });

        InsertShape3D(view, sp3d);
    }

    /// <summary>
    /// The keys the bevel object form accepts. The five v1.25 keys come FIRST (stable error-candidate order);
    /// the four v1.26 keys — bevelBottom (a:bevelB), contour (a:contourClr + @contourW), material (@prstMaterial)
    /// and z (@z) — extend the SAME single a:sp3d the bevel has always owned.
    /// </summary>
    private static readonly IReadOnlyList<string> BevelKeys =
        ["preset", "width", "height", "depth", "depthColor", "bevelBottom", "contour", "material", "z"];

    /// <summary>
    /// Builds an a:sp3d from the {preset?, width?, height?, depth?, depthColor?, bevelBottom?, contour?,
    /// material?, z?} object form. Child APPEND order honors the schema: a:bevelT (ctor), then a:bevelB,
    /// then a:extrusionClr, then a:contourClr; @z/@extrusionH/@contourW/@prstMaterial are attributes of a:sp3d
    /// (the SDK serializes them in schema order regardless of assignment order). For any v1.25 input this
    /// reduces to [a:bevelT, (a:extrusionClr)] with only @extrusionH — bit-identical to what v1.25 wrote.
    /// </summary>
    private static A.Shape3DType BuildBevel(JsonObject obj)
    {
        foreach (var (key, _) in obj)
        {
            if (!BevelKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown bevel prop '{key}'.",
                    "bevel object props: preset, width, height, depth, depthColor, bevelBottom, contour, material, z.",
                    candidates: BevelKeys);
            }
        }

        var bevelTop = new A.BevelTop
        {
            Preset = obj.TryGetPropertyValue("preset", out var presetNode)
                ? ParseBevelPreset(presetNode)
                : A.BevelPresetValues.Circle,
            Width = obj.TryGetPropertyValue("width", out var widthNode)
                ? Units.ParseLengthEmu("width", widthNode)
                : DefaultBevelSize,
            Height = obj.TryGetPropertyValue("height", out var heightNode)
                ? Units.ParseLengthEmu("height", heightNode)
                : DefaultBevelSize,
        };

        var sp3d = new A.Shape3DType(bevelTop);

        // @z — a length; ST_Coordinate is signed but a bevel depth-below-surface is nonsensical, so reject negatives.
        if (obj.TryGetPropertyValue("z", out var zNode))
        {
            var z = Units.ParseLengthEmu("z", zNode);
            if (z < 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"The bevel 'z' must be non-negative: {zNode?.ToJsonString() ?? "null"}.",
                    "Pass z as a non-negative length, e.g. 0, \"6pt\", \"2mm\".");
            }

            sp3d.Z = z;
        }

        if (obj.TryGetPropertyValue("depth", out var depthNode))
        {
            sp3d.ExtrusionHeight = Units.ParseLengthEmu("depth", depthNode);
        }

        if (obj.TryGetPropertyValue("material", out var materialNode))
        {
            sp3d.PresetMaterial = ParseMaterial(materialNode);
        }

        // contour contributes both @contourW (attr) and a:contourClr (child, appended below in schema order).
        A.ContourColor? contourColor = null;
        if (obj.TryGetPropertyValue("contour", out var contourNode))
        {
            var (contourHex, contourWidth) = ParseContour(contourNode);
            sp3d.ContourWidth = contourWidth;
            contourColor = new A.ContourColor(new A.RgbColorModelHex { Val = contourHex });
        }

        // Child order: a:bevelT (ctor) -> a:bevelB -> a:extrusionClr -> a:contourClr.
        if (obj.TryGetPropertyValue("bevelBottom", out var bevelBottomNode))
        {
            sp3d.Append(BuildBevelBottom(bevelBottomNode));
        }

        if (obj.TryGetPropertyValue("depthColor", out var depthColorNode))
        {
            sp3d.Append(new A.ExtrusionColor(new A.RgbColorModelHex { Val = Units.ParseColorHex("depthColor", depthColorNode) }));
        }

        if (contourColor is not null)
        {
            sp3d.Append(contourColor);
        }

        return sp3d;
    }

    /// <summary>The keys the bevelBottom object form accepts (mirrors the bevelT preset/size shape).</summary>
    private static readonly IReadOnlyList<string> BevelBottomKeys = ["preset", "width", "height"];

    /// <summary>
    /// Builds an a:bevelB from a bare preset string (preset at the 6pt default) or a {preset?, width?, height?}
    /// object. The 12 presets are the same set as bevelT.
    /// </summary>
    private static A.BevelBottom BuildBevelBottom(JsonNode? node)
    {
        if (node is JsonObject obj)
        {
            foreach (var (key, _) in obj)
            {
                if (!BevelBottomKeys.Contains(key, StringComparer.Ordinal))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Unknown bevelBottom prop '{key}'.",
                        "bevelBottom object props: preset, width, height.",
                        candidates: BevelBottomKeys);
                }
            }

            return new A.BevelBottom
            {
                Preset = obj.TryGetPropertyValue("preset", out var presetNode)
                    ? ParseBevelPreset(presetNode)
                    : A.BevelPresetValues.Circle,
                Width = obj.TryGetPropertyValue("width", out var widthNode)
                    ? Units.ParseLengthEmu("width", widthNode)
                    : DefaultBevelSize,
                Height = obj.TryGetPropertyValue("height", out var heightNode)
                    ? Units.ParseLengthEmu("height", heightNode)
                    : DefaultBevelSize,
            };
        }

        return new A.BevelBottom { Preset = ParseBevelPreset(node), Width = DefaultBevelSize, Height = DefaultBevelSize };
    }

    /// <summary>The keys the contour object form accepts.</summary>
    private static readonly IReadOnlyList<string> ContourKeys = ["color", "width"];

    /// <summary>
    /// Parses the contour {color, width?} object into (hex, widthEmu). color is REQUIRED; width defaults to
    /// 1pt so a color-only contour never renders invisibly (the @contourW=0 trap). srgbClr only (no theme).
    /// </summary>
    private static (string Hex, long Width) ParseContour(JsonNode? node)
    {
        if (node is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"The bevel 'contour' must be an object with a color: {node?.ToJsonString() ?? "null"}.",
                "Pass contour as {\"color\":\"C00000\"} or {\"color\":\"C00000\",\"width\":\"2pt\"}.",
                candidates: ContourKeys);
        }

        foreach (var (key, _) in obj)
        {
            if (!ContourKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown contour prop '{key}'.",
                    "contour object props: color, width.",
                    candidates: ContourKeys);
            }
        }

        if (!obj.TryGetPropertyValue("color", out var colorNode) || colorNode is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "The bevel 'contour' requires a color.",
                "Pass contour as {\"color\":\"C00000\"} (width defaults to 1pt).",
                candidates: ContourKeys);
        }

        var hex = Units.ParseColorHex("color", colorNode);
        var width = obj.TryGetPropertyValue("width", out var widthNode)
            ? Units.ParseLengthEmu("width", widthNode)
            : DefaultContourWidth;
        return (hex, width);
    }

    /// <summary>The 11 preset-material tokens the bevel 'material' accepts (a:sp3d @prstMaterial), in candidate order.</summary>
    private static readonly IReadOnlyList<string> MaterialTokens =
        ["matte", "warmMatte", "metal", "plastic", "powder", "translucentPowder",
         "clear", "flat", "dkEdge", "softEdge", "softmetal"];

    /// <summary>
    /// Maps a material token to a:sp3d @prstMaterial. Accepts the 11 modern materials; rejects the 4 legacy
    /// ones (legacyMatte/legacyPlastic/legacyMetal/legacyWireframe) with invalid_args + the 11 candidates.
    /// </summary>
    private static A.PresetMaterialTypeValues ParseMaterial(JsonNode? node)
    {
        var raw = node is null ? string.Empty : J.ScalarText(node).Trim();
        return raw switch
        {
            "matte" => A.PresetMaterialTypeValues.Matte,
            "warmMatte" => A.PresetMaterialTypeValues.WarmMatte,
            "metal" => A.PresetMaterialTypeValues.Metal,
            "plastic" => A.PresetMaterialTypeValues.Plastic,
            "powder" => A.PresetMaterialTypeValues.Powder,
            "translucentPowder" => A.PresetMaterialTypeValues.TranslucentPowder,
            "clear" => A.PresetMaterialTypeValues.Clear,
            "flat" => A.PresetMaterialTypeValues.Flat,
            "dkEdge" => A.PresetMaterialTypeValues.DarkEdge,
            "softEdge" => A.PresetMaterialTypeValues.SoftEdge,
            "softmetal" => A.PresetMaterialTypeValues.SoftMetal,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid bevel material: {node?.ToJsonString() ?? "null"}",
                "bevel materials: matte, warmMatte, metal, plastic, powder, translucentPowder, clear, flat, dkEdge, softEdge, softmetal.",
                candidates: MaterialTokens),
        };
    }

    /// <summary>Maps an a:sp3d @prstMaterial back to its material token; null when absent or a legacy material.</summary>
    private static string? MaterialToken(A.PresetMaterialTypeValues? val)
    {
        if (val is null)
        {
            return null;
        }

        return val.Value switch
        {
            _ when val.Value == A.PresetMaterialTypeValues.Matte => "matte",
            _ when val.Value == A.PresetMaterialTypeValues.WarmMatte => "warmMatte",
            _ when val.Value == A.PresetMaterialTypeValues.Metal => "metal",
            _ when val.Value == A.PresetMaterialTypeValues.Plastic => "plastic",
            _ when val.Value == A.PresetMaterialTypeValues.Powder => "powder",
            _ when val.Value == A.PresetMaterialTypeValues.TranslucentPowder => "translucentPowder",
            _ when val.Value == A.PresetMaterialTypeValues.Clear => "clear",
            _ when val.Value == A.PresetMaterialTypeValues.Flat => "flat",
            _ when val.Value == A.PresetMaterialTypeValues.DarkEdge => "dkEdge",
            _ when val.Value == A.PresetMaterialTypeValues.SoftEdge => "softEdge",
            _ when val.Value == A.PresetMaterialTypeValues.SoftMetal => "softmetal",
            _ => null,
        };
    }

    /// <summary>Inserts (replacing any existing) a:sp3d into spPr: after a:ln/a:effectLst/a:scene3d, before a:extLst.</summary>
    private static void InsertShape3D(ShapeView view, A.Shape3DType sp3d)
    {
        var properties = RequireShapeProperties(view);
        properties.GetFirstChild<A.Shape3DType>()?.Remove(); // idempotent replace

        var extLst = (OpenXmlElement?)properties.GetFirstChild<A.ExtensionList>();
        if (extLst is not null)
        {
            properties.InsertBefore(sp3d, extLst);
        }
        else
        {
            properties.Append(sp3d);
        }
    }

    /// <summary>The bevel presets accepted (a:bevelT @prst), in error-candidate order.</summary>
    private static readonly IReadOnlyList<string> BevelPresets =
        ["relaxedInset", "circle", "slope", "cross", "angle", "softRound",
         "convex", "coolSlant", "divot", "riblet", "hardEdge", "artDeco"];

    /// <summary>Maps a bevel preset token to its a:bevelT @prst value; throws invalid_args with candidates otherwise.</summary>
    private static A.BevelPresetValues ParseBevelPreset(JsonNode? node)
    {
        var raw = node is null ? string.Empty : J.ScalarText(node).Trim();
        return raw switch
        {
            "relaxedInset" => A.BevelPresetValues.RelaxedInset,
            "circle" => A.BevelPresetValues.Circle,
            "slope" => A.BevelPresetValues.Slope,
            "cross" => A.BevelPresetValues.Cross,
            "angle" => A.BevelPresetValues.Angle,
            "softRound" => A.BevelPresetValues.SoftRound,
            "convex" => A.BevelPresetValues.Convex,
            "coolSlant" => A.BevelPresetValues.CoolSlant,
            "divot" => A.BevelPresetValues.Divot,
            "riblet" => A.BevelPresetValues.Riblet,
            "hardEdge" => A.BevelPresetValues.HardEdge,
            "artDeco" => A.BevelPresetValues.ArtDeco,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid bevel preset: {node?.ToJsonString() ?? "null"}",
                "bevel presets: relaxedInset, circle, slope, cross, angle, softRound, convex, coolSlant, divot, riblet, hardEdge, artDeco.",
                candidates: BevelPresets),
        };
    }

    // ---- scene3d ------------------------------------------------------------

    /// <summary>
    /// Sets (or clears, on false/"") a shape scene (a:scene3d) — the camera + optional light rig that frame the
    /// 3-D shape, living OUTSIDE a:effectLst and BEFORE a:sp3d in spPr. A bare camera preset string
    /// (e.g. "perspectiveFront") writes a camera-only scene; a {camera, lightRig?} object adds a light rig and/or
    /// rotations. scene3d and bevel are orthogonal — setting one never synthesizes the other.
    /// </summary>
    public static void SetScene3D(ShapeView view, JsonNode? value)
    {
        if (IsFalse(value) || IsEmptyString(value))
        {
            RequireShapeProperties(view).GetFirstChild<A.Scene3DType>()?.Remove();
            return;
        }

        InsertScene3D(view, BuildScene3D(value));
    }

    /// <summary>The keys the scene3d object form accepts.</summary>
    private static readonly IReadOnlyList<string> Scene3DKeys = ["camera", "lightRig"];

    /// <summary>
    /// Builds an a:scene3d from a bare camera preset string or a {camera, lightRig?} object. a:scene3d REQUIRES
    /// a camera; the light rig is optional in the input but REQUIRED by the schema, so when the caller omits it
    /// we synthesize the default rig (three-point, lit from the top) — which the read side collapses back so a
    /// camera-only input round-trips to the bare camera string. Child order: a:camera then a:lightRig.
    /// </summary>
    private static A.Scene3DType BuildScene3D(JsonNode? value)
    {
        A.Camera camera;
        A.LightRig? lightRig = null;

        if (value is JsonObject obj)
        {
            foreach (var (key, _) in obj)
            {
                if (!Scene3DKeys.Contains(key, StringComparer.Ordinal))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Unknown scene3d prop '{key}'.",
                        "scene3d object props: camera, lightRig.",
                        candidates: Scene3DKeys);
                }
            }

            if (!obj.TryGetPropertyValue("camera", out var cameraNode) || cameraNode is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "A scene3d requires a camera.",
                    "Pass scene3d as a camera preset string, or {\"camera\":\"perspectiveFront\", \"lightRig\":\"threePt\"}.",
                    candidates: Scene3DKeys);
            }

            camera = BuildCamera(cameraNode);
            if (obj.TryGetPropertyValue("lightRig", out var lightRigNode))
            {
                lightRig = BuildLightRig(lightRigNode);
            }
        }
        else
        {
            camera = BuildCamera(value);
        }

        return new A.Scene3DType(camera, lightRig ?? DefaultLightRig());
    }

    /// <summary>
    /// The schema-required default light rig for a camera-only scene: three-point lighting from the top. The
    /// read side treats this exact rig (three-point, Top, no rotation) as "no explicit light rig".
    /// </summary>
    private static A.LightRig DefaultLightRig() =>
        new() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };

    /// <summary>True when a rig is the synthesized default (three-point, Top, no rotation) — read collapses it.</summary>
    private static bool IsDefaultLightRig(A.LightRig rig) =>
        rig.Rig?.Value == A.LightRigValues.ThreePoints &&
        rig.Direction?.Value == A.LightRigDirectionValues.Top &&
        rig.GetFirstChild<A.Rotation>() is null;

    /// <summary>The keys the camera object form accepts.</summary>
    private static readonly IReadOnlyList<string> CameraKeys = ["preset", "rotation"];

    /// <summary>Builds an a:camera from a bare preset string or a {preset, rotation?} object.</summary>
    private static A.Camera BuildCamera(JsonNode? node)
    {
        if (node is JsonObject obj)
        {
            foreach (var (key, _) in obj)
            {
                if (!CameraKeys.Contains(key, StringComparer.Ordinal))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Unknown camera prop '{key}'.",
                        "camera object props: preset, rotation.",
                        candidates: CameraKeys);
                }
            }

            if (!obj.TryGetPropertyValue("preset", out var presetNode) || presetNode is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "A scene3d camera requires a preset.",
                    "Pass camera as a preset string, or {\"preset\":\"perspectiveFront\", \"rotation\":{\"lat\":20,\"lon\":30}}.",
                    candidates: CameraKeys);
            }

            var camera = new A.Camera { Preset = ParseCameraPreset(presetNode) };
            if (obj.TryGetPropertyValue("rotation", out var rotationNode))
            {
                camera.Append(BuildRotation(rotationNode));
            }

            return camera;
        }

        return new A.Camera { Preset = ParseCameraPreset(node) };
    }

    /// <summary>The keys the lightRig object form accepts.</summary>
    private static readonly IReadOnlyList<string> LightRigKeys = ["rig", "dir", "rotation"];

    /// <summary>
    /// Builds an a:lightRig from a bare rig string (dir synthesized as 't'=Top) or a {rig, dir?, rotation?}
    /// object. a:lightRig requires both @rig and @dir.
    /// </summary>
    private static A.LightRig BuildLightRig(JsonNode? node)
    {
        if (node is JsonObject obj)
        {
            foreach (var (key, _) in obj)
            {
                if (!LightRigKeys.Contains(key, StringComparer.Ordinal))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Unknown lightRig prop '{key}'.",
                        "lightRig object props: rig, dir, rotation.",
                        candidates: LightRigKeys);
                }
            }

            if (!obj.TryGetPropertyValue("rig", out var rigNode) || rigNode is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "A scene3d lightRig requires a rig.",
                    "Pass lightRig as a rig string, or {\"rig\":\"threePt\", \"dir\":\"tl\"}.",
                    candidates: LightRigKeys);
            }

            var lightRig = new A.LightRig
            {
                Rig = ParseLightRig(rigNode),
                Direction = obj.TryGetPropertyValue("dir", out var dirNode)
                    ? ParseLightDirection(dirNode)
                    : A.LightRigDirectionValues.Top,
            };
            if (obj.TryGetPropertyValue("rotation", out var rotationNode))
            {
                lightRig.Append(BuildRotation(rotationNode));
            }

            return lightRig;
        }

        return new A.LightRig { Rig = ParseLightRig(node), Direction = A.LightRigDirectionValues.Top };
    }

    /// <summary>The keys the rotation object form accepts.</summary>
    private static readonly IReadOnlyList<string> RotationKeys = ["lat", "lon", "rev"];

    /// <summary>
    /// Builds an a:rot from {lat?, lon?, rev?} degrees. All three attributes are required by the schema; an
    /// omitted axis is 0. Angles are normalized into [0,360) and written as 60000ths of a degree (like PptxFill).
    /// </summary>
    private static A.Rotation BuildRotation(JsonNode? node)
    {
        if (node is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A scene3d 'rotation' must be an object: {node?.ToJsonString() ?? "null"}.",
                "Pass rotation as {\"lat\":20, \"lon\":30, \"rev\":0} (degrees; any axis may be omitted).",
                candidates: RotationKeys);
        }

        foreach (var (key, _) in obj)
        {
            if (!RotationKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown rotation prop '{key}'.",
                    "rotation object props: lat, lon, rev.",
                    candidates: RotationKeys);
            }
        }

        return new A.Rotation
        {
            Latitude = RotationAxis(obj, "lat"),
            Longitude = RotationAxis(obj, "lon"),
            Revolution = RotationAxis(obj, "rev"),
        };
    }

    /// <summary>Reads one rotation axis (degrees) as 60000ths, normalized into [0,360); an absent axis is 0.</summary>
    private static int RotationAxis(JsonObject obj, string key)
    {
        if (!obj.TryGetPropertyValue(key, out var node) || node is null)
        {
            return 0;
        }

        if (node is not JsonValue value || !Units.TryNumber(value, out var degrees))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"The rotation '{key}' is not a number: {node.ToJsonString()}.",
                "Pass the axis in degrees, e.g. 20, 45, 315.");
        }

        var normalized = ((degrees % 360) + 360) % 360;
        return (int)Math.Round(normalized * DirUnitsPerDegree);
    }

    /// <summary>Inserts (replacing any existing) a:scene3d into spPr: after a:ln/a:effectLst, before a:sp3d/a:extLst.</summary>
    private static void InsertScene3D(ShapeView view, A.Scene3DType scene)
    {
        var properties = RequireShapeProperties(view);
        properties.GetFirstChild<A.Scene3DType>()?.Remove(); // idempotent replace

        var anchor = (OpenXmlElement?)properties.GetFirstChild<A.Shape3DType>()
            ?? properties.GetFirstChild<A.ExtensionList>();
        if (anchor is not null)
        {
            properties.InsertBefore(scene, anchor);
        }
        else
        {
            properties.Append(scene);
        }
    }

    /// <summary>The 44 camera presets accepted (a:camera @prst), in error-candidate order.</summary>
    private static readonly IReadOnlyList<string> CameraPresets =
        ["orthographicFront",
         "isometricTopUp", "isometricTopDown", "isometricBottomUp", "isometricBottomDown",
         "isometricLeftUp", "isometricLeftDown", "isometricRightUp", "isometricRightDown",
         "isometricOffAxis1Left", "isometricOffAxis1Right", "isometricOffAxis1Top",
         "isometricOffAxis2Left", "isometricOffAxis2Right", "isometricOffAxis2Top",
         "isometricOffAxis3Left", "isometricOffAxis3Right", "isometricOffAxis3Bottom",
         "isometricOffAxis4Left", "isometricOffAxis4Right", "isometricOffAxis4Bottom",
         "obliqueTopLeft", "obliqueTop", "obliqueTopRight", "obliqueLeft", "obliqueRight",
         "obliqueBottomLeft", "obliqueBottom", "obliqueBottomRight",
         "perspectiveFront", "perspectiveLeft", "perspectiveRight", "perspectiveAbove", "perspectiveBelow",
         "perspectiveAboveLeftFacing", "perspectiveAboveRightFacing",
         "perspectiveContrastingLeftFacing", "perspectiveContrastingRightFacing",
         "perspectiveHeroicLeftFacing", "perspectiveHeroicRightFacing",
         "perspectiveHeroicExtremeLeftFacing", "perspectiveHeroicExtremeRightFacing",
         "perspectiveRelaxed", "perspectiveRelaxedModerately"];

    /// <summary>Maps a camera-preset token to a:camera @prst; rejects the 18 legacy presets with candidates.</summary>
    private static A.PresetCameraValues ParseCameraPreset(JsonNode? node)
    {
        var raw = node is null ? string.Empty : J.ScalarText(node).Trim();
        return raw switch
        {
            "orthographicFront" => A.PresetCameraValues.OrthographicFront,
            "isometricTopUp" => A.PresetCameraValues.IsometricTopUp,
            "isometricTopDown" => A.PresetCameraValues.IsometricTopDown,
            "isometricBottomUp" => A.PresetCameraValues.IsometricBottomUp,
            "isometricBottomDown" => A.PresetCameraValues.IsometricBottomDown,
            "isometricLeftUp" => A.PresetCameraValues.IsometricLeftUp,
            "isometricLeftDown" => A.PresetCameraValues.IsometricLeftDown,
            "isometricRightUp" => A.PresetCameraValues.IsometricRightUp,
            "isometricRightDown" => A.PresetCameraValues.IsometricRightDown,
            "isometricOffAxis1Left" => A.PresetCameraValues.IsometricOffAxis1Left,
            "isometricOffAxis1Right" => A.PresetCameraValues.IsometricOffAxis1Right,
            "isometricOffAxis1Top" => A.PresetCameraValues.IsometricOffAxis1Top,
            "isometricOffAxis2Left" => A.PresetCameraValues.IsometricOffAxis2Left,
            "isometricOffAxis2Right" => A.PresetCameraValues.IsometricOffAxis2Right,
            "isometricOffAxis2Top" => A.PresetCameraValues.IsometricOffAxis2Top,
            "isometricOffAxis3Left" => A.PresetCameraValues.IsometricOffAxis3Left,
            "isometricOffAxis3Right" => A.PresetCameraValues.IsometricOffAxis3Right,
            "isometricOffAxis3Bottom" => A.PresetCameraValues.IsometricOffAxis3Bottom,
            "isometricOffAxis4Left" => A.PresetCameraValues.IsometricOffAxis4Left,
            "isometricOffAxis4Right" => A.PresetCameraValues.IsometricOffAxis4Right,
            "isometricOffAxis4Bottom" => A.PresetCameraValues.IsometricOffAxis4Bottom,
            "obliqueTopLeft" => A.PresetCameraValues.ObliqueTopLeft,
            "obliqueTop" => A.PresetCameraValues.ObliqueTop,
            "obliqueTopRight" => A.PresetCameraValues.ObliqueTopRight,
            "obliqueLeft" => A.PresetCameraValues.ObliqueLeft,
            "obliqueRight" => A.PresetCameraValues.ObliqueRight,
            "obliqueBottomLeft" => A.PresetCameraValues.ObliqueBottomLeft,
            "obliqueBottom" => A.PresetCameraValues.ObliqueBottom,
            "obliqueBottomRight" => A.PresetCameraValues.ObliqueBottomRight,
            "perspectiveFront" => A.PresetCameraValues.PerspectiveFront,
            "perspectiveLeft" => A.PresetCameraValues.PerspectiveLeft,
            "perspectiveRight" => A.PresetCameraValues.PerspectiveRight,
            "perspectiveAbove" => A.PresetCameraValues.PerspectiveAbove,
            "perspectiveBelow" => A.PresetCameraValues.PerspectiveBelow,
            "perspectiveAboveLeftFacing" => A.PresetCameraValues.PerspectiveAboveLeftFacing,
            "perspectiveAboveRightFacing" => A.PresetCameraValues.PerspectiveAboveRightFacing,
            "perspectiveContrastingLeftFacing" => A.PresetCameraValues.PerspectiveContrastingLeftFacing,
            "perspectiveContrastingRightFacing" => A.PresetCameraValues.PerspectiveContrastingRightFacing,
            "perspectiveHeroicLeftFacing" => A.PresetCameraValues.PerspectiveHeroicLeftFacing,
            "perspectiveHeroicRightFacing" => A.PresetCameraValues.PerspectiveHeroicRightFacing,
            "perspectiveHeroicExtremeLeftFacing" => A.PresetCameraValues.PerspectiveHeroicExtremeLeftFacing,
            "perspectiveHeroicExtremeRightFacing" => A.PresetCameraValues.PerspectiveHeroicExtremeRightFacing,
            "perspectiveRelaxed" => A.PresetCameraValues.PerspectiveRelaxed,
            "perspectiveRelaxedModerately" => A.PresetCameraValues.PerspectiveRelaxedModerately,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid scene3d camera preset: {node?.ToJsonString() ?? "null"}",
                "camera presets include orthographicFront, isometric*, isometricOffAxis*, oblique*, perspective* (44 total); legacy presets are not accepted.",
                candidates: CameraPresets),
        };
    }

    /// <summary>Maps an a:camera @prst back to its token; null when absent or a legacy preset.</summary>
    private static string? CameraPresetToken(A.PresetCameraValues? val)
    {
        if (val is null)
        {
            return null;
        }

        // Reuse the forward map (the *Values structs have no member-name ToString) to stay single-sourced.
        foreach (var token in CameraPresets)
        {
            if (ParseCameraPreset(JsonValue.Create(token)) == val.Value)
            {
                return token;
            }
        }

        return null;
    }

    /// <summary>The 15 light-rig tokens accepted (a:lightRig @rig), in error-candidate order.</summary>
    private static readonly IReadOnlyList<string> LightRigTokens =
        ["threePt", "twoPt", "balanced", "soft", "harsh", "flood", "contrasting", "morning",
         "sunrise", "sunset", "chilly", "freezing", "flat", "glow", "brightRoom"];

    /// <summary>Maps a light-rig token to a:lightRig @rig; rejects the 12 legacy rigs with candidates.</summary>
    private static A.LightRigValues ParseLightRig(JsonNode? node)
    {
        var raw = node is null ? string.Empty : J.ScalarText(node).Trim();
        return raw switch
        {
            "threePt" => A.LightRigValues.ThreePoints,
            "twoPt" => A.LightRigValues.TwoPoints,
            "balanced" => A.LightRigValues.Balanced,
            "soft" => A.LightRigValues.Soft,
            "harsh" => A.LightRigValues.Harsh,
            "flood" => A.LightRigValues.Flood,
            "contrasting" => A.LightRigValues.Contrasting,
            "morning" => A.LightRigValues.Morning,
            "sunrise" => A.LightRigValues.Sunrise,
            "sunset" => A.LightRigValues.Sunset,
            "chilly" => A.LightRigValues.Chilly,
            "freezing" => A.LightRigValues.Freezing,
            "flat" => A.LightRigValues.Flat,
            "glow" => A.LightRigValues.Glow,
            "brightRoom" => A.LightRigValues.BrightRoom,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid scene3d light rig: {node?.ToJsonString() ?? "null"}",
                "light rigs: threePt, twoPt, balanced, soft, harsh, flood, contrasting, morning, sunrise, sunset, chilly, freezing, flat, glow, brightRoom.",
                candidates: LightRigTokens),
        };
    }

    /// <summary>Maps an a:lightRig @rig back to its token; null when absent or a legacy rig.</summary>
    private static string? LightRigToken(A.LightRigValues? val)
    {
        if (val is null)
        {
            return null;
        }

        return val.Value switch
        {
            _ when val.Value == A.LightRigValues.ThreePoints => "threePt",
            _ when val.Value == A.LightRigValues.TwoPoints => "twoPt",
            _ when val.Value == A.LightRigValues.Balanced => "balanced",
            _ when val.Value == A.LightRigValues.Soft => "soft",
            _ when val.Value == A.LightRigValues.Harsh => "harsh",
            _ when val.Value == A.LightRigValues.Flood => "flood",
            _ when val.Value == A.LightRigValues.Contrasting => "contrasting",
            _ when val.Value == A.LightRigValues.Morning => "morning",
            _ when val.Value == A.LightRigValues.Sunrise => "sunrise",
            _ when val.Value == A.LightRigValues.Sunset => "sunset",
            _ when val.Value == A.LightRigValues.Chilly => "chilly",
            _ when val.Value == A.LightRigValues.Freezing => "freezing",
            _ when val.Value == A.LightRigValues.Flat => "flat",
            _ when val.Value == A.LightRigValues.Glow => "glow",
            _ when val.Value == A.LightRigValues.BrightRoom => "brightRoom",
            _ => null,
        };
    }

    /// <summary>The 8 light-direction tokens accepted (a:lightRig @dir), in error-candidate order.</summary>
    private static readonly IReadOnlyList<string> LightDirectionTokens =
        ["tl", "t", "tr", "l", "r", "bl", "b", "br"];

    /// <summary>Maps a light-direction token to a:lightRig @dir; throws invalid_args with candidates otherwise.</summary>
    private static A.LightRigDirectionValues ParseLightDirection(JsonNode? node)
    {
        var raw = node is null ? string.Empty : J.ScalarText(node).Trim();
        return raw switch
        {
            "tl" => A.LightRigDirectionValues.TopLeft,
            "t" => A.LightRigDirectionValues.Top,
            "tr" => A.LightRigDirectionValues.TopRight,
            "l" => A.LightRigDirectionValues.Left,
            "r" => A.LightRigDirectionValues.Right,
            "bl" => A.LightRigDirectionValues.BottomLeft,
            "b" => A.LightRigDirectionValues.Bottom,
            "br" => A.LightRigDirectionValues.BottomRight,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid light direction: {node?.ToJsonString() ?? "null"}",
                "light directions: tl, t, tr, l, r, bl, b, br.",
                candidates: LightDirectionTokens),
        };
    }

    /// <summary>Maps an a:lightRig @dir back to its token; null when absent or unrecognized.</summary>
    private static string? LightDirectionToken(A.LightRigDirectionValues? val)
    {
        if (val is null)
        {
            return null;
        }

        return val.Value switch
        {
            _ when val.Value == A.LightRigDirectionValues.TopLeft => "tl",
            _ when val.Value == A.LightRigDirectionValues.Top => "t",
            _ when val.Value == A.LightRigDirectionValues.TopRight => "tr",
            _ when val.Value == A.LightRigDirectionValues.Left => "l",
            _ when val.Value == A.LightRigDirectionValues.Right => "r",
            _ when val.Value == A.LightRigDirectionValues.BottomLeft => "bl",
            _ when val.Value == A.LightRigDirectionValues.Bottom => "b",
            _ when val.Value == A.LightRigDirectionValues.BottomRight => "br",
            _ => null,
        };
    }

    // ---- outline ------------------------------------------------------------

    /// <summary>
    /// Sets (or clears, on false) a shape outline (a:ln). A color value is the stroke; a
    /// {color?, width?, dash?, compound?} object additionally tunes width, dash pattern and
    /// compound line type. The bare-string / true / false / null / "" forms are unchanged.
    /// </summary>
    public static void SetOutline(ShapeView view, JsonNode? value)
    {
        var properties = RequireShapeProperties(view);
        foreach (var existing in properties.Elements<A.Outline>().ToList())
        {
            existing.Remove();
        }

        if (IsFalse(value))
        {
            return;
        }

        var outline = value is JsonObject obj
            ? BuildOutline(obj)
            : new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = ColorOrDefault(value, "000000") }))
            {
                Width = 12_700, // 1pt
            };

        // a:ln follows the fill/geometry but precedes a:effectLst / a:scene3d / a:sp3d / a:extLst
        // in CT_ShapeProperties. Anchor before the FIRST of them that is present — otherwise a
        // shape carrying a 3-D element (bevel/scene3d) but no effectLst would get the outline
        // appended AFTER it, producing schema-invalid XML. (Legacy: with only an effectLst or
        // nothing present, this resolves exactly as before, so no valid v1.25 output changes.)
        var anchor = (OpenXmlElement?)properties.GetFirstChild<A.EffectList>()
            ?? (OpenXmlElement?)properties.GetFirstChild<A.Scene3DType>()
            ?? (OpenXmlElement?)properties.GetFirstChild<A.Shape3DType>()
            ?? properties.GetFirstChild<A.ExtensionList>();
        if (anchor is not null)
        {
            properties.InsertBefore(outline, anchor);
        }
        else
        {
            properties.Append(outline);
        }
    }

    /// <summary>The keys the outline object form accepts.</summary>
    private static readonly IReadOnlyList<string> OutlineKeys = ["color", "width", "dash", "compound"];

    /// <summary>
    /// Builds an a:ln from the {color?, width?, dash?, compound?} object form. Child order honors
    /// the schema: a:solidFill precedes a:prstDash; @w and @cmpd are attributes of a:ln.
    /// </summary>
    private static A.Outline BuildOutline(JsonObject obj)
    {
        foreach (var (key, _) in obj)
        {
            if (!OutlineKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown outline prop '{key}'.",
                    "outline object props: color, width, dash, compound.",
                    candidates: OutlineKeys);
            }
        }

        var hex = obj.TryGetPropertyValue("color", out var colorNode)
            ? Units.ParseColorHex("color", colorNode)
            : "000000";
        var outline = new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = hex }))
        {
            Width = obj.TryGetPropertyValue("width", out var widthNode)
                ? (int)Units.ParseLengthEmu("width", widthNode)
                : 12_700, // 1pt
        };

        // a:prstDash follows a:solidFill in the a:ln child order.
        if (obj.TryGetPropertyValue("dash", out var dashNode))
        {
            outline.Append(new A.PresetDash { Val = ParseDashToken(dashNode) });
        }

        if (obj.TryGetPropertyValue("compound", out var compoundNode))
        {
            outline.CompoundLineType = ParseCompoundToken(compoundNode);
        }

        return outline;
    }

    /// <summary>The dash tokens the outline object accepts (a:prstDash @val).</summary>
    private static readonly IReadOnlyList<string> DashTokens =
        ["solid", "dash", "dot", "dashDot", "dashDotDot", "lgDash", "lgDashDot", "lgDashDotDot"];

    /// <summary>The compound tokens the outline object accepts (a:ln @cmpd).</summary>
    private static readonly IReadOnlyList<string> CompoundTokens =
        ["single", "double", "thickThin", "thinThick", "triple"];

    /// <summary>Maps a dash token to its a:prstDash value; throws invalid_args with candidates otherwise.</summary>
    private static A.PresetLineDashValues ParseDashToken(JsonNode? node)
    {
        var raw = node is null ? string.Empty : J.ScalarText(node).Trim();
        return raw switch
        {
            "solid" => A.PresetLineDashValues.Solid,
            "dash" => A.PresetLineDashValues.Dash,
            "dot" => A.PresetLineDashValues.Dot,
            "dashDot" => A.PresetLineDashValues.DashDot,
            "dashDotDot" => A.PresetLineDashValues.SystemDashDotDot,
            "lgDash" => A.PresetLineDashValues.LargeDash,
            "lgDashDot" => A.PresetLineDashValues.LargeDashDot,
            "lgDashDotDot" => A.PresetLineDashValues.LargeDashDotDot,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid outline dash: {node?.ToJsonString() ?? "null"}",
                "outline dash tokens: solid, dash, dot, dashDot, dashDotDot, lgDash, lgDashDot, lgDashDotDot.",
                candidates: DashTokens),
        };
    }

    /// <summary>Maps a compound token to its a:ln @cmpd value; throws invalid_args with candidates otherwise.</summary>
    private static A.CompoundLineValues ParseCompoundToken(JsonNode? node)
    {
        var raw = node is null ? string.Empty : J.ScalarText(node).Trim();
        return raw switch
        {
            "single" => A.CompoundLineValues.Single,
            "double" => A.CompoundLineValues.Double,
            "thickThin" => A.CompoundLineValues.ThickThin,
            "thinThick" => A.CompoundLineValues.ThinThick,
            "triple" => A.CompoundLineValues.Triple,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid outline compound: {node?.ToJsonString() ?? "null"}",
                "outline compound tokens: single, double, thickThin, thinThick, triple.",
                candidates: CompoundTokens),
        };
    }

    // ---- read-back ----------------------------------------------------------

    /// <summary>The effects set on a shape (for the get projection); null when none are present.</summary>
    public static object? Read(OpenXmlCompositeElement element)
    {
        var properties = ShapePropertiesOf(element);
        if (properties is null)
        {
            return null;
        }

        var effectList = properties.GetFirstChild<A.EffectList>();
        var shadow = EffectColor(effectList?.GetFirstChild<A.OuterShadow>());
        var glow = EffectColor(effectList?.GetFirstChild<A.Glow>());
        var hasReflection = effectList?.GetFirstChild<A.Reflection>() is not null;
        var outline = ReadOutline(properties.GetFirstChild<A.Outline>());
        // Trim to 3 decimals so a foreign deck's arbitrary EMU radius projects a clean "N.NNNpt"
        // rather than a float-noise string like "3.1496062992125984pt"; the documented values
        // (2.5pt=31750, 5pt=63500) are exact and unaffected.
        var softEdge = effectList?.GetFirstChild<A.SoftEdge>()?.Radius?.Value is { } rad
            ? Units.Inv($"{rad / EmuPerPoint:0.###}pt")
            : null;
        var innerShadow = ReadInnerShadow(effectList?.GetFirstChild<A.InnerShadow>());
        var bevel = ReadBevel(properties.GetFirstChild<A.Shape3DType>());
        var scene3d = ReadScene3D(properties.GetFirstChild<A.Scene3DType>());

        if (shadow is null && glow is null && !hasReflection && outline is null && softEdge is null
            && innerShadow is null && bevel is null && scene3d is null)
        {
            return null;
        }

        return new
        {
            Shadow = shadow,
            Glow = glow,
            Reflection = hasReflection ? true : (bool?)null,
            Outline = outline,
            SoftEdge = softEdge,
            InnerShadow = innerShadow,
            Bevel = bevel,
            Scene3d = scene3d,
        };
    }

    /// <summary>
    /// Projects an a:innerShdw as a bare hex color STRING when its geometry is the PowerPoint default,
    /// or as {color, blur, dist, dir} when a non-default blur radius, distance or direction is present.
    /// Discriminated by what is set (like ReadOutline), so a bare-color set round-trips to the string
    /// and an object set round-trips to the object.
    /// </summary>
    private static object? ReadInnerShadow(A.InnerShadow? inner)
    {
        if (inner is null)
        {
            return null;
        }

        var color = inner.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value?.ToUpperInvariant() ?? "000000";
        var blur = inner.BlurRadius?.Value;
        var dist = inner.Distance?.Value;
        var dir = inner.Direction?.Value;

        // Default geometry projects as the bare hex string; any non-default value yields the object.
        if ((blur is null || blur == InnerShadowBlur) &&
            (dist is null || dist == InnerShadowDistance) &&
            (dir is null || dir == InnerShadowDirection))
        {
            return color;
        }

        return new
        {
            Color = color,
            Blur = blur is null ? null : Units.Inv($"{blur.Value / EmuPerPoint:0.###}pt"),
            Dist = dist is null ? null : Units.Inv($"{dist.Value / EmuPerPoint:0.###}pt"),
            Dir = dir is null ? (int?)null : (int)Math.Round(dir.Value / DirUnitsPerDegree),
        };
    }

    /// <summary>
    /// Projects an a:sp3d/a:bevelT as a bare preset STRING when the bevel is a plain preset at the
    /// 6pt default (no extrusion depth or color), or as {preset, width?, height?, depth?, depthColor?}
    /// when a non-default width/height, an extrusion depth or an extrusion color is present.
    /// Discriminated by what is set (like ReadOutline / ReadInnerShadow); null when there is no a:sp3d.
    /// </summary>
    private static object? ReadBevel(A.Shape3DType? sp3d)
    {
        var bevelTop = sp3d?.GetFirstChild<A.BevelTop>();
        if (bevelTop is null)
        {
            return null;
        }

        var preset = BevelPresetToken(bevelTop.Preset?.Value) ?? "circle";
        var width = bevelTop.Width?.Value;
        var height = bevelTop.Height?.Value;
        var depth = sp3d!.ExtrusionHeight?.Value;
        var depthColor = sp3d.GetFirstChild<A.ExtrusionColor>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value?.ToUpperInvariant();
        var bevelBottom = ReadBevelSide(sp3d.GetFirstChild<A.BevelBottom>());
        var contour = ReadContour(sp3d);
        var material = MaterialToken(sp3d.PresetMaterial?.Value);
        var z = sp3d.Z?.Value;

        // A plain preset at the 6pt default (no depth/color and none of the v1.26 fields) projects as the
        // bare preset string — so a v1.25 bevel reads back byte-identically and gains no new keys.
        if ((width is null || width == DefaultBevelSize) &&
            (height is null || height == DefaultBevelSize) &&
            depth is null && depthColor is null &&
            bevelBottom is null && contour is null && material is null && z is null)
        {
            return preset;
        }

        return new
        {
            Preset = preset,
            Width = width is null ? null : Units.Inv($"{width.Value / EmuPerPoint:0.###}pt"),
            Height = height is null ? null : Units.Inv($"{height.Value / EmuPerPoint:0.###}pt"),
            Depth = depth is null ? null : Units.Inv($"{depth.Value / EmuPerPoint:0.###}pt"),
            DepthColor = depthColor,
            BevelBottom = bevelBottom,
            Contour = contour,
            Material = material,
            Z = z is null ? null : Units.Inv($"{z.Value / EmuPerPoint:0.###}pt"),
        };
    }

    /// <summary>
    /// Projects an a:bevelB as a bare preset STRING when it is a plain preset at the 6pt default, or as
    /// {preset, width?, height?} when a non-default width/height is present; null when there is no a:bevelB.
    /// </summary>
    private static object? ReadBevelSide(A.BevelBottom? bevel)
    {
        if (bevel is null)
        {
            return null;
        }

        var preset = BevelPresetToken(bevel.Preset?.Value) ?? "circle";
        var width = bevel.Width?.Value;
        var height = bevel.Height?.Value;

        if ((width is null || width == DefaultBevelSize) && (height is null || height == DefaultBevelSize))
        {
            return preset;
        }

        return new
        {
            Preset = preset,
            Width = width is null ? null : Units.Inv($"{width.Value / EmuPerPoint:0.###}pt"),
            Height = height is null ? null : Units.Inv($"{height.Value / EmuPerPoint:0.###}pt"),
        };
    }

    /// <summary>
    /// Projects the contour as {color, width?} when a:contourClr is present; width is omitted when it is the
    /// 1pt default (so a color-only contour round-trips to {color}). Null when there is no a:contourClr.
    /// </summary>
    private static object? ReadContour(A.Shape3DType sp3d)
    {
        var contourColor = sp3d.GetFirstChild<A.ContourColor>();
        if (contourColor is null)
        {
            return null;
        }

        var color = contourColor.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value?.ToUpperInvariant();
        var width = sp3d.ContourWidth?.Value;
        return new
        {
            Color = color,
            Width = (width is null || width == DefaultContourWidth) ? null : Units.Inv($"{width.Value / EmuPerPoint:0.###}pt"),
        };
    }

    /// <summary>Maps an a:bevelT @prst back to its bevel preset token; null when absent or unrecognized.</summary>
    private static string? BevelPresetToken(A.BevelPresetValues? val)
    {
        if (val is null)
        {
            return null;
        }

        return val.Value switch
        {
            _ when val.Value == A.BevelPresetValues.RelaxedInset => "relaxedInset",
            _ when val.Value == A.BevelPresetValues.Circle => "circle",
            _ when val.Value == A.BevelPresetValues.Slope => "slope",
            _ when val.Value == A.BevelPresetValues.Cross => "cross",
            _ when val.Value == A.BevelPresetValues.Angle => "angle",
            _ when val.Value == A.BevelPresetValues.SoftRound => "softRound",
            _ when val.Value == A.BevelPresetValues.Convex => "convex",
            _ when val.Value == A.BevelPresetValues.CoolSlant => "coolSlant",
            _ when val.Value == A.BevelPresetValues.Divot => "divot",
            _ when val.Value == A.BevelPresetValues.Riblet => "riblet",
            _ when val.Value == A.BevelPresetValues.HardEdge => "hardEdge",
            _ when val.Value == A.BevelPresetValues.ArtDeco => "artDeco",
            _ => null,
        };
    }

    /// <summary>
    /// Projects an a:scene3d as a bare camera preset STRING when there is no light rig and no camera rotation,
    /// or as {camera, lightRig?} otherwise; null when there is no a:scene3d (or it has no a:camera).
    /// </summary>
    private static object? ReadScene3D(A.Scene3DType? scene)
    {
        var camera = scene?.GetFirstChild<A.Camera>();
        if (camera is null)
        {
            return null;
        }

        var cameraPreset = CameraPresetToken(camera.Preset?.Value);
        var cameraRotation = ReadRotation(camera.GetFirstChild<A.Rotation>());
        var lightRigElement = scene!.GetFirstChild<A.LightRig>();
        var lightRig = lightRigElement is null || IsDefaultLightRig(lightRigElement)
            ? null
            : ReadLightRig(lightRigElement);

        // Bare camera string when there is no light rig and no camera rotation.
        if (lightRig is null && cameraRotation is null)
        {
            return cameraPreset;
        }

        return new
        {
            Camera = cameraRotation is null
                ? (object?)cameraPreset
                : new { Preset = cameraPreset, Rotation = cameraRotation },
            LightRig = lightRig,
        };
    }

    /// <summary>
    /// Projects an a:lightRig as a bare rig STRING when its direction is 't' (Top) and it has no rotation, or
    /// as {rig, dir, rotation?} otherwise; null when there is no a:lightRig.
    /// </summary>
    private static object? ReadLightRig(A.LightRig? rig)
    {
        if (rig is null)
        {
            return null;
        }

        var rigToken = LightRigToken(rig.Rig?.Value);
        var dir = LightDirectionToken(rig.Direction?.Value);
        var rotation = ReadRotation(rig.GetFirstChild<A.Rotation>());

        // Bare rig string when the direction is the synthesized default 't' (Top) and there is no rotation.
        if (dir == "t" && rotation is null)
        {
            return rigToken;
        }

        return new { Rig = rigToken, Dir = dir, Rotation = rotation };
    }

    /// <summary>Projects an a:rot as {lat, lon, rev} in whole degrees; null when there is no a:rot.</summary>
    private static object? ReadRotation(A.Rotation? rotation)
    {
        if (rotation is null)
        {
            return null;
        }

        return new
        {
            Lat = DegreesFrom(rotation.Latitude?.Value),
            Lon = DegreesFrom(rotation.Longitude?.Value),
            Rev = DegreesFrom(rotation.Revolution?.Value),
        };
    }

    /// <summary>Converts a 60000ths-of-a-degree angle back to whole degrees; null when absent.</summary>
    private static int? DegreesFrom(int? units) =>
        units is null ? null : (int)Math.Round(units.Value / DirUnitsPerDegree);

    private static string? EffectColor(OpenXmlElement? effect) =>
        effect is null ? null : effect.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value?.ToUpperInvariant() ?? DefaultEffectColor;

    /// <summary>
    /// Projects an a:ln as a bare hex color STRING (the legacy shape) when it is a plain 1pt
    /// solid stroke, or as {color, width?, dash?, compound?} when a non-default width, a:prstDash
    /// or @cmpd is present. Discriminated by what is set, not a mode flag — so a legacy outline set
    /// round-trips to the bare string and an object set round-trips to the object.
    /// </summary>
    private static object? ReadOutline(A.Outline? outline)
    {
        if (outline is null)
        {
            return null;
        }

        var color = outline.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();
        var width = outline.Width?.Value;
        var dash = DashToken(outline.GetFirstChild<A.PresetDash>()?.Val?.Value);
        var compound = CompoundToken(outline.CompoundLineType?.Value);

        // The legacy/default stroke (1pt, no dash, no compound) projects as the bare hex string.
        if ((width is null || width == 12_700) && dash is null && compound is null)
        {
            return color;
        }

        return new
        {
            Color = color,
            Width = width is null ? (string?)null : Units.Inv($"{width.Value}emu"),
            Dash = dash,
            Compound = compound,
        };
    }

    /// <summary>Maps an a:prstDash @val back to its outline dash token; null when absent or unrecognized.</summary>
    private static string? DashToken(A.PresetLineDashValues? val)
    {
        if (val is null)
        {
            return null;
        }

        return val.Value switch
        {
            _ when val.Value == A.PresetLineDashValues.Solid => "solid",
            _ when val.Value == A.PresetLineDashValues.Dash => "dash",
            _ when val.Value == A.PresetLineDashValues.Dot => "dot",
            _ when val.Value == A.PresetLineDashValues.DashDot => "dashDot",
            _ when val.Value == A.PresetLineDashValues.SystemDashDotDot => "dashDotDot",
            _ when val.Value == A.PresetLineDashValues.LargeDash => "lgDash",
            _ when val.Value == A.PresetLineDashValues.LargeDashDot => "lgDashDot",
            _ when val.Value == A.PresetLineDashValues.LargeDashDotDot => "lgDashDotDot",
            _ => null,
        };
    }

    /// <summary>Maps an a:ln @cmpd back to its outline compound token; null when absent or unrecognized.</summary>
    private static string? CompoundToken(A.CompoundLineValues? val)
    {
        if (val is null)
        {
            return null;
        }

        return val.Value switch
        {
            _ when val.Value == A.CompoundLineValues.Single => "single",
            _ when val.Value == A.CompoundLineValues.Double => "double",
            _ when val.Value == A.CompoundLineValues.ThickThin => "thickThin",
            _ when val.Value == A.CompoundLineValues.ThinThick => "thinThick",
            _ when val.Value == A.CompoundLineValues.Triple => "triple",
            _ => null,
        };
    }

    // ---- shared plumbing ----------------------------------------------------

    /// <summary>Inserts an effect into the shape's a:effectLst, honoring the schema child order.</summary>
    private static void InsertEffect(ShapeView view, OpenXmlElement effect, string[] after, string[] before)
    {
        var properties = RequireShapeProperties(view);
        var effectList = properties.GetFirstChild<A.EffectList>();
        if (effectList is null)
        {
            effectList = new A.EffectList();
            // a:effectLst precedes the 3-D family (a:scene3d then a:sp3d) in spPr; when either is already
            // present, insert before the FIRST of them, else append (a:effectLst is otherwise the last child).
            var anchor = (OpenXmlElement?)properties.GetFirstChild<A.Scene3DType>()
                ?? properties.GetFirstChild<A.Shape3DType>();
            if (anchor is not null)
            {
                properties.InsertBefore(effectList, anchor);
            }
            else
            {
                properties.Append(effectList);
            }
        }

        // Replace any existing effect of the same element type (idempotent set).
        foreach (var existing in effectList.ChildElements.Where(c => c.LocalName == effect.LocalName).ToList())
        {
            existing.Remove();
        }

        if (before.Length > 0 &&
            effectList.ChildElements.FirstOrDefault(c => before.Contains(c.LocalName, StringComparer.Ordinal)) is { } anchorBefore)
        {
            effectList.InsertBefore(effect, anchorBefore);
            return;
        }

        if (after.Length > 0 &&
            effectList.ChildElements.LastOrDefault(c => after.Contains(c.LocalName, StringComparer.Ordinal)) is { } anchorAfter)
        {
            effectList.InsertAfter(effect, anchorAfter);
            return;
        }

        effectList.Append(effect);
    }

    /// <summary>Removes every effect of the given type from the shape's a:effectLst (and the list when empty).</summary>
    private static void RemoveEffect<TEffect>(ShapeView view)
        where TEffect : OpenXmlElement
    {
        var properties = ShapePropertiesOf(view.Element);
        var effectList = properties?.GetFirstChild<A.EffectList>();
        if (effectList is null)
        {
            return;
        }

        foreach (var effect in effectList.Elements<TEffect>().ToList())
        {
            effect.Remove();
        }

        if (!effectList.HasChildren)
        {
            effectList.Remove();
        }
    }

    private static P.ShapeProperties RequireShapeProperties(ShapeView view) =>
        view.Element switch
        {
            P.Shape s => s.ShapeProperties ?? InsertShapeProperties(s),
            P.Picture p => p.ShapeProperties ??= new P.ShapeProperties(),
            P.ConnectionShape c => c.ShapeProperties ??= new P.ShapeProperties(),
            _ => throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Effects are not supported on a '{view.Kind}'.",
                "Shadow, glow, reflection, outline and soft edge apply to shapes, pictures and lines; " +
                "ungroup grouped content in PowerPoint first."),
        };

    private static P.ShapeProperties? ShapePropertiesOf(OpenXmlCompositeElement element) => element switch
    {
        P.Shape s => s.ShapeProperties,
        P.Picture p => p.ShapeProperties,
        P.ConnectionShape c => c.ShapeProperties,
        _ => null,
    };

    private static P.ShapeProperties InsertShapeProperties(P.Shape shape)
    {
        var properties = new P.ShapeProperties();
        shape.InsertAfter(properties, shape.NonVisualShapeProperties);
        return properties;
    }

    /// <summary>True when the value clears the effect (false/"false"); otherwise yields the color hex (or the default).</summary>
    private static bool IsClear(JsonNode? value, out string hex, string defaultHex)
    {
        if (IsFalse(value))
        {
            hex = defaultHex;
            return true;
        }

        hex = ColorOrDefault(value, defaultHex);
        return false;
    }

    /// <summary>The color hex of the value, or the default when the value is true/null (not a color).</summary>
    private static string ColorOrDefault(JsonNode? value, string defaultHex)
    {
        if (value is JsonValue jv && jv.TryGetValue<bool>(out _))
        {
            return defaultHex; // true means "on, default color"
        }

        if (value is null)
        {
            return defaultHex;
        }

        var text = J.ScalarText(value).Trim();
        if (string.Equals(text, "true", StringComparison.OrdinalIgnoreCase))
        {
            return defaultHex;
        }

        return Units.ParseColorHex("color", value);
    }

    /// <summary>True when the value is the empty string (the "" clear form for soft edge).</summary>
    private static bool IsEmptyString(JsonNode? value) =>
        value is JsonValue jv && jv.TryGetValue<string>(out var text) && text.Trim().Length == 0;

    private static bool IsFalse(JsonNode? value)
    {
        if (value is JsonValue jv)
        {
            if (jv.TryGetValue<bool>(out var flag))
            {
                return !flag;
            }

            if (jv.TryGetValue<string>(out var text) && bool.TryParse(text, out var parsed))
            {
                return !parsed;
            }
        }

        return false;
    }
}
