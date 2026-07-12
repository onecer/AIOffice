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

    /// <summary>The keys the bevel object form accepts.</summary>
    private static readonly IReadOnlyList<string> BevelKeys = ["preset", "width", "height", "depth", "depthColor"];

    /// <summary>
    /// Builds an a:sp3d from the {preset?, width?, height?, depth?, depthColor?} object form. Child
    /// order honors the schema: a:bevelT precedes a:extrusionClr; @extrusionH is an attribute of a:sp3d.
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
                    "bevel object props: preset, width, height, depth, depthColor.",
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
        if (obj.TryGetPropertyValue("depth", out var depthNode))
        {
            sp3d.ExtrusionHeight = Units.ParseLengthEmu("depth", depthNode);
        }

        // a:extrusionClr follows a:bevelT in the a:sp3d child order.
        if (obj.TryGetPropertyValue("depthColor", out var depthColorNode))
        {
            sp3d.Append(new A.ExtrusionColor(new A.RgbColorModelHex { Val = Units.ParseColorHex("depthColor", depthColorNode) }));
        }

        return sp3d;
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

        // a:ln follows the fill / geometry but precedes a:effectLst in spPr.
        var anchor = (OpenXmlElement?)properties.GetFirstChild<A.EffectList>();
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

        if (shadow is null && glow is null && !hasReflection && outline is null && softEdge is null && innerShadow is null && bevel is null)
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

        // A plain preset at the 6pt default (no depth, no color) projects as the bare preset string.
        if ((width is null || width == DefaultBevelSize) &&
            (height is null || height == DefaultBevelSize) &&
            depth is null && depthColor is null)
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
            // a:effectLst precedes a:sp3d (the 3-D family) in spPr; when a bevel has already put an
            // a:sp3d there, insert before it, else append (a:effectLst is otherwise the last child).
            var anchor = properties.GetFirstChild<A.Shape3DType>();
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
