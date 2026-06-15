using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Shape text/visual effects: shadow (a:outerShdw), glow (a:glow), reflection
/// (a:reflection) — all children of the shape's a:effectLst — and outline
/// (a:ln) on the shape's spPr. Each setter accepts a color hex (the effect's
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

    // ---- outline ------------------------------------------------------------

    /// <summary>Sets (or clears, on false) a shape outline (a:ln); a color value is the stroke.</summary>
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

        var hex = ColorOrDefault(value, "000000");
        var outline = new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = hex }))
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
        var outline = properties.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()?
            .RgbColorModelHex?.Val?.Value?.ToUpperInvariant();

        if (shadow is null && glow is null && !hasReflection && outline is null)
        {
            return null;
        }

        return new
        {
            Shadow = shadow,
            Glow = glow,
            Reflection = hasReflection ? true : (bool?)null,
            Outline = outline,
        };
    }

    private static string? EffectColor(OpenXmlElement? effect) =>
        effect is null ? null : effect.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value?.ToUpperInvariant() ?? DefaultEffectColor;

    // ---- shared plumbing ----------------------------------------------------

    /// <summary>Inserts an effect into the shape's a:effectLst, honoring the schema child order.</summary>
    private static void InsertEffect(ShapeView view, OpenXmlElement effect, string[] after, string[] before)
    {
        var properties = RequireShapeProperties(view);
        var effectList = properties.GetFirstChild<A.EffectList>();
        if (effectList is null)
        {
            effectList = new A.EffectList();
            // a:effectLst is the last child of spPr (after fill, ln, …).
            properties.Append(effectList);
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
                "Shadow, glow, reflection and outline apply to shapes, pictures and lines; " +
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
