using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>
    /// The v1.1.0 text-effect props an agent can <c>set</c> on a run (or a
    /// run-bearing paragraph): <c>shadow</c>, <c>glow</c>, <c>reflection</c> and
    /// <c>outline</c> (alias <c>textOutline</c>). They emit the Word 2010 text
    /// effects (<c>w14:shadow</c>/<c>w14:glow</c>/<c>w14:reflection</c>/<c>w14:textOutline</c>)
    /// into the run properties; the w14 namespace is declared mc:Ignorable on the
    /// document root so the validator passes and Word renders them. A run inside a
    /// body shape (a textbox's <c>w:txbxContent</c>) is a plain <c>w:r</c> too, so
    /// the same set reaches shape text.
    /// </summary>
    internal static readonly IReadOnlyList<string> TextEffectProps =
        ["shadow", "glow", "reflection", "outline", "textOutline"];

    /// <summary>EMU per point: w14 offsets/blur radii are expressed in EMUs (914400 per inch, 72 pt per inch).</summary>
    private const long EmuPerPoint = 12700L;

    /// <summary>
    /// Detaches every text-effect key from <paramref name="props"/> (mutating it)
    /// into a fresh object the caller applies later, after the plain-formatting
    /// loop has rebuilt the runs. Returns null when no effect prop is present, so
    /// the caller can skip the effect pass entirely.
    /// </summary>
    private static JsonObject? ExtractTextEffectProps(JsonObject props)
    {
        var requested = TextEffectProps.Where(props.ContainsKey).ToList();
        if (requested.Count == 0)
        {
            return null;
        }

        var detached = new JsonObject();
        foreach (var key in requested)
        {
            var node = props[key];
            props.Remove(key);
            detached[key] = node?.DeepClone();
        }

        return detached;
    }

    /// <summary>
    /// Applies each effect in <paramref name="effectProps"/> to every run of
    /// <paramref name="node"/> and returns the canonical effect names applied
    /// (outline absorbs the textOutline alias) so the caller can report them.
    /// </summary>
    private static List<string> ApplyTextEffects(WordprocessingDocument doc, ResolvedNode node, JsonObject effectProps)
    {
        var requested = TextEffectProps.Where(effectProps.ContainsKey).ToList();
        if (requested.Count == 0)
        {
            return [];
        }

        var runs = node.Element switch
        {
            Run run => [run],
            Paragraph paragraph => EnsureParagraphRun(paragraph),
            _ => throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Text effects apply to runs and paragraphs, not '{node.Type}'.",
                "Target a run (e.g. /body/p[1]/run[1]) or a paragraph; effects are set on its run text."),
        };

        var applied = new List<string>();
        foreach (var effect in requested)
        {
            var value = effectProps[effect];
            foreach (var run in runs)
            {
                ApplyOneEffect(EnsureRunProperties(run), effect, value);
            }

            applied.Add(effect == "textOutline" ? "outline" : effect);
        }

        DeclareDocumentW14Ignorable(doc);
        return applied.Distinct().ToList();
    }

    /// <summary>Sets (or replaces) one w14 effect on a run-properties element from its JSON value.</summary>
    private static void ApplyOneEffect(RunProperties rPr, string effect, JsonNode? value)
    {
        switch (effect)
        {
            case "shadow":
                rPr.RemoveAllChildren<W14.Shadow>();
                rPr.AppendChild(BuildShadow(value));
                break;

            case "glow":
                rPr.RemoveAllChildren<W14.Glow>();
                rPr.AppendChild(BuildGlow(value));
                break;

            case "reflection":
                rPr.RemoveAllChildren<W14.Reflection>();
                rPr.AppendChild(BuildReflection(value));
                break;

            case "outline":
            case "textOutline":
                rPr.RemoveAllChildren<W14.TextOutlineEffect>();
                rPr.AppendChild(BuildOutline(value));
                break;

            default:
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"Unknown text effect '{effect}'.",
                    $"Supported effects: {string.Join(", ", TextEffectProps)}.",
                    candidates: TextEffectProps);
        }
    }

    // -------------------------------------------------------------- build w14

    /// <summary>
    /// <c>shadow: true</c> for a soft default drop shadow, or
    /// <c>{color?, blur?, distance?, direction?}</c> with color as RRGGBB hex,
    /// blur/distance in points and direction in degrees.
    /// </summary>
    private static W14.Shadow BuildShadow(JsonNode? value)
    {
        var o = AsEffectObject(value, "shadow");
        var shadow = new W14.Shadow
        {
            BlurRadius = PointsToEmu(EffectNumber(o, "blur") ?? 4),
            DistanceFromText = PointsToEmu(EffectNumber(o, "distance") ?? 3),
            DirectionAngle = DegreesToAngle(EffectNumber(o, "direction") ?? 90),
            // The 2010 schema requires the geometric scale/skew quartet to be present.
            HorizontalScalingFactor = 100000,
            VerticalScalingFactor = 100000,
            HorizontalSkewAngle = 0,
            VerticalSkewAngle = 0,
            Alignment = W14.RectangleAlignmentValues.BottomLeft,
        };
        shadow.RgbColorModelHex = EffectColor(o, "color") ?? new W14.RgbColorModelHex { Val = "808080" };
        return shadow;
    }

    /// <summary><c>glow: {color, radius}</c> — radius in points (default 5), color RRGGBB (default a soft blue).</summary>
    private static W14.Glow BuildGlow(JsonNode? value)
    {
        var o = AsEffectObject(value, "glow");
        var glow = new W14.Glow
        {
            GlowRadius = PointsToEmu(EffectNumber(o, "radius") ?? 5),
        };
        glow.RgbColorModelHex = EffectColor(o, "color") ?? new W14.RgbColorModelHex { Val = "4F81BD" };
        return glow;
    }

    /// <summary><c>reflection: true</c> for a default reflection, or <c>{transparency, size}</c> as 0-100 percentages.</summary>
    private static W14.Reflection BuildReflection(JsonNode? value)
    {
        var o = AsEffectObject(value, "reflection");
        var transparency = PercentToThousandths(EffectNumber(o, "transparency") ?? 50);
        var size = PercentToThousandths(EffectNumber(o, "size") ?? 35);
        return new W14.Reflection
        {
            BlurRadius = 6350L,
            StartingOpacity = Math.Clamp(100000 - transparency, 0, 100000),
            StartPosition = 0,
            EndingOpacity = 0,
            EndPosition = Math.Clamp(size, 0, 100000),
            DistanceFromText = 0,
            DirectionAngle = 5400000,
            FadeDirection = 5400000,
            HorizontalScalingFactor = 100000,
            VerticalScalingFactor = -100000,
            HorizontalSkewAngle = 0,
            VerticalSkewAngle = 0,
            Alignment = W14.RectangleAlignmentValues.BottomLeft,
        };
    }

    /// <summary><c>outline: {color, width}</c> — width in points (default 1), color RRGGBB (default black).</summary>
    private static W14.TextOutlineEffect BuildOutline(JsonNode? value)
    {
        var o = AsEffectObject(value, "outline");
        var color = EffectColor(o, "color") ?? new W14.RgbColorModelHex { Val = "000000" };
        return new W14.TextOutlineEffect(new W14.SolidColorFillProperties(color))
        {
            LineWidth = (int)PointsToEmu(EffectNumber(o, "width") ?? 1),
            CapType = W14.LineCapValues.Round,
            Compound = W14.CompoundLineValues.Simple,
            Alignment = W14.PenAlignmentValues.Center,
        };
    }

    // ------------------------------------------------------------------ read

    /// <summary>
    /// Read-back of the w14 effects on a run's properties, for <c>get</c>: a map
    /// keyed by effect name with the decoded prop values, or null when the run
    /// carries no effects (so the key is omitted from the wire form).
    /// </summary>
    internal static Dictionary<string, object?>? ReadTextEffects(RunProperties? rPr)
    {
        if (rPr is null)
        {
            return null;
        }

        var effects = new Dictionary<string, object?>();

        if (rPr.GetFirstChild<W14.Shadow>() is { } shadow)
        {
            effects["shadow"] = new Dictionary<string, object?>
            {
                ["color"] = ColorHex(shadow.RgbColorModelHex),
                ["blur"] = EmuToPoints(shadow.BlurRadius),
                ["distance"] = EmuToPoints(shadow.DistanceFromText),
                ["direction"] = AngleToDegrees(shadow.DirectionAngle),
            };
        }

        if (rPr.GetFirstChild<W14.Glow>() is { } glow)
        {
            effects["glow"] = new Dictionary<string, object?>
            {
                ["color"] = ColorHex(glow.RgbColorModelHex),
                ["radius"] = EmuToPoints(glow.GlowRadius),
            };
        }

        if (rPr.GetFirstChild<W14.Reflection>() is { } reflection)
        {
            effects["reflection"] = new Dictionary<string, object?>
            {
                ["transparency"] = reflection.StartingOpacity?.Value is { } so ? Math.Round((100000 - so) / 1000.0, 2) : null,
                ["size"] = reflection.EndPosition?.Value is { } ep ? Math.Round(ep / 1000.0, 2) : null,
            };
        }

        if (rPr.GetFirstChild<W14.TextOutlineEffect>() is { } outline)
        {
            effects["outline"] = new Dictionary<string, object?>
            {
                ["color"] = ColorHex(outline.Descendants<W14.RgbColorModelHex>().FirstOrDefault()),
                ["width"] = outline.LineWidth?.Value is { } w ? EmuToPoints((long)w) : null,
            };
        }

        return effects.Count > 0 ? effects : null;
    }

    // --------------------------------------------------------------- helpers

    /// <summary>A bare <c>true</c> means "default effect"; an object carries the tuned props; anything else is invalid.</summary>
    private static JsonObject? AsEffectObject(JsonNode? value, string effect)
    {
        switch (value)
        {
            case JsonObject o:
                return o;
            case JsonValue v when v.TryGetValue<bool>(out var b):
                return b
                    ? null // default-tuned effect
                    : throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"'{effect}: false' does not remove the effect.",
                        $"Pass {effect}:true for a default effect, or an object of its props; to clear it, re-set the run text.");
            case null:
                return null;
            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"'{effect}' must be true or an object of effect props.",
                    $"Example: {{\"{effect}\":{{\"color\":\"FF0000\"}}}} or {{\"{effect}\":true}}.");
        }
    }

    private static double? EffectNumber(JsonObject? o, string key)
    {
        if (o?[key] is not { } node)
        {
            return null;
        }

        var raw = NodeToString(node);
        if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Text-effect '{key}' must be a number, got '{raw}'.",
                $"Pass a number, e.g. {{\"{key}\":3}}.");
        }

        return value;
    }

    private static W14.RgbColorModelHex? EffectColor(JsonObject? o, string key) =>
        o?[key] is { } node && NodeToString(node) is { Length: > 0 } raw
            ? new W14.RgbColorModelHex { Val = WordFormatting.ParseHexColor(raw) }
            : null;

    private static long PointsToEmu(double points) => (long)Math.Round(points * EmuPerPoint);

    private static double EmuToPoints(Int64Value? emu) => emu?.Value is { } v ? Math.Round(v / (double)EmuPerPoint, 2) : 0;

    private static double EmuToPoints(long emu) => Math.Round(emu / (double)EmuPerPoint, 2);

    /// <summary>w14 direction is 60000ths of a degree; Word measures clockwise from "east".</summary>
    private static int DegreesToAngle(double degrees) => (int)Math.Round(((degrees % 360) + 360) % 360 * 60000);

    private static double AngleToDegrees(Int32Value? angle) => angle is { } v ? Math.Round(v.Value / 60000.0, 2) : 0;

    /// <summary>A 0-100 percentage as the schema's 0-100000 thousandths-of-a-percent.</summary>
    private static int PercentToThousandths(double percent) => (int)Math.Round(Math.Clamp(percent, 0, 100) * 1000);

    private static string? ColorHex(W14.RgbColorModelHex? color) => color?.Val?.Value;

    private static RunProperties EnsureRunProperties(Run run) => run.RunProperties ??= new RunProperties();

    /// <summary>The runs of a paragraph, creating one empty run when the paragraph has none (so effects have somewhere to land).</summary>
    private static List<Run> EnsureParagraphRun(Paragraph paragraph)
    {
        var runs = paragraph.ChildElements.OfType<Run>().ToList();
        if (runs.Count == 0)
        {
            var run = new Run(NewText(string.Empty));
            paragraph.AppendChild(run);
            runs.Add(run);
        }

        return runs;
    }
}
