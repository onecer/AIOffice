using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Paragraph-level visual primitives (v1.8.0) an agent can <c>set</c> on a body
/// (or header/footer) paragraph: <c>shading</c> (a w:shd fill) and <c>border</c>
/// (a w:pBdr box). Both are structured props peeled out of the stringly-typed
/// formatting loop so they can take an object value; the scalar companions
/// (<c>spacingBefore/After</c>, <c>indentLeft/Right</c>) ride through
/// <see cref="WordFormatting.SetParagraphProp"/>. Shading reuses the table
/// <see cref="ApplyShading"/> clear-pattern; the border reuses the page-border
/// edge builder (<see cref="NewParagraphBorder{T}"/> mirrors
/// <see cref="NewPageBorder{T}"/>) and the shared <see cref="PageBorderStyles"/>
/// style map.
/// </summary>
public sealed partial class WordHandler
{
    /// <summary>The structured paragraph-visual keys handled outside the stringly loop.</summary>
    private static readonly IReadOnlyList<string> ParagraphVisualProps = ["shading", "border"];

    private static readonly string[] ParagraphBorderSides = ["all", "top", "bottom", "left", "right"];

    /// <summary>
    /// Detaches the structured visual keys (<c>shading</c>, <c>border</c>) from
    /// <paramref name="props"/> (mutating it) so the stringly-typed formatting loop
    /// never sees them. Returns null when neither key is present.
    /// </summary>
    private static JsonObject? ExtractParagraphVisualProps(JsonObject props)
    {
        var requested = ParagraphVisualProps.Where(props.ContainsKey).ToList();
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
    /// Applies the peeled <c>shading</c>/<c>border</c> props to a paragraph and
    /// returns the keys applied (for the op summary).
    /// </summary>
    private static List<string> ApplyParagraphVisuals(Paragraph paragraph, JsonObject props)
    {
        var applied = new List<string>();
        foreach (var (name, value) in props)
        {
            switch (name)
            {
                case "shading":
                    ApplyShading(EnsureParagraphPPr(paragraph), NodeToString(value), set: s => EnsureParagraphPPr(paragraph).Shading = s);
                    break;

                case "border":
                    ApplyParagraphBorder(paragraph, value);
                    break;
            }

            applied.Add(name);
        }

        return applied;
    }

    private static ParagraphProperties EnsureParagraphPPr(Paragraph paragraph) =>
        paragraph.ParagraphProperties ??= new ParagraphProperties();

    /// <summary>
    /// Builds (or clears) the paragraph's <c>w:pBdr</c>. <paramref name="node"/> is
    /// either the string <c>"none"</c> (clear) or an object
    /// <c>{style, color?, widthPt?, sides?}</c> — the same grammar as the page
    /// border. Only the selected sides are emitted; the style map and width units
    /// (eighths of a point) match <see cref="ApplyPageBorder"/>.
    /// </summary>
    private static void ApplyParagraphBorder(Paragraph paragraph, JsonNode? node)
    {
        var pPr = EnsureParagraphPPr(paragraph);

        // "none" (or a bare null/empty string) clears any existing paragraph border.
        if (node is null || (node is JsonValue v && v.TryGetValue<string>(out var raw) &&
            (raw.Equals("none", StringComparison.OrdinalIgnoreCase) || raw.Length == 0)))
        {
            pPr.ParagraphBorders = null;
            return;
        }

        if (node is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "border must be an object {style,color?,widthPt?,sides?} or the string \"none\".",
                "Example: {\"border\":{\"style\":\"single\",\"color\":\"38BDF8\",\"widthPt\":1,\"sides\":\"all\"}}.");
        }

        var styleRaw = obj["style"] is { } styleNode ? NodeToString(styleNode).Trim() : null;
        if (string.IsNullOrEmpty(styleRaw) || !PageBorderStyles.TryGetValue(styleRaw, out var borderValue))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                styleRaw is null or ""
                    ? "border needs a style."
                    : $"Unknown border style '{styleRaw}'.",
                "Use one of: single, double, thick, dashed, dotted, wave.",
                candidates: [.. PageBorderStyles.Keys]);
        }

        var color = obj["color"] is { } colorNode && NodeToString(colorNode) is { Length: > 0 } colorRaw
            ? WordFormatting.ParseHexColor(colorRaw)
            : "auto";

        // w:border sz is in eighths of a point; default 0.5pt (4 eighths) like Word.
        uint sizeEighths = 4;
        if (obj["widthPt"] is { } widthNode)
        {
            var widthRaw = NodeToString(widthNode);
            if (!double.TryParse(widthRaw, NumberStyles.Float, CultureInfo.InvariantCulture, out var pt) || pt <= 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"border.widthPt must be a positive number of points, got '{widthRaw}'.",
                    "Pass a width in points, e.g. \"widthPt\":1.5.");
            }

            // OOXML caps a paragraph-border width at 31.5pt (sz 0..255).
            sizeEighths = (uint)Math.Clamp(Math.Round(pt * 8), 1, 255);
        }

        var sides = obj["sides"] is { } sidesNode && NodeToString(sidesNode) is { Length: > 0 } sidesRaw
            ? sidesRaw.Trim().ToLowerInvariant()
            : "all";
        if (!ParagraphBorderSides.Contains(sides, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"border.sides '{sides}' is not valid.",
                "Use all, top, bottom, left or right.",
                candidates: ParagraphBorderSides);
        }

        var pBdr = new ParagraphBorders();

        if (sides is "all" or "top")
        {
            pBdr.AppendChild(NewParagraphBorder<TopBorder>(borderValue, color, sizeEighths));
        }

        if (sides is "all" or "left")
        {
            pBdr.AppendChild(NewParagraphBorder<LeftBorder>(borderValue, color, sizeEighths));
        }

        if (sides is "all" or "bottom")
        {
            pBdr.AppendChild(NewParagraphBorder<BottomBorder>(borderValue, color, sizeEighths));
        }

        if (sides is "all" or "right")
        {
            pBdr.AppendChild(NewParagraphBorder<RightBorder>(borderValue, color, sizeEighths));
        }

        pPr.ParagraphBorders = pBdr;
    }

    private static T NewParagraphBorder<T>(BorderValues style, string color, uint sizeEighths)
        where T : BorderType, new() =>
        new()
        {
            Val = style,
            Color = color,
            Size = sizeEighths,
            Space = 1, // points between the border and the text (Word's paragraph default)
        };

    /// <summary>The get shape of a paragraph's w:pBdr, or null when none is set.</summary>
    private static Dictionary<string, object?>? ParagraphBorderShape(ParagraphProperties? pPr)
    {
        var pBdr = pPr?.ParagraphBorders;
        if (pBdr is null)
        {
            return null;
        }

        var edges = new (string Side, BorderType? Border)[]
        {
            ("top", pBdr.TopBorder),
            ("bottom", pBdr.BottomBorder),
            ("left", pBdr.LeftBorder),
            ("right", pBdr.RightBorder),
        };

        var present = edges.Where(e => e.Border?.Val is not null).ToList();
        if (present.Count == 0)
        {
            return null;
        }

        var first = present[0].Border!;
        var sides = present.Count == 4 ? "all" : string.Join(",", present.Select(e => e.Side));

        return new Dictionary<string, object?>
        {
            ["style"] = PageBorderStyleName(first.Val!.Value),
            ["color"] = first.Color?.Value,
            ["widthPt"] = first.Size?.Value is { } sz ? Math.Round(sz / 8.0, 2) : (double?)null,
            ["sides"] = sides,
        };
    }
}
