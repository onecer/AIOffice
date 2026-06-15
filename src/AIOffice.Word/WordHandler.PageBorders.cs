using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Page borders (v1.4.0): a <c>w:pgBorders</c> on a section, set through
/// <c>{"op":"set","path":"/section[1]","props":{"pageBorder":{…}}}</c>. The
/// border style maps to a w:border value, the color to RRGGBB hex, the width to
/// eighths of a point, and <c>sides</c> selects which of the four edges carry the
/// border. Passing <c>"none"</c> clears it. <c>get /section[1]</c> reports the
/// border under <c>pageBorder</c>.
/// </summary>
public sealed partial class WordHandler
{
    /// <summary>The page-border styles we accept, mapped to their w:border value.</summary>
    private static readonly Dictionary<string, BorderValues> PageBorderStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["single"] = BorderValues.Single,
        ["double"] = BorderValues.Double,
        ["thick"] = BorderValues.Thick,
        ["dashed"] = BorderValues.Dashed,
        ["dotted"] = BorderValues.Dotted,
        ["wave"] = BorderValues.Wave,
    };

    private static readonly string[] PageBorderSides = ["all", "top", "bottom", "left", "right"];

    /// <summary>
    /// Applies (or clears) the section's page border. <paramref name="node"/> is
    /// either the string <c>"none"</c> (clear) or an object
    /// <c>{style, color?, widthPt?, sides?}</c>. The four edge elements live inside
    /// <c>w:pgBorders</c>; only the selected sides are emitted.
    /// </summary>
    private static void ApplyPageBorder(SectionProperties sectPr, JsonNode? node)
    {
        // "none" (or a bare null/empty string) clears any existing page border.
        if (node is null || (node is JsonValue v && v.TryGetValue<string>(out var raw) &&
            (raw.Equals("none", StringComparison.OrdinalIgnoreCase) || raw.Length == 0)))
        {
            sectPr.RemoveAllChildren<PageBorders>();
            return;
        }

        if (node is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "pageBorder must be an object {style,color?,widthPt?,sides?} or the string \"none\".",
                "Example: {\"pageBorder\":{\"style\":\"single\",\"color\":\"38BDF8\",\"widthPt\":1.5,\"sides\":\"all\"}}.");
        }

        var styleRaw = obj["style"] is { } styleNode ? NodeToString(styleNode).Trim() : null;
        if (string.IsNullOrEmpty(styleRaw) || !PageBorderStyles.TryGetValue(styleRaw, out var borderValue))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                styleRaw is null or ""
                    ? "pageBorder needs a style."
                    : $"Unknown pageBorder style '{styleRaw}'.",
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
                    $"pageBorder.widthPt must be a positive number of points, got '{widthRaw}'.",
                    "Pass a width in points, e.g. \"widthPt\":1.5.");
            }

            // OOXML caps page-border width at 31.5pt (sz 0..255).
            sizeEighths = (uint)Math.Clamp(Math.Round(pt * 8), 1, 255);
        }

        var sides = obj["sides"] is { } sidesNode && NodeToString(sidesNode) is { Length: > 0 } sidesRaw
            ? sidesRaw.Trim().ToLowerInvariant()
            : "all";
        if (!PageBorderSides.Contains(sides, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"pageBorder.sides '{sides}' is not valid.",
                "Use all, top, bottom, left or right.",
                candidates: PageBorderSides);
        }

        var pgBorders = new PageBorders { OffsetFrom = PageBorderOffsetValues.Text };

        if (sides is "all" or "top")
        {
            pgBorders.AppendChild(NewPageBorder<TopBorder>(borderValue, color, sizeEighths));
        }

        if (sides is "all" or "left")
        {
            pgBorders.AppendChild(NewPageBorder<LeftBorder>(borderValue, color, sizeEighths));
        }

        if (sides is "all" or "bottom")
        {
            pgBorders.AppendChild(NewPageBorder<BottomBorder>(borderValue, color, sizeEighths));
        }

        if (sides is "all" or "right")
        {
            pgBorders.AppendChild(NewPageBorder<RightBorder>(borderValue, color, sizeEighths));
        }

        sectPr.RemoveAllChildren<PageBorders>();
        InsertSectionChild(sectPr, pgBorders);
    }

    private static T NewPageBorder<T>(BorderValues style, string color, uint sizeEighths)
        where T : BorderType, new() =>
        new()
        {
            Val = style,
            Color = color,
            Size = sizeEighths,
            Space = 24, // Word's default page-border spacing from text (points)
        };

    /// <summary>The get shape of a section's page border, or null when none is set.</summary>
    private static Dictionary<string, object?>? PageBorderShape(SectionProperties? sectPr)
    {
        var pgBorders = sectPr?.GetFirstChild<PageBorders>();
        if (pgBorders is null)
        {
            return null;
        }

        var edges = new (string Side, BorderType? Border)[]
        {
            ("top", pgBorders.TopBorder),
            ("bottom", pgBorders.BottomBorder),
            ("left", pgBorders.LeftBorder),
            ("right", pgBorders.RightBorder),
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

    private static string PageBorderStyleName(BorderValues value)
    {
        foreach (var (name, mapped) in PageBorderStyles)
        {
            if (mapped == value)
            {
                return name;
            }
        }

        return value.ToString() ?? "single";
    }
}
