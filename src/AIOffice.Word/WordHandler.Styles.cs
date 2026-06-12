using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>Word built-ins we refuse to remove (Heading1..9 are matched by pattern).</summary>
    private static readonly string[] BuiltinStyleIds =
    [
        "Normal", "Title", "Subtitle", "ListParagraph", "Quote", "IntenseQuote",
        "Caption", "Hyperlink", "CommentText", "CommentReference", "TableNormal",
        "NoList", "DefaultParagraphFont",
    ];

    private static readonly string[] StyleEditableProps =
        ["name", "basedOn", "bold", "italic", "underline", "color", "fontSize",
         "alignment", "spacingBefore", "spacingAfter"];

    private static bool IsBuiltinStyleId(string id) =>
        BuiltinStyleIds.Contains(id, StringComparer.OrdinalIgnoreCase) || HeadingLevel(id) is not null;

    // ------------------------------------------------------------------ read

    private static object StylesView(WordprocessingDocument doc)
    {
        var inUse = CollectUsedStyleIds(doc);
        var styles = (doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>() ?? [])
            .Where(s => s.Type is null || s.Type.Value == StyleValues.Paragraph || s.Type.Value == StyleValues.Character)
            .Select(s => StyleShape(s, inUse))
            .ToList();

        return new { view = "styles", count = styles.Count, styles };
    }

    private static object StyleShape(Style style, HashSet<string> inUse)
    {
        var id = style.StyleId?.Value ?? string.Empty;
        return new
        {
            id,
            name = style.StyleName?.Val?.Value ?? id,
            kind = style.Type is { } t && t.Value == StyleValues.Character ? "character" : "paragraph",
            basedOn = style.BasedOn?.Val?.Value,
            builtin = IsBuiltinStyleId(id),
            inUse = inUse.Contains(id),
        };
    }

    /// <summary>Style ids referenced by any paragraph or run in body, headers and footers.</summary>
    private static HashSet<string> CollectUsedStyleIds(WordprocessingDocument doc)
    {
        var used = new HashSet<string>(StringComparer.Ordinal);
        var roots = new List<OpenXmlElement>();
        if (doc.MainDocumentPart?.Document?.Body is { } body)
        {
            roots.Add(body);
        }

        roots.AddRange(WordAddress.HeaderFooterRoots(doc).Select(r => r.Element));

        foreach (var root in roots)
        {
            foreach (var styleId in root.Descendants<ParagraphStyleId>())
            {
                if (styleId.Val?.Value is { Length: > 0 } id)
                {
                    used.Add(id);
                }
            }

            foreach (var runStyle in root.Descendants<RunStyle>())
            {
                if (runStyle.Val?.Value is { Length: > 0 } id)
                {
                    used.Add(id);
                }
            }
        }

        return used;
    }

    /// <summary>get /style[@id=X] data.</summary>
    private static Dictionary<string, object?> GetStyleProperties(WordprocessingDocument doc, DocPath path)
    {
        var style = ResolveStyle(doc, path);
        var inUse = CollectUsedStyleIds(doc);
        var rPr = style.StyleRunProperties;
        var pPr = style.StyleParagraphProperties;
        var spacing = pPr?.SpacingBetweenLines;

        return new Dictionary<string, object?>
        {
            ["id"] = style.StyleId?.Value,
            ["name"] = style.StyleName?.Val?.Value ?? style.StyleId?.Value,
            ["kind"] = style.Type is { } t && t.Value == StyleValues.Character ? "character" : "paragraph",
            ["basedOn"] = style.BasedOn?.Val?.Value,
            ["builtin"] = IsBuiltinStyleId(style.StyleId?.Value ?? string.Empty),
            ["inUse"] = inUse.Contains(style.StyleId?.Value ?? string.Empty),
            ["bold"] = WordFormatting.IsOn(rPr?.Bold),
            ["italic"] = WordFormatting.IsOn(rPr?.Italic),
            ["underline"] = rPr?.Underline is { } u ? (u.Val?.Value ?? UnderlineValues.Single) != UnderlineValues.None : null,
            ["color"] = rPr?.Color?.Val?.Value,
            ["fontSize"] = rPr?.FontSize?.Val?.Value is { } hp &&
                double.TryParse(hp, NumberStyles.Float, CultureInfo.InvariantCulture, out var halfPoints)
                    ? halfPoints / 2.0
                    : null,
            ["alignment"] = WordFormatting.AlignmentName(pPr?.Justification),
            ["spacingBefore"] = TwentiethsToPoints(spacing?.Before?.Value),
            ["spacingAfter"] = TwentiethsToPoints(spacing?.After?.Value),
        };
    }

    private static double? TwentiethsToPoints(string? twentieths) =>
        twentieths is not null &&
        double.TryParse(twentieths, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)
            ? v / 20.0
            : null;

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/styles","type":"style","props":{"id":…,"kind":…,…}}</c>:
    /// defines a new paragraph or character style in the styles part.
    /// </summary>
    private static object ApplyAddStyle(WordprocessingDocument doc, EditOp op)
    {
        if (op.Path is not "/styles")
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add --type style targets /styles, not '{op.Path}'.",
                "Use {\"op\":\"add\",\"path\":\"/styles\",\"type\":\"style\",\"props\":{\"id\":\"Callout\",…}}.",
                candidates: ["/styles"]);
        }

        var props = op.Props?.DeepClone().AsObject() ?? [];
        var id = props["id"] is { } idNode ? NodeToString(idNode) : null;
        props.Remove("id");
        if (id is null || !Regex.IsMatch(id, "^[A-Za-z][A-Za-z0-9_-]*$"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A style needs props.id (letters/digits/_/-, starting with a letter); got '{id}'.",
                "Example: {\"op\":\"add\",\"path\":\"/styles\",\"type\":\"style\",\"props\":{\"id\":\"Callout\",\"kind\":\"paragraph\"}}.");
        }

        var kind = props["kind"] is { } kindNode ? NodeToString(kindNode) : "paragraph";
        props.Remove("kind");
        if (kind is not ("paragraph" or "character"))
        {
            throw new AiofficeException(
                kind is "table" or "numbering" ? ErrorCodes.UnsupportedFeature : ErrorCodes.InvalidArgs,
                $"Style kind '{kind}' is not supported (M2 covers paragraph and character styles).",
                "Use kind paragraph (default) or character.",
                candidates: ["paragraph", "character"]);
        }

        var styles = EnsureStylesRoot(doc);
        if (FindStyle(styles, id) is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Style '{id}' already exists.",
                "Modify it instead: {\"op\":\"set\",\"path\":\"/style[@id=" + id + "]\",\"props\":{…}}.");
        }

        var style = new Style
        {
            Type = kind == "character" ? StyleValues.Character : StyleValues.Paragraph,
            StyleId = id,
            CustomStyle = true,
        };
        style.AppendChild(new StyleName { Val = id });
        styles.AppendChild(style);

        ApplyStyleProps(doc, style, props, kind == "paragraph");
        return new { op = "add", type = "style", path = StylePath(id), id, kind };
    }

    // ------------------------------------------------------------ set/remove

    /// <summary>set /style[@id=X]: modifies an existing style definition.</summary>
    private static object ApplySetStyle(WordprocessingDocument doc, EditOp op)
    {
        var style = ResolveStyle(doc, DocPath.Parse(op.Path));
        var props = RequireProps(op).DeepClone().AsObject();
        var id = style.StyleId?.Value ?? string.Empty;

        foreach (var fixedKey in (string[])["id", "kind"])
        {
            if (props.ContainsKey(fixedKey))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"A style's {fixedKey} cannot change after creation.",
                    "Add a new style with the wanted " + fixedKey + " and re-style the content, then remove the old one.");
            }
        }

        var isParagraph = style.Type is not { } t || t.Value != StyleValues.Character;
        ApplyStyleProps(doc, style, props, isParagraph);
        return new { op = "set", path = StylePath(id), type = "style" };
    }

    /// <summary>remove /style[@id=X]: only non-builtin styles can go.</summary>
    private static object ApplyRemoveStyle(WordprocessingDocument doc, EditOp op)
    {
        var style = ResolveStyle(doc, DocPath.Parse(op.Path));
        var id = style.StyleId?.Value ?? string.Empty;

        if (IsBuiltinStyleId(id))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{id}' is a built-in style and cannot be removed.",
                "Modify it instead with {\"op\":\"set\",\"path\":\"" + StylePath(id) + "\",\"props\":{…}}, or remove a custom style.");
        }

        var wasInUse = CollectUsedStyleIds(doc).Contains(id);
        style.Remove();
        return new
        {
            op = "remove",
            path = StylePath(id),
            type = "style",
            note = wasInUse ? "the style was in use; affected content falls back to Normal" : null,
        };
    }

    // ----------------------------------------------------------------- props

    /// <summary>Applies the supported style props in schema order (name/basedOn, then pPr, then rPr).</summary>
    private static void ApplyStyleProps(WordprocessingDocument doc, Style style, JsonObject props, bool isParagraph)
    {
        foreach (var (key, node) in props.ToList())
        {
            var value = NodeToString(node);
            switch (key)
            {
                case "name":
                    style.StyleName = new StyleName { Val = value };
                    break;

                case "basedOn":
                    var styles = EnsureStylesRoot(doc);
                    if (FindStyle(styles, value) is null && HeadingLevel(value) is null && value != "Normal")
                    {
                        throw new AiofficeException(
                            ErrorCodes.InvalidArgs,
                            $"basedOn style '{value}' does not exist.",
                            "Base the style on an existing one; run 'aioffice read <file> --view styles' to list them.",
                            candidates: [.. styles.Elements<Style>().Select(s => s.StyleId?.Value ?? string.Empty).Where(s => s.Length > 0).Take(8)]);
                    }

                    style.BasedOn = new BasedOn { Val = value };
                    break;

                case "bold":
                    EnsureStyleRunProperties(style).Bold = new Bold { Val = OnOffValue.FromBoolean(WordFormatting.ParseBool(key, value)) };
                    break;

                case "italic":
                    EnsureStyleRunProperties(style).Italic = new Italic { Val = OnOffValue.FromBoolean(WordFormatting.ParseBool(key, value)) };
                    break;

                case "underline":
                    EnsureStyleRunProperties(style).Underline = new Underline
                    {
                        Val = WordFormatting.ParseBool(key, value) ? UnderlineValues.Single : UnderlineValues.None,
                    };
                    break;

                case "color":
                    EnsureStyleRunProperties(style).Color = new Color { Val = WordFormatting.ParseHexColor(value) };
                    break;

                case "fontSize":
                    EnsureStyleRunProperties(style).FontSize = new FontSize { Val = WordFormatting.ParseFontSizeHalfPoints(value) };
                    break;

                case "alignment" when isParagraph:
                    EnsureStyleParagraphProperties(style).Justification = new Justification { Val = WordFormatting.ParseAlignment(value) };
                    break;

                case "spacingBefore" when isParagraph:
                    EnsureSpacing(style).Before = PointsToTwentieths(key, value);
                    break;

                case "spacingAfter" when isParagraph:
                    EnsureSpacing(style).After = PointsToTwentieths(key, value);
                    break;

                case "alignment" or "spacingBefore" or "spacingAfter":
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"'{key}' applies to paragraph styles only; this is a character style.",
                        "Drop the paragraph-level props, or define the style with kind paragraph.");

                default:
                    throw new AiofficeException(
                        ErrorCodes.UnsupportedFeature,
                        $"Style property '{key}' is not supported.",
                        $"Did you mean '{WordFormatting.Nearest(key, StyleEditableProps)}'? Supported: {string.Join(", ", StyleEditableProps)}.",
                        candidates: StyleEditableProps);
            }
        }
    }

    private static string PointsToTwentieths(string key, string value)
    {
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var points) && points is >= 0 and <= 1584)
        {
            return ((int)Math.Round(points * 20)).ToString(CultureInfo.InvariantCulture);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{key}' must be spacing in points (0–1584), got '{value}'.",
            $"Pass points, e.g. {key}=6 for 6pt.");
    }

    private static StyleRunProperties EnsureStyleRunProperties(Style style) =>
        style.StyleRunProperties ??= new StyleRunProperties();

    private static StyleParagraphProperties EnsureStyleParagraphProperties(Style style) =>
        style.StyleParagraphProperties ??= new StyleParagraphProperties();

    private static SpacingBetweenLines EnsureSpacing(Style style)
    {
        var pPr = EnsureStyleParagraphProperties(style);
        return pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
    }

    // ------------------------------------------------------------ resolution

    private static Styles EnsureStylesRoot(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The document has no main part.",
            "Re-export the file from Word.");

        if (main.StyleDefinitionsPart is null)
        {
            WordFactory.AddDefaultStylesPart(main);
        }

        return main.StyleDefinitionsPart!.Styles ??= new Styles();
    }

    private static Style? FindStyle(Styles styles, string id) =>
        styles.Elements<Style>().FirstOrDefault(s => string.Equals(s.StyleId?.Value, id, StringComparison.Ordinal));

    /// <summary>Resolves /style[@id=X] (or positional /style[i]) or throws invalid_path with candidates.</summary>
    private static Style ResolveStyle(WordprocessingDocument doc, DocPath path)
    {
        var styles = EnsureStylesRoot(doc);
        var all = styles.Elements<Style>().ToList();
        var segment = path.Segments[0];

        if (path.Segments.Count == 1 && segment.Id is { } id)
        {
            return FindStyle(styles, id) ?? throw StyleNotFound($"No style has id '{id}'.", all);
        }

        if (path.Segments.Count == 1 && segment.Index is { } index)
        {
            if (index <= all.Count)
            {
                return all[index - 1];
            }

            throw StyleNotFound($"/style[{index}] does not exist; there are {all.Count} style(s).", all);
        }

        throw StyleNotFound($"'{path.ToCanonicalString()}' is not a style path; use /style[@id=X].", all);
    }

    private static AiofficeException StyleNotFound(string message, List<Style> styles) => new(
        ErrorCodes.InvalidPath,
        message,
        "Run 'aioffice read <file> --view styles' to list style ids.",
        candidates: [.. styles.Select(s => s.StyleId?.Value).OfType<string>().Take(8).Select(StylePath)]);

    private static string StylePath(string id) => $"/style[@id={id}]";
}
