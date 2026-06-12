using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Reads and writes the supported formatting properties on paragraphs and runs:
/// text, style, bold, italic, underline, color, alignment, fontSize.
/// Anything else is a typed <c>unsupported_feature</c> that names the nearest
/// supported property.
/// </summary>
internal static class WordFormatting
{
    public static readonly IReadOnlyList<string> ParagraphProps =
        ["text", "style", "bold", "italic", "underline", "color", "alignment", "fontSize"];

    public static readonly IReadOnlyList<string> RunProps =
        ["text", "style", "bold", "italic", "underline", "color", "fontSize"];

    // ----------------------------------------------------------------- read

    public static Dictionary<string, object?> ReadParagraphProps(Paragraph p)
    {
        var firstRun = p.ChildElements.OfType<Run>().FirstOrDefault();
        return new Dictionary<string, object?>
        {
            ["text"] = p.InnerText,
            ["style"] = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value,
            ["alignment"] = AlignmentName(p.ParagraphProperties?.Justification),
            ["bold"] = firstRun is null ? null : IsOn(firstRun.RunProperties?.Bold),
            ["italic"] = firstRun is null ? null : IsOn(firstRun.RunProperties?.Italic),
            ["underline"] = firstRun is null ? null : IsUnderlined(firstRun.RunProperties),
            ["fontSize"] = firstRun is null ? null : FontSizePoints(firstRun.RunProperties),
            ["color"] = firstRun?.RunProperties?.Color?.Val?.Value,
            ["runs"] = p.ChildElements.OfType<Run>().Count(),
        };
    }

    public static Dictionary<string, object?> ReadRunProps(Run run) => new()
    {
        ["text"] = run.InnerText,
        ["style"] = run.RunProperties?.RunStyle?.Val?.Value,
        ["bold"] = IsOn(run.RunProperties?.Bold),
        ["italic"] = IsOn(run.RunProperties?.Italic),
        ["underline"] = IsUnderlined(run.RunProperties),
        ["fontSize"] = FontSizePoints(run.RunProperties),
        ["color"] = run.RunProperties?.Color?.Val?.Value,
    };

    /// <summary>w:b style toggles: presence means on unless val says off.</summary>
    public static bool? IsOn(OnOffType? element) =>
        element is null ? null : element.Val?.Value ?? true;

    public static bool? IsUnderlined(RunProperties? rPr) =>
        rPr?.Underline is not { } u ? null : (u.Val?.Value ?? UnderlineValues.Single) != UnderlineValues.None;

    public static double? FontSizePoints(RunProperties? rPr) =>
        rPr?.FontSize?.Val?.Value is { } halfPoints &&
        double.TryParse(halfPoints, NumberStyles.Float, CultureInfo.InvariantCulture, out var hp)
            ? hp / 2.0
            : null;

    public static string? AlignmentName(Justification? justification)
    {
        if (justification?.Val is not { } val)
        {
            return null;
        }

        if (val.Value == JustificationValues.Center)
        {
            return "center";
        }

        if (val.Value == JustificationValues.Right || val.Value == JustificationValues.End)
        {
            return "right";
        }

        if (val.Value == JustificationValues.Both || val.Value == JustificationValues.Distribute)
        {
            return "justify";
        }

        return "left";
    }

    // ---------------------------------------------------------------- write

    /// <summary>Applies one property to a paragraph (text/style/alignment, or every run for run-level props).</summary>
    public static void SetParagraphProp(Paragraph p, string name, string value)
    {
        switch (name)
        {
            case "text":
                ReplaceParagraphText(p, value);
                break;

            case "style":
                EnsurePPr(p).ParagraphStyleId = new ParagraphStyleId { Val = value };
                break;

            case "alignment":
                EnsurePPr(p).Justification = new Justification { Val = ParseAlignment(value) };
                break;

            case "bold" or "italic" or "underline" or "color" or "fontSize":
                foreach (var run in EnsureRun(p))
                {
                    SetRunProp(run, name, value);
                }

                break;

            default:
                throw UnsupportedProp(name, "p", ParagraphProps);
        }
    }

    public static void SetRunProp(Run run, string name, string value)
    {
        switch (name)
        {
            case "text":
                run.RemoveAllChildren<Text>();
                run.RemoveAllChildren<Break>();
                run.RemoveAllChildren<TabChar>();
                run.AppendChild(WordHandler.NewText(value));
                break;

            case "style":
                EnsureRPr(run).RunStyle = new RunStyle { Val = value };
                break;

            case "bold":
                EnsureRPr(run).Bold = new Bold { Val = OnOffValue.FromBoolean(ParseBool(name, value)) };
                break;

            case "italic":
                EnsureRPr(run).Italic = new Italic { Val = OnOffValue.FromBoolean(ParseBool(name, value)) };
                break;

            case "underline":
                EnsureRPr(run).Underline = new Underline
                {
                    Val = ParseBool(name, value) ? UnderlineValues.Single : UnderlineValues.None,
                };
                break;

            case "color":
                EnsureRPr(run).Color = new Color { Val = ParseHexColor(value) };
                break;

            case "fontSize":
                EnsureRPr(run).FontSize = new FontSize { Val = ParseFontSizeHalfPoints(value) };
                break;

            default:
                throw UnsupportedProp(name, "run", RunProps);
        }
    }

    /// <summary>Replaces all inline content with one run, preserving the first run's formatting.</summary>
    public static void ReplaceParagraphText(Paragraph p, string text)
    {
        var keepFormatting = p.ChildElements.OfType<Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
        foreach (var child in p.ChildElements.Where(c => c is not ParagraphProperties).ToList())
        {
            child.Remove();
        }

        var run = new Run();
        if (keepFormatting is not null)
        {
            run.RunProperties = keepFormatting;
        }

        run.AppendChild(WordHandler.NewText(text));
        p.AppendChild(run);
    }

    private static IEnumerable<Run> EnsureRun(Paragraph p)
    {
        if (!p.ChildElements.OfType<Run>().Any())
        {
            p.AppendChild(new Run(WordHandler.NewText(string.Empty)));
        }

        return p.ChildElements.OfType<Run>();
    }

    private static ParagraphProperties EnsurePPr(Paragraph p) =>
        p.ParagraphProperties ??= new ParagraphProperties();

    private static RunProperties EnsureRPr(Run run) =>
        run.RunProperties ??= new RunProperties();

    // --------------------------------------------------------------- parsing

    public static JustificationValues ParseAlignment(string value) => value.ToLowerInvariant() switch
    {
        "left" or "start" => JustificationValues.Left,
        "center" or "centre" => JustificationValues.Center,
        "right" or "end" => JustificationValues.Right,
        "justify" or "both" => JustificationValues.Both,
        _ => throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Unknown alignment '{value}'.",
            "Use one of: left, center, right, justify.",
            candidates: ["left", "center", "right", "justify"]),
    };

    public static bool ParseBool(string name, string value) => value.ToLowerInvariant() switch
    {
        "true" or "1" or "on" or "yes" => true,
        "false" or "0" or "off" or "no" => false,
        _ => throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{name}' expects true or false, got '{value}'.",
            $"Pass {name}=true or {name}=false.",
            candidates: ["true", "false"]),
    };

    public static string ParseHexColor(string value)
    {
        var hex = value.TrimStart('#');
        if (hex.Length == 6 && hex.All(Uri.IsHexDigit))
        {
            return hex.ToUpperInvariant();
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{value}' is not a color; expected 6 hex digits.",
            "Use RRGGBB hex, e.g. color=FF0000 for red.");
    }

    public static string ParseFontSizeHalfPoints(string value)
    {
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var points) && points is > 0 and <= 1638)
        {
            return ((int)Math.Round(points * 2)).ToString(CultureInfo.InvariantCulture);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{value}' is not a font size in points.",
            "Pass a positive number of points, e.g. fontSize=12.");
    }

    // --------------------------------------------------- unsupported helpers

    private static AiofficeException UnsupportedProp(string name, string elementType, IReadOnlyList<string> supported)
    {
        var nearest = Nearest(name, supported);
        return new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"Property '{name}' is not supported on {elementType}.",
            $"Did you mean '{nearest}'? Supported {elementType} properties: {string.Join(", ", supported)}.",
            candidates: supported);
    }

    /// <summary>Nearest supported name by alias table, then edit distance.</summary>
    internal static string Nearest(string name, IReadOnlyList<string> supported)
    {
        var aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["size"] = "fontSize",
            ["font"] = "fontSize",
            ["font-size"] = "fontSize",
            ["align"] = "alignment",
            ["justify"] = "alignment",
            ["justification"] = "alignment",
            ["colour"] = "color",
            ["heading"] = "style",
            ["content"] = "text",
            ["value"] = "text",
        };

        if (aliases.TryGetValue(name, out var alias) && supported.Contains(alias, StringComparer.Ordinal))
        {
            return alias;
        }

        return supported.MinBy(s => Levenshtein(name.ToLowerInvariant(), s.ToLowerInvariant()))!;
    }

    private static int Levenshtein(string a, string b)
    {
        var previous = new int[b.Length + 1];
        var current = new int[b.Length + 1];
        for (var j = 0; j <= b.Length; j++)
        {
            previous[j] = j;
        }

        for (var i = 1; i <= a.Length; i++)
        {
            current[0] = i;
            for (var j = 1; j <= b.Length; j++)
            {
                var cost = a[i - 1] == b[j - 1] ? 0 : 1;
                current[j] = Math.Min(Math.Min(current[j - 1] + 1, previous[j] + 1), previous[j - 1] + cost);
            }

            (previous, current) = (current, previous);
        }

        return previous[b.Length];
    }
}

/// <summary>Builds new, schema-valid wordprocessing elements.</summary>
internal static class WordFactory
{
    /// <summary>A paragraph with one run of text, optionally styled.</summary>
    public static Paragraph Paragraph(string? text, string? styleId = null)
    {
        var p = new Paragraph();
        if (styleId is { Length: > 0 })
        {
            p.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = styleId });
        }

        if (text is { Length: > 0 })
        {
            p.AppendChild(new Run(WordHandler.NewText(text)));
        }

        return p;
    }

    /// <summary>A table row whose cells each hold one paragraph of text.</summary>
    public static TableRow Row(IReadOnlyList<string> cells)
    {
        var row = new TableRow();
        foreach (var cell in cells)
        {
            row.AppendChild(new TableCell(Paragraph(cell)));
        }

        return row;
    }

    /// <summary>An empty rows x columns table with the required tblPr + tblGrid (single borders so Word shows it).</summary>
    public static Table Table(int rows, int columns)
    {
        var table = new Table();
        table.AppendChild(new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));
        var grid = new TableGrid();
        for (var c = 0; c < columns; c++)
        {
            grid.AppendChild(new GridColumn());
        }

        table.AppendChild(grid);
        for (var r = 0; r < rows; r++)
        {
            table.AppendChild(Row([.. Enumerable.Repeat(string.Empty, columns)]));
        }

        return table;
    }

    /// <summary>Minimal styles part: Normal + Heading1..3 so created docs render properly in Word.</summary>
    public static void AddDefaultStylesPart(DocumentFormat.OpenXml.Packaging.MainDocumentPart main)
    {
        var styles = new Styles();
        styles.AppendChild(new Style(new StyleName { Val = "Normal" })
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true,
        });

        foreach (var (level, halfPoints) in new[] { (1, "32"), (2, "28"), (3, "26") })
        {
            styles.AppendChild(HeadingStyle(level, halfPoints));
        }

        var stylesPart = main.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = styles;
    }

    /// <summary>A built-in heading definition (also added on demand when an edit references one).</summary>
    public static Style HeadingStyle(int level, string halfPoints) => new(
        new StyleName { Val = $"heading {level}" },
        new BasedOn { Val = "Normal" },
        new StyleParagraphProperties(new OutlineLevel { Val = level - 1 }),
        new StyleRunProperties(new Bold(), new FontSize { Val = halfPoints }))
    {
        Type = StyleValues.Paragraph,
        StyleId = $"Heading{level}",
    };
}
