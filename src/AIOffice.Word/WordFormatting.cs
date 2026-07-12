using System.Globalization;
using System.Text.Json.Nodes;
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
        ["text", "style", "bold", "italic", "underline", "color", "font", "alignment", "fontSize", "rtl",
         "shading", "border", "spacingBefore", "spacingAfter", "indentLeft", "indentRight",
         // 1.10 typography: paragraph-level toggles + line spacing, outline level, tab stops.
         "lineSpacing", "keepNext", "keepLines", "pageBreakBefore", "widowControl", "outlineLevel", "tabStops",
         // 1.10 run props that fan out to every run when set on a paragraph (like bold/font).
         "highlight", "strike", "doubleStrike", "smallCaps", "allCaps", "superscript", "subscript", "characterSpacing"];

    public static readonly IReadOnlyList<string> RunProps =
        ["text", "style", "bold", "italic", "underline", "color", "font", "fontSize", "rtl",
         // 1.10 character typography primitives.
         "highlight", "strike", "doubleStrike", "smallCaps", "allCaps", "superscript", "subscript", "characterSpacing"];

    /// <summary>1.10 run props that fan out to every run when set on a paragraph (alongside the v1.8 set).</summary>
    public static readonly string[] RunFanoutProps =
        ["bold", "italic", "underline", "color", "fontSize", "font",
         "highlight", "strike", "doubleStrike", "smallCaps", "allCaps", "superscript", "subscript", "characterSpacing"];

    /// <summary>The fixed set of named highlight colors Word's w:highlight @val accepts (no hex form).</summary>
    public static readonly IReadOnlyList<string> HighlightColors =
        ["yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan", "darkGreen",
         "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "black", "white", "none"];

    /// <summary>
    /// The underline styles the 'underline' prop accepts as a string (alongside the bool form).
    /// bool true maps to "single", bool false / "none" to "none"; every other entry is a named
    /// w:u @val. Runtime-validated candidate list for the value-shape relaxation.
    /// </summary>
    public static readonly IReadOnlyList<string> UnderlineStyles =
        ["double", "thick", "dotted", "dash", "dashLong", "dotDash", "dotDotDash",
         "wave", "wavyHeavy", "wavyDouble", "words", "single", "none"];

    // ----------------------------------------------------------------- read

    public static Dictionary<string, object?> ReadParagraphProps(Paragraph p)
    {
        var firstRun = p.ChildElements.OfType<Run>().FirstOrDefault();
        var firstRunPr = firstRun?.RunProperties;
        var pPr = p.ParagraphProperties;
        var spacing = pPr?.SpacingBetweenLines;
        var indentation = pPr?.Indentation;
        var props = new Dictionary<string, object?>
        {
            ["text"] = p.InnerText,
            ["style"] = pPr?.ParagraphStyleId?.Val?.Value,
            ["alignment"] = AlignmentName(pPr?.Justification),
            ["bold"] = firstRun is null ? null : IsOn(firstRunPr?.Bold),
            ["italic"] = firstRun is null ? null : IsOn(firstRunPr?.Italic),
            ["underline"] = firstRun is null ? null : UnderlineValue(firstRunPr),
            ["fontSize"] = firstRun is null ? null : FontSizePoints(firstRunPr),
            ["color"] = firstRunPr?.Color?.Val?.Value,
            ["font"] = firstRunPr?.RunFonts?.Ascii?.Value,
            ["rtl"] = IsParagraphRtl(p),
            // 1.8 paragraph-level visuals: shading fill (w:shd), spacing (pt) and
            // indentation (cm). The border box (w:pBdr) reports as a structured object.
            ["shading"] = pPr?.Shading?.Fill?.Value,
            ["spacingBefore"] = TwentiethsToPointsValue(spacing?.Before?.Value),
            ["spacingAfter"] = TwentiethsToPointsValue(spacing?.After?.Value),
            ["indentLeft"] = TwipsToCentimeters(indentation?.Left?.Value),
            ["indentRight"] = TwipsToCentimeters(indentation?.Right?.Value),
            // 1.10 character typography echoed from the first run (like bold/font).
            ["highlight"] = firstRun is null ? null : HighlightName(firstRunPr),
            ["strike"] = firstRun is null ? null : IsOn(firstRunPr?.Strike),
            ["doubleStrike"] = firstRun is null ? null : IsOn(firstRunPr?.DoubleStrike),
            ["smallCaps"] = firstRun is null ? null : IsOn(firstRunPr?.SmallCaps),
            ["allCaps"] = firstRun is null ? null : IsOn(firstRunPr?.Caps),
            ["superscript"] = firstRun is null ? null : VertAlignName(firstRunPr) == "superscript",
            ["subscript"] = firstRun is null ? null : VertAlignName(firstRunPr) == "subscript",
            ["characterSpacing"] = firstRun is null ? null : CharacterSpacingPoints(firstRunPr),
            // 1.10 paragraph typography: line spacing (multiple or {atLeast|exactly}),
            // keep/break toggles, outline level, tab stops.
            ["lineSpacing"] = LineSpacingValue(spacing),
            ["keepNext"] = IsOn(pPr?.KeepNext),
            ["keepLines"] = IsOn(pPr?.KeepLines),
            ["pageBreakBefore"] = IsOn(pPr?.PageBreakBefore),
            ["widowControl"] = IsOn(pPr?.WidowControl),
            ["outlineLevel"] = pPr?.OutlineLevel?.Val?.Value is { } lvl ? (int?)lvl : null,
            ["tabStops"] = TabStopsValue(pPr?.Tabs),
            ["runs"] = p.ChildElements.OfType<Run>().Count(),
        };
        return props;
    }

    /// <summary>w:spacing before/after are twentieths of a point; null stays null.</summary>
    internal static double? TwentiethsToPointsValue(string? twentieths) =>
        twentieths is not null &&
        double.TryParse(twentieths, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)
            ? v / 20.0
            : null;

    /// <summary>w:ind left/right are twips; report them in centimeters (2dp), null stays null.</summary>
    internal static double? TwipsToCentimeters(string? twips) =>
        twips is not null &&
        double.TryParse(twips, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)
            ? Math.Round(v / (1440.0 / 2.54), 2)
            : null;

    /// <summary>
    /// w:spacing line/lineRule -> the get shape of <c>lineSpacing</c>: a bare multiple
    /// for @lineRule="auto" (@line/240), or {atLeast|exactly: points} (@line/20) for the
    /// fixed rules. null when no line rule is set (only before/after spacing present).
    /// </summary>
    internal static object? LineSpacingValue(SpacingBetweenLines? spacing)
    {
        if (spacing?.Line?.Value is not { } lineStr ||
            !double.TryParse(lineStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var line))
        {
            return null;
        }

        var rule = spacing.LineRule?.Value ?? LineSpacingRuleValues.Auto;
        if (rule == LineSpacingRuleValues.AtLeast)
        {
            return new Dictionary<string, object?> { ["atLeast"] = Math.Round(line / 20.0, 2) };
        }

        if (rule == LineSpacingRuleValues.Exact)
        {
            return new Dictionary<string, object?> { ["exactly"] = Math.Round(line / 20.0, 2) };
        }

        // auto: a line-height multiple of 240ths.
        return Math.Round(line / 240.0, 4);
    }

    /// <summary>w:tabs -> the get shape of <c>tabStops</c>: an array of {pos(cm,2dp), align, leader}. null when none.</summary>
    internal static object? TabStopsValue(Tabs? tabs)
    {
        var stops = tabs?.Elements<TabStop>().Where(t => t.Val?.Value != TabStopValues.Clear).ToList();
        if (stops is null || stops.Count == 0)
        {
            return null;
        }

        return stops.Select(t => new Dictionary<string, object?>
        {
            ["pos"] = t.Position?.Value is { } pos ? Math.Round(pos / 567.0, 2) : 0.0,
            ["align"] = TabAlignName(t.Val?.Value),
            ["leader"] = TabLeaderName(t.Leader?.Value),
        }).ToList();
    }

    private static string TabAlignName(TabStopValues? val)
    {
        if (val == TabStopValues.Center)
        {
            return "center";
        }

        if (val == TabStopValues.Right || val == TabStopValues.End)
        {
            return "right";
        }

        if (val == TabStopValues.Decimal)
        {
            return "decimal";
        }

        if (val == TabStopValues.Bar)
        {
            return "bar";
        }

        return "left";
    }

    private static string TabLeaderName(TabStopLeaderCharValues? leader)
    {
        if (leader == TabStopLeaderCharValues.Dot)
        {
            return "dot";
        }

        if (leader == TabStopLeaderCharValues.Hyphen)
        {
            return "hyphen";
        }

        if (leader == TabStopLeaderCharValues.Underscore)
        {
            return "underscore";
        }

        return "none";
    }

    public static Dictionary<string, object?> ReadRunProps(Run run)
    {
        var rPr = run.RunProperties;
        return new Dictionary<string, object?>
        {
            ["text"] = run.InnerText,
            ["style"] = rPr?.RunStyle?.Val?.Value,
            ["bold"] = IsOn(rPr?.Bold),
            ["italic"] = IsOn(rPr?.Italic),
            ["underline"] = UnderlineValue(rPr),
            ["fontSize"] = FontSizePoints(rPr),
            ["color"] = rPr?.Color?.Val?.Value,
            ["font"] = rPr?.RunFonts?.Ascii?.Value,
            ["rtl"] = IsOn(rPr?.RightToLeftText) ?? false,
            // 1.10 character typography primitives.
            ["highlight"] = HighlightName(rPr),
            ["strike"] = IsOn(rPr?.Strike),
            ["doubleStrike"] = IsOn(rPr?.DoubleStrike),
            ["smallCaps"] = IsOn(rPr?.SmallCaps),
            ["allCaps"] = IsOn(rPr?.Caps),
            ["superscript"] = VertAlignName(rPr) == "superscript",
            ["subscript"] = VertAlignName(rPr) == "subscript",
            ["characterSpacing"] = CharacterSpacingPoints(rPr),
        };
    }

    /// <summary>w:highlight @val (a named color string), or null when absent.</summary>
    public static string? HighlightName(RunProperties? rPr) =>
        rPr?.Highlight?.Val is { } v ? v.ToString() : null;

    /// <summary>w:vertAlign @val -> "superscript"/"subscript"/null (baseline or absent).</summary>
    public static string? VertAlignName(RunProperties? rPr)
    {
        if (rPr?.VerticalTextAlignment?.Val?.Value is not { } v)
        {
            return null;
        }

        if (v == VerticalPositionValues.Superscript)
        {
            return "superscript";
        }

        if (v == VerticalPositionValues.Subscript)
        {
            return "subscript";
        }

        return null;
    }

    /// <summary>w:spacing @val (twentieths of a point) -> points; null when absent.</summary>
    public static double? CharacterSpacingPoints(RunProperties? rPr) =>
        rPr?.Spacing?.Val?.Value is { } twentieths ? twentieths / 20.0 : null;

    /// <summary>w:b style toggles: presence means on unless val says off.</summary>
    public static bool? IsOn(OnOffType? element) =>
        element is null ? null : element.Val?.Value ?? true;

    public static bool? IsUnderlined(RunProperties? rPr) =>
        rPr?.Underline is not { } u ? null : (u.Val?.Value ?? UnderlineValues.Single) != UnderlineValues.None;

    /// <summary>
    /// The discriminated read of w:u, following the shipped ReadOutline precedent (bare form -> a
    /// primitive, enriched -> a richer shape): absent -> null, none -> bool false, single/default ->
    /// bool true, any other style -> the style STRING. single/none/absent are byte-stable with the
    /// legacy <see cref="IsUnderlined"/> bool; only non-single content the tool didn't author widens.
    /// </summary>
    public static object? UnderlineValue(RunProperties? rPr) => UnderlineValue(rPr?.Underline);

    /// <summary>The discriminated read off a w:u element (shared by run and style-def rPr).</summary>
    public static object? UnderlineValue(Underline? underline)
    {
        if (underline is null)
        {
            return null;
        }

        var valEnum = underline.Val;
        var val = valEnum?.Value ?? UnderlineValues.Single;
        if (val == UnderlineValues.None)
        {
            return false;
        }

        // single/default -> bool true (byte-stable); any other style -> its @val STRING (e.g. "double").
        return val == UnderlineValues.Single ? true : valEnum!.ToString();
    }

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

            case "rtl":
                SetParagraphRtl(p, ParseBool(name, value));
                break;

            case "spacingBefore":
                EnsureSpacing(p).Before = ParseSpacingTwentieths(name, value);
                break;

            case "spacingAfter":
                EnsureSpacing(p).After = ParseSpacingTwentieths(name, value);
                break;

            case "indentLeft":
                EnsureIndentation(p).Left = ParseIndentTwips(name, value);
                break;

            case "indentRight":
                EnsureIndentation(p).Right = ParseIndentTwips(name, value);
                break;

            // 1.10 paragraph typography.
            case "lineSpacing":
                ApplyLineSpacing(p, value);
                break;

            case "keepNext":
                EnsurePPr(p).KeepNext = ToggleParagraph<KeepNext>(name, value);
                break;

            case "keepLines":
                EnsurePPr(p).KeepLines = ToggleParagraph<KeepLines>(name, value);
                break;

            case "pageBreakBefore":
                EnsurePPr(p).PageBreakBefore = ToggleParagraph<PageBreakBefore>(name, value);
                break;

            case "widowControl":
                EnsurePPr(p).WidowControl = ToggleParagraph<WidowControl>(name, value);
                break;

            case "outlineLevel":
                EnsurePPr(p).OutlineLevel = ParseOutlineLevel(name, value);
                break;

            case "tabStops":
                ApplyTabStops(p, value);
                break;

            case var fanout when RunFanoutProps.Contains(fanout, StringComparer.Ordinal):
                foreach (var run in EnsureRun(p))
                {
                    SetRunProp(run, fanout, value);
                }

                break;

            default:
                throw UnsupportedProp(name, "p", ParagraphProps);
        }
    }

    /// <summary>A present/absent toggle paragraph element (false removes it, true writes it on).</summary>
    private static T? ToggleParagraph<T>(string name, string value)
        where T : OnOffType, new() =>
        ParseBool(name, value) ? new T() : null;

    private static OutlineLevel? ParseOutlineLevel(string name, string value)
    {
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var level) && level is >= 0 and <= 9)
        {
            return new OutlineLevel { Val = level };
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{name}' must be an integer 0–9 (a level-1 heading is 0), got '{value}'.",
            "Pass a 0-based outline level, e.g. outlineLevel=0 for a top-level heading.");
    }

    private static SpacingBetweenLines EnsureSpacing(Paragraph p)
    {
        var pPr = EnsurePPr(p);
        return pPr.SpacingBetweenLines ??= new SpacingBetweenLines();
    }

    private static Indentation EnsureIndentation(Paragraph p)
    {
        var pPr = EnsurePPr(p);
        return pPr.Indentation ??= new Indentation();
    }

    /// <summary>Spacing in points -> w:spacing twentieths-of-a-point string.</summary>
    private static string ParseSpacingTwentieths(string name, string value)
    {
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var points) && points is >= 0 and <= 1584)
        {
            return ((int)Math.Round(points * 20)).ToString(CultureInfo.InvariantCulture);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{name}' must be spacing in points (0–1584), got '{value}'.",
            $"Pass points, e.g. {name}=6 for 6pt.");
    }

    /// <summary>Indentation in centimeters -> w:ind twips string (negative pulls into the margin).</summary>
    private static string ParseIndentTwips(string name, string value)
    {
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var cm) && cm is >= -10 and <= 30)
        {
            return ((int)Math.Round(cm * (1440.0 / 2.54))).ToString(CultureInfo.InvariantCulture);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{name}' must be an indentation in centimeters (-10–30), got '{value}'.",
            $"Pass centimeters, e.g. {name}=1.5 for a 1.5cm indent.");
    }

    /// <summary>
    /// Writes w:spacing @line/@lineRule. <paramref name="value"/> is either a bare number
    /// (a line-height multiple -> @lineRule=auto, @line=round(multiple*240)) or a JSON object
    /// {atLeast: pts} / {exactly: pts} (-> @lineRule=atLeast|exactly, @line=pts*20). The @before/@after
    /// already on the same w:spacing are preserved (EnsureSpacing only touches @line/@lineRule).
    /// </summary>
    private static void ApplyLineSpacing(Paragraph p, string value)
    {
        var spacing = EnsureSpacing(p);

        // Bare number: a line-height multiple.
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var multiple))
        {
            if (multiple is <= 0 or > 132)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"lineSpacing multiple must be a positive number (e.g. 1.5, 2), got '{value}'.",
                    "Pass a line-height multiple like lineSpacing=1.5, or an object {\"atLeast\":12} / {\"exactly\":14}.");
            }

            spacing.Line = ((int)Math.Round(multiple * 240)).ToString(CultureInfo.InvariantCulture);
            spacing.LineRule = LineSpacingRuleValues.Auto;
            return;
        }

        // Object form: {atLeast: points} or {exactly: points}.
        var obj = TryParseJson(value) as JsonObject ?? throw LineSpacingInvalid(value);
        var hasAtLeast = obj.TryGetPropertyValue("atLeast", out var atLeastNode);
        var hasExactly = obj.TryGetPropertyValue("exactly", out var exactlyNode);
        if (hasAtLeast == hasExactly)
        {
            // both or neither set.
            throw LineSpacingInvalid(value);
        }

        var pointsNode = hasAtLeast ? atLeastNode : exactlyNode;
        if (pointsNode is null ||
            !double.TryParse(pointsNode.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var points) ||
            points is <= 0 or > 1584)
        {
            throw LineSpacingInvalid(value);
        }

        spacing.Line = ((int)Math.Round(points * 20)).ToString(CultureInfo.InvariantCulture);
        spacing.LineRule = hasAtLeast ? LineSpacingRuleValues.AtLeast : LineSpacingRuleValues.Exact;
    }

    private static AiofficeException LineSpacingInvalid(string value) => new(
        ErrorCodes.InvalidArgs,
        $"lineSpacing '{value}' is not a multiple or an {{atLeast|exactly: points}} object.",
        "Pass a line-height multiple (lineSpacing=2 for double), or {\"atLeast\":12} / {\"exactly\":14} in points.");

    /// <summary>
    /// Replaces the paragraph's w:tabs. <paramref name="value"/> is a JSON array of
    /// {pos: cm, align?, leader?}; an empty array clears the tab set entirely.
    /// </summary>
    private static void ApplyTabStops(Paragraph p, string value)
    {
        var pPr = EnsurePPr(p);
        var array = TryParseJson(value) as JsonArray ?? throw TabStopsInvalid(value);

        if (array.Count == 0)
        {
            pPr.Tabs = null;
            return;
        }

        var tabs = new Tabs();
        foreach (var node in array)
        {
            if (node is not JsonObject obj || !obj.TryGetPropertyValue("pos", out var posNode) || posNode is null ||
                !double.TryParse(posNode.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var cm) ||
                cm is < 0 or > 56)
            {
                throw TabStopsInvalid(value);
            }

            var align = obj.TryGetPropertyValue("align", out var alignNode) && alignNode is not null
                ? ParseTabAlign(alignNode.ToString())
                : TabStopValues.Left;
            var leader = obj.TryGetPropertyValue("leader", out var leaderNode) && leaderNode is not null
                ? ParseTabLeader(leaderNode.ToString())
                : (TabStopLeaderCharValues?)null;

            var tab = new TabStop
            {
                Val = align,
                Position = (int)Math.Round(cm * 567),
            };
            if (leader is { } l)
            {
                tab.Leader = l;
            }

            tabs.AppendChild(tab);
        }

        pPr.Tabs = tabs;
    }

    private static AiofficeException TabStopsInvalid(string value) => new(
        ErrorCodes.InvalidArgs,
        $"tabStops must be an array of {{pos(cm), align?, leader?}} (or [] to clear), got '{value}'.",
        "Example: tabStops=[{\"pos\":5,\"align\":\"decimal\",\"leader\":\"dot\"}].");

    private static TabStopValues ParseTabAlign(string raw) => raw.Trim().ToLowerInvariant() switch
    {
        "left" => TabStopValues.Left,
        "center" or "centre" => TabStopValues.Center,
        "right" => TabStopValues.Right,
        "decimal" => TabStopValues.Decimal,
        "bar" => TabStopValues.Bar,
        _ => throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"tab align '{raw}' is not valid.",
            "Use left, center, right, decimal or bar.",
            candidates: ["left", "center", "right", "decimal", "bar"]),
    };

    private static TabStopLeaderCharValues ParseTabLeader(string raw) => raw.Trim().ToLowerInvariant() switch
    {
        "none" => TabStopLeaderCharValues.None,
        "dot" => TabStopLeaderCharValues.Dot,
        "hyphen" => TabStopLeaderCharValues.Hyphen,
        "underscore" => TabStopLeaderCharValues.Underscore,
        _ => throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"tab leader '{raw}' is not valid.",
            "Use none, dot, hyphen or underscore.",
            candidates: ["none", "dot", "hyphen", "underscore"]),
    };

    private static JsonNode? TryParseJson(string value)
    {
        try
        {
            return JsonNode.Parse(value);
        }
        catch (System.Text.Json.JsonException)
        {
            return null;
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
                EnsureRPr(run).Underline = new Underline { Val = ParseUnderline(name, value) };
                break;

            case "color":
                EnsureRPr(run).Color = new Color { Val = ParseHexColor(value) };
                break;

            case "fontSize":
                EnsureRPr(run).FontSize = new FontSize { Val = ParseFontSizeHalfPoints(value) };
                break;

            case "font":
                EnsureRPr(run).RunFonts = new RunFonts
                {
                    Ascii = value,
                    HighAnsi = value,
                    ComplexScript = value,
                };
                break;

            case "rtl":
                SetRunRtl(run, ParseBool(name, value));
                break;

            // 1.10 character typography.
            case "highlight":
                EnsureRPr(run).Highlight = new Highlight { Val = ParseHighlight(value) };
                break;

            case "strike":
                EnsureRPr(run).Strike = new Strike { Val = OnOffValue.FromBoolean(ParseBool(name, value)) };
                break;

            case "doubleStrike":
                EnsureRPr(run).DoubleStrike = new DoubleStrike { Val = OnOffValue.FromBoolean(ParseBool(name, value)) };
                break;

            case "smallCaps":
                EnsureRPr(run).SmallCaps = new SmallCaps { Val = OnOffValue.FromBoolean(ParseBool(name, value)) };
                break;

            case "allCaps":
                EnsureRPr(run).Caps = new Caps { Val = OnOffValue.FromBoolean(ParseBool(name, value)) };
                break;

            // superscript/subscript share one w:vertAlign — keep exactly one; baseline removes it.
            case "superscript":
                SetVertAlign(run, ParseBool(name, value) ? VerticalPositionValues.Superscript : null);
                break;

            case "subscript":
                SetVertAlign(run, ParseBool(name, value) ? VerticalPositionValues.Subscript : null);
                break;

            case "characterSpacing":
                EnsureRPr(run).Spacing = new DocumentFormat.OpenXml.Wordprocessing.Spacing { Val = ParseCharacterSpacingTwentieths(value) };
                break;

            default:
                throw UnsupportedProp(name, "run", RunProps);
        }
    }

    /// <summary>Sets (or clears) the run's single w:vertAlign — only one of super/subscript may exist.</summary>
    private static void SetVertAlign(Run run, VerticalPositionValues? position)
    {
        if (position is { } p)
        {
            EnsureRPr(run).VerticalTextAlignment = new VerticalTextAlignment { Val = p };
        }
        else if (run.RunProperties is { } rPr)
        {
            rPr.VerticalTextAlignment = null;
        }
    }

    /// <summary>
    /// The 'underline' value-shape: the BOOL form (true -> single, false -> none) is tried FIRST and
    /// stays byte-stable; otherwise a named underline style (case-insensitive) off <see cref="UnderlineStyles"/>.
    /// An unrecognized string is invalid_args with the candidate list.
    /// </summary>
    public static UnderlineValues ParseUnderline(string name, string value)
    {
        switch (value.ToLowerInvariant())
        {
            case "true" or "1" or "on" or "yes":
                return UnderlineValues.Single;
            case "false" or "0" or "off" or "no":
                return UnderlineValues.None;
        }

        var match = UnderlineStyles.FirstOrDefault(s => string.Equals(s, value, StringComparison.OrdinalIgnoreCase));
        if (match is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{name}' expects true/false or an underline style, got '{value}'.",
                $"Pass {name}=true/false, or a style: {string.Join(", ", UnderlineStyles)}.",
                candidates: UnderlineStyles);
        }

        return new UnderlineValues(match);
    }

    /// <summary>A named highlight color (case-insensitive), or invalid_args with the fixed candidate list.</summary>
    public static HighlightColorValues ParseHighlight(string value)
    {
        var match = HighlightColors.FirstOrDefault(c => string.Equals(c, value, StringComparison.OrdinalIgnoreCase));
        if (match is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{value}' is not a Word highlight color (w:highlight has no hex form; use shading for an arbitrary fill).",
                $"Use one of: {string.Join(", ", HighlightColors)}.",
                candidates: HighlightColors);
        }

        return new HighlightColorValues(match);
    }

    /// <summary>Character spacing in points -> w:spacing @val in twentieths of a point (may be negative to condense).</summary>
    public static int ParseCharacterSpacingTwentieths(string value)
    {
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var points) && points is >= -158 and <= 158)
        {
            return (int)Math.Round(points * 20);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'characterSpacing' must be a number of points (-158–158, may be negative to condense), got '{value}'.",
            "Pass points, e.g. characterSpacing=1 (expand 1pt) or characterSpacing=-0.5 (condense).");
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

    /// <summary>
    /// Sets paragraph right-to-left flow: w:bidi on/off. Turning it on also
    /// right-aligns the paragraph unless it already carries an explicit
    /// alignment (Word's New ▸ RTL paragraph behavior); turning it off clears
    /// the bidi flag and any right alignment we added.
    /// </summary>
    public static void SetParagraphRtl(Paragraph p, bool rtl)
    {
        var pPr = EnsurePPr(p);
        if (rtl)
        {
            pPr.BiDi = new BiDi();
            pPr.Justification ??= new Justification { Val = JustificationValues.Right };
        }
        else
        {
            pPr.BiDi = new BiDi { Val = OnOffValue.FromBoolean(false) };
        }
    }

    /// <summary>Sets a run's right-to-left mark (w:rtl) for mixed-direction content.</summary>
    public static void SetRunRtl(Run run, bool rtl) =>
        EnsureRPr(run).RightToLeftText = new RightToLeftText { Val = OnOffValue.FromBoolean(rtl) };

    /// <summary>Reads a paragraph's direction: true when w:bidi is on.</summary>
    public static bool IsParagraphRtl(Paragraph p) =>
        IsOn(p.ParagraphProperties?.BiDi) == true;

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
            ["fontFamily"] = "font",
            ["typeface"] = "font",
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
