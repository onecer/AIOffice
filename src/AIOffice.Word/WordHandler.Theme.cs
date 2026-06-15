using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace AIOffice.Word;

/// <summary>
/// v1.3.0 document theme &amp; color-scheme editing. The theme part
/// (<c>theme1.xml</c>) holds the <c>a:clrScheme</c> (dk1/lt1/dk2/lt2,
/// accent1..6, hlink/folHlink) and <c>a:fontScheme</c> (major/minor fonts) that
/// styles reference by name. <c>set /theme</c> edits those slots; a document
/// without a theme part gets a complete Office-default one created on demand so
/// the edit lands somewhere valid. <c>get /theme</c> reports the scheme.
/// </summary>
public sealed partial class WordHandler
{
    /// <summary>The theme color slots editable through set /theme -> their clrScheme child accessor.</summary>
    private static readonly string[] ThemeColorSlots =
        ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

    private static readonly string[] ThemeFontSlots = ["majorFont", "minorFont"];

    // ------------------------------------------------------------------- set

    /// <summary>
    /// <c>{"op":"set","path":"/theme","props":{"accent1":"38BDF8","majorFont":"Calibri Light",…}}</c>:
    /// edits the theme part's color scheme and/or font scheme slots.
    /// </summary>
    private static object ApplySetTheme(WordprocessingDocument doc, EditOp op)
    {
        var props = RequireProps(op);
        var theme = EnsureThemePart(doc);
        var elements = theme.ThemeElements ??= new A.ThemeElements();
        var scheme = elements.ColorScheme ??= BuildDefaultColorScheme();
        var fontScheme = elements.FontScheme ??= BuildDefaultFontScheme();

        var changed = new List<string>();
        foreach (var (rawName, node) in props)
        {
            var name = rawName;
            var value = NodeToString(node);

            if (ThemeColorSlots.Contains(name, StringComparer.Ordinal))
            {
                SetThemeColor(scheme, name, WordFormatting.ParseHexColor(value));
                changed.Add(name);
            }
            else if (name == "majorFont")
            {
                SetThemeFont(fontScheme.MajorFont ??= new A.MajorFont(new A.LatinFont { Typeface = "Calibri Light" }), value);
                changed.Add(name);
            }
            else if (name == "minorFont")
            {
                SetThemeFont(fontScheme.MinorFont ??= new A.MinorFont(new A.LatinFont { Typeface = "Calibri" }), value);
                changed.Add(name);
            }
            else
            {
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"Theme slot '{name}' is not editable.",
                    "Set color slots dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink, " +
                    "or fonts majorFont / minorFont.",
                    candidates: [.. ThemeColorSlots, .. ThemeFontSlots]);
            }
        }

        return new { op = "set", path = "/theme", type = "theme", changed };
    }

    /// <summary>Writes an explicit RGB hex into one clrScheme color slot, replacing whatever child it had.</summary>
    private static void SetThemeColor(A.ColorScheme scheme, string slot, string hex)
    {
        DocumentFormat.OpenXml.OpenXmlElement color = slot switch
        {
            "dk1" => scheme.Dark1Color ??= new A.Dark1Color(),
            "lt1" => scheme.Light1Color ??= new A.Light1Color(),
            "dk2" => scheme.Dark2Color ??= new A.Dark2Color(),
            "lt2" => scheme.Light2Color ??= new A.Light2Color(),
            "accent1" => scheme.Accent1Color ??= new A.Accent1Color(),
            "accent2" => scheme.Accent2Color ??= new A.Accent2Color(),
            "accent3" => scheme.Accent3Color ??= new A.Accent3Color(),
            "accent4" => scheme.Accent4Color ??= new A.Accent4Color(),
            "accent5" => scheme.Accent5Color ??= new A.Accent5Color(),
            "accent6" => scheme.Accent6Color ??= new A.Accent6Color(),
            "hlink" => scheme.Hyperlink ??= new A.Hyperlink(),
            _ => scheme.FollowedHyperlinkColor ??= new A.FollowedHyperlinkColor(),
        };

        color.RemoveAllChildren();
        color.AppendChild(new A.RgbColorModelHex { Val = hex });
    }

    private static void SetThemeFont(A.FontCollectionType font, string typeface)
    {
        var latin = font.GetFirstChild<A.LatinFont>();
        if (latin is null)
        {
            latin = new A.LatinFont();
            font.InsertAt(latin, 0);
        }

        latin.Typeface = typeface;
    }

    // ------------------------------------------------------------------- get

    /// <summary>get /theme -> the color scheme and font scheme as hex/name maps.</summary>
    private static Dictionary<string, object?> GetThemeProperties(WordprocessingDocument doc)
    {
        var theme = doc.MainDocumentPart?.ThemePart?.Theme;
        var scheme = theme?.ThemeElements?.ColorScheme;
        var fonts = theme?.ThemeElements?.FontScheme;

        return new Dictionary<string, object?>
        {
            ["name"] = theme?.Name?.Value,
            ["dk1"] = ThemeColorHex(scheme?.Dark1Color),
            ["lt1"] = ThemeColorHex(scheme?.Light1Color),
            ["dk2"] = ThemeColorHex(scheme?.Dark2Color),
            ["lt2"] = ThemeColorHex(scheme?.Light2Color),
            ["accent1"] = ThemeColorHex(scheme?.Accent1Color),
            ["accent2"] = ThemeColorHex(scheme?.Accent2Color),
            ["accent3"] = ThemeColorHex(scheme?.Accent3Color),
            ["accent4"] = ThemeColorHex(scheme?.Accent4Color),
            ["accent5"] = ThemeColorHex(scheme?.Accent5Color),
            ["accent6"] = ThemeColorHex(scheme?.Accent6Color),
            ["hlink"] = ThemeColorHex(scheme?.Hyperlink),
            ["folHlink"] = ThemeColorHex(scheme?.FollowedHyperlinkColor),
            ["majorFont"] = fonts?.MajorFont?.LatinFont?.Typeface?.Value,
            ["minorFont"] = fonts?.MinorFont?.LatinFont?.Typeface?.Value,
        };
    }

    /// <summary>The hex of a clrScheme slot: the explicit srgbClr, or the sysColor's lastClr cache.</summary>
    private static string? ThemeColorHex(A.Color2Type? color)
    {
        if (color is null)
        {
            return null;
        }

        if (color.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value is { } rgb)
        {
            return rgb.ToUpperInvariant();
        }

        return color.GetFirstChild<A.SystemColor>()?.LastColor?.Value?.ToUpperInvariant();
    }

    // ------------------------------------------------------------ theme part

    /// <summary>Returns the document theme, creating a complete Office-default theme part when the document has none.</summary>
    private static A.Theme EnsureThemePart(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The document has no main part to carry a theme.",
            "Re-export the file from Word.");

        var part = main.ThemePart;
        if (part is null)
        {
            part = main.AddNewPart<ThemePart>();
            part.Theme = BuildDefaultTheme();
        }

        return part.Theme ??= BuildDefaultTheme();
    }

    /// <summary>A complete, valid Office-default theme (clrScheme + fontScheme + fmtScheme).</summary>
    private static A.Theme BuildDefaultTheme()
    {
        var theme = new A.Theme(
            new A.ThemeElements(BuildDefaultColorScheme(), BuildDefaultFontScheme(), BuildDefaultFormatScheme()))
        {
            Name = "Office Theme",
        };

        // Declare a: on the theme root explicitly. A reopen+save emits the xmlns
        // declaration before the name attribute; authoring it here matches that
        // ordering so the round-trip law stays byte-identical.
        theme.AddNamespaceDeclaration("a", DrawingMainNamespace);
        return theme;
    }

    private static A.ColorScheme BuildDefaultColorScheme() => new(
        new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
        new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
        new A.Dark2Color(new A.RgbColorModelHex { Val = "44546A" }),
        new A.Light2Color(new A.RgbColorModelHex { Val = "E7E6E6" }),
        new A.Accent1Color(new A.RgbColorModelHex { Val = "4472C4" }),
        new A.Accent2Color(new A.RgbColorModelHex { Val = "ED7D31" }),
        new A.Accent3Color(new A.RgbColorModelHex { Val = "A5A5A5" }),
        new A.Accent4Color(new A.RgbColorModelHex { Val = "FFC000" }),
        new A.Accent5Color(new A.RgbColorModelHex { Val = "5B9BD5" }),
        new A.Accent6Color(new A.RgbColorModelHex { Val = "70AD47" }),
        new A.Hyperlink(new A.RgbColorModelHex { Val = "0563C1" }),
        new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" }))
    {
        Name = "Office",
    };

    private static A.FontScheme BuildDefaultFontScheme() => new(
        FillFont(new A.MajorFont(), "Calibri Light"),
        FillFont(new A.MinorFont(), "Calibri"))
    {
        Name = "Office",
    };

    /// <summary>Populates a major/minor font collection with the standard latin/east-asian/cs entries.</summary>
    private static T FillFont<T>(T font, string typeface)
        where T : A.FontCollectionType
    {
        font.AppendChild(new A.LatinFont { Typeface = typeface });
        font.AppendChild(new A.EastAsianFont { Typeface = string.Empty });
        font.AppendChild(new A.ComplexScriptFont { Typeface = string.Empty });
        return font;
    }

    private static A.FormatScheme BuildDefaultFormatScheme()
    {
        A.SolidFill Ph() => new(new A.SchemeColor { Val = A.SchemeColorValues.PhColor });
        A.Outline Line() => new(Ph()) { Width = 9525 };
        A.EffectStyle Effect() => new(new A.EffectList());

        return new A.FormatScheme(
            new A.FillStyleList(Ph(), Ph(), Ph()),
            new A.LineStyleList(Line(), Line(), Line()),
            new A.EffectStyleList(Effect(), Effect(), Effect()),
            new A.BackgroundFillStyleList(Ph(), Ph(), Ph()))
        {
            Name = "Office",
        };
    }
}
