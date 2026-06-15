using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Embedded fonts (v1.5.0). A deck can carry its own font files so it renders the
/// same on a machine that does not have them installed. Each embedded font is a
/// <see cref="FontPart"/> referenced from a <c>p:font</c> slot
/// (<c>regular</c>/<c>bold</c>/<c>italic</c>/<c>boldItalic</c>) inside a
/// <c>p:embeddedFont</c> in the presentation's <c>p:embeddedFontLst</c>:
/// <list type="bullet">
/// <item><c>add /fonts {type:font, src:"MyFont.ttf"}</c> embeds the sandbox-resolved
///   font file as a font part and registers it (regular slot by default); the
///   media type is sniffed (ttf/otf).</item>
/// <item><c>get /fonts</c> lists every embedded font; <c>get /fonts/font[@name=...]</c>
///   reports one.</item>
/// <item><c>remove /fonts/font[@name=...]</c> drops the registration and its parts.</item>
/// </list>
/// We can only embed a font FILE the agent supplies — there is no system-font
/// lookup — so <c>src</c> is required and must point inside the workspace.
/// </summary>
internal static class PptxFonts
{
    private static readonly IReadOnlyList<string> AddPropKeys = ["src", "name", "embedAll", "bold", "italic", "boldItalic"];

    /// <summary>The four style slots a p:embeddedFont can fill, in schema order.</summary>
    private enum Slot
    {
        Regular,
        Bold,
        Italic,
        BoldItalic,
    }

    // ---- add ----------------------------------------------------------------

    /// <summary>
    /// add /fonts {type:font}: embeds the sandbox-resolved props.src font file and
    /// registers it under props.name (defaulting to the font file name). With
    /// embedAll plus the per-style srcs (bold/italic/boldItalic), all four slots are
    /// embedded; otherwise just the regular slot. Returns the canonical font path.
    /// </summary>
    public static string Add(PresentationPart presentation, JsonObject? props, Workspace workspace)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddPropKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown font prop '{key}'.",
                    "add font props: src (required, the regular .ttf/.otf), name (the typeface name), " +
                    "embedAll, and bold/italic/boldItalic (per-style files). " +
                    "Embedding needs the font FILE — aioffice cannot pull a system font.",
                    candidates: AddPropKeys);
            }
        }

        if (!props.TryGetPropertyValue("src", out var srcNode) || srcNode is null || J.ScalarText(srcNode).Trim().Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add font requires props.src (the regular-style font file).",
                "Point src at a .ttf/.otf inside the workspace, e.g. " +
                "{\"op\":\"add\",\"path\":\"/fonts\",\"type\":\"font\",\"props\":{\"src\":\"MyFont.ttf\"}}. " +
                "Embedding needs the font FILE — aioffice cannot pull a system font.");
        }

        // Sandbox first: a src outside the workspace is sandbox_denied, never read.
        var regularSrc = workspace.Resolve(J.ScalarText(srcNode).Trim(), mustExist: true);
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null && J.ScalarText(nameNode).Trim().Length > 0
            ? J.ScalarText(nameNode).Trim()
            : Path.GetFileNameWithoutExtension(regularSrc);

        var embeddedFont = ResolveByName(presentation, name);
        if (embeddedFont is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A font named '{name}' is already embedded.",
                "Remove it first ({\"op\":\"remove\",\"path\":\"" + PptxAddress.CanonicalFontPath(name) + "\"}) " +
                "then re-add, or pick a different name.");
        }

        // Collect the (slot -> source) embeds. The regular slot is always embedded;
        // each per-style src (bold/italic/boldItalic) the agent supplies brings in
        // that style. We never reuse the regular file for another style — embedding a
        // regular face as "bold" would silently misrepresent the glyphs.
        var embeds = new List<(Slot Slot, string Src)> { (Slot.Regular, regularSrc) };
        var embedAll = props.TryGetPropertyValue("embedAll", out var allNode) && allNode is not null && AsBool("embedAll", allNode);

        foreach (var (slot, key) in new[] { (Slot.Bold, "bold"), (Slot.Italic, "italic"), (Slot.BoldItalic, "boldItalic") })
        {
            if (props.TryGetPropertyValue(key, out var styleNode) && styleNode is not null && J.ScalarText(styleNode).Trim().Length > 0)
            {
                embeds.Add((slot, workspace.Resolve(J.ScalarText(styleNode).Trim(), mustExist: true)));
            }
        }

        // embedAll declares the intent to embed all four style variants — it only
        // holds when their files were actually supplied. Without them we cannot honor
        // it (there is no system-font lookup), so we say so rather than fake the glyphs.
        if (embedAll && embeds.Count < 4)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "embedAll needs the bold, italic and boldItalic font files too.",
                "Supply bold/italic/boldItalic srcs alongside src, e.g. " +
                "{\"src\":\"Brand.ttf\",\"bold\":\"Brand-Bold.ttf\",\"italic\":\"Brand-Italic.ttf\"," +
                "\"boldItalic\":\"Brand-BoldItalic.ttf\",\"embedAll\":true}; " +
                "or drop embedAll to embed just the regular face. aioffice cannot synthesize a style from the regular file.");
        }

        var slots = new List<(Slot Slot, string RelId)>();
        foreach (var (slot, src) in embeds)
        {
            RequireFontFile(src);
            var part = presentation.AddFontPart(FontPartType.FontTtf);
            using (var stream = File.OpenRead(src))
            {
                part.FeedData(stream);
            }

            slots.Add((slot, presentation.GetIdOfPart(part)));
        }

        var fontElement = new P.EmbeddedFont(new P.Font { Typeface = name });
        foreach (var (slot, relId) in slots)
        {
            fontElement.Append(BuildSlot(slot, relId));
        }

        var list = EnsureList(presentation);
        list.Append(fontElement);

        return PptxAddress.CanonicalFontPath(name);
    }

    // ---- get / list ---------------------------------------------------------

    /// <summary>get /fonts: every embedded font with its name, style slots and total embedded bytes.</summary>
    public static object List(PresentationPart presentation)
    {
        var fonts = EmbeddedFonts(presentation);
        return new
        {
            Path = "/fonts",
            EmbedTrueTypeFonts = presentation.Presentation?.EmbedTrueTypeFonts?.Value,
            Count = fonts.Count,
            Fonts = fonts.Select(f => (object)Project(presentation, f)).ToList(),
        };
    }

    /// <summary>get /fonts/font[@name=...]: one embedded font's name, slots and bytes.</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var font = ResolveByName(presentation, address.FontName!) ?? throw NotFound(presentation, address.FontName!);
        return Project(presentation, font);
    }

    private static object Project(PresentationPart presentation, P.EmbeddedFont font)
    {
        var name = font.Font?.Typeface?.Value ?? string.Empty;
        var slots = new List<string>();
        long bytes = 0;
        foreach (var (slot, relId) in SlotsOf(font))
        {
            slots.Add(SlotName(slot));
            if (relId is { Length: > 0 } && presentation.TryGetPartById(relId, out var part) && part is FontPart fontPart)
            {
                bytes += FontPartSize(fontPart);
            }
        }

        return new
        {
            Path = PptxAddress.CanonicalFontPath(name),
            Name = name,
            Styles = slots,
            Size = bytes,
        };
    }

    // ---- remove -------------------------------------------------------------

    /// <summary>remove /fonts/font[@name=...]: drops the registration and every font part it referenced.</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var font = ResolveByName(presentation, address.FontName!) ?? throw NotFound(presentation, address.FontName!);
        var name = font.Font?.Typeface?.Value ?? address.FontName!;

        // Drop the referenced font parts so removing the element does not orphan them.
        foreach (var (_, relId) in SlotsOf(font))
        {
            if (relId is { Length: > 0 } && presentation.TryGetPartById(relId, out _))
            {
                presentation.DeletePart(relId);
            }
        }

        var list = (P.EmbeddedFontList?)font.Parent;
        font.Remove();

        // An empty p:embeddedFontLst is not valid (it needs >=1 child); drop it and
        // the now-meaningless embedTrueTypeFonts flag when the last font is gone.
        if (list is not null && !list.Elements<P.EmbeddedFont>().Any())
        {
            list.Remove();
            if (presentation.Presentation is { } pres)
            {
                pres.EmbedTrueTypeFonts = null;
            }
        }

        return PptxAddress.CanonicalFontPath(name);
    }

    // ---- resolution ---------------------------------------------------------

    /// <summary>Every p:embeddedFont in the deck, in document order.</summary>
    private static List<P.EmbeddedFont> EmbeddedFonts(PresentationPart presentation) =>
        presentation.Presentation?.EmbeddedFontList?.Elements<P.EmbeddedFont>().ToList() ?? [];

    private static P.EmbeddedFont? ResolveByName(PresentationPart presentation, string name) =>
        EmbeddedFonts(presentation).FirstOrDefault(f =>
            string.Equals(f.Font?.Typeface?.Value, name, StringComparison.OrdinalIgnoreCase));

    /// <summary>The (slot, relationship-id) pairs a p:embeddedFont fills, in schema order.</summary>
    private static IEnumerable<(Slot Slot, string? RelId)> SlotsOf(P.EmbeddedFont font)
    {
        if (font.RegularFont is { } regular)
        {
            yield return (Slot.Regular, regular.Id?.Value);
        }

        if (font.BoldFont is { } bold)
        {
            yield return (Slot.Bold, bold.Id?.Value);
        }

        if (font.ItalicFont is { } italic)
        {
            yield return (Slot.Italic, italic.Id?.Value);
        }

        if (font.BoldItalicFont is { } boldItalic)
        {
            yield return (Slot.BoldItalic, boldItalic.Id?.Value);
        }
    }

    /// <summary>The presentation's p:embeddedFontLst, created (in schema position) and the flag set on first use.</summary>
    private static P.EmbeddedFontList EnsureList(PresentationPart presentation)
    {
        var presentationElement = presentation.Presentation
            ?? throw Corrupt("the package has no p:presentation");

        if (presentationElement.EmbeddedFontList is { } existing)
        {
            return existing;
        }

        var list = new P.EmbeddedFontList();

        // embeddedFontLst sits before defaultTextStyle in CT_Presentation; insert it
        // there when defaultTextStyle exists, otherwise append (the validator agrees).
        if (presentationElement.DefaultTextStyle is { } defaultTextStyle)
        {
            presentationElement.InsertBefore(list, defaultTextStyle);
        }
        else
        {
            presentationElement.Append(list);
        }

        presentationElement.EmbedTrueTypeFonts = true;
        return list;
    }

    private static OpenXmlElement BuildSlot(Slot slot, string relId) => slot switch
    {
        Slot.Regular => new P.RegularFont { Id = relId },
        Slot.Bold => new P.BoldFont { Id = relId },
        Slot.Italic => new P.ItalicFont { Id = relId },
        _ => new P.BoldItalicFont { Id = relId },
    };

    private static string SlotName(Slot slot) => slot switch
    {
        Slot.Regular => "regular",
        Slot.Bold => "bold",
        Slot.Italic => "italic",
        _ => "boldItalic",
    };

    /// <summary>
    /// Verifies a file is a font (an sfnt-wrapped TrueType/OpenType: the magic is
    /// 0x00010000, "OTTO", "true" or "ttcf"). We embed the bytes verbatim, so the
    /// guard is the only thing standing between "embed a font" and "embed a random
    /// file as a font" — sniffing the header keeps us honest. Both .ttf and .otf
    /// embed as the same ttf font part (the obfuscated variant is a PowerPoint
    /// export concern we do not emit).
    /// </summary>
    private static void RequireFontFile(string src)
    {
        byte[] header;
        using (var stream = File.OpenRead(src))
        {
            header = new byte[(int)Math.Min(4, stream.Length)];
            _ = stream.Read(header, 0, header.Length);
        }

        var isSfnt = header.Length >= 4 &&
            ((header[0] == 0x00 && header[1] == 0x01 && header[2] == 0x00 && header[3] == 0x00) || // TrueType 1.0
             (header[0] == (byte)'O' && header[1] == (byte)'T' && header[2] == (byte)'T' && header[3] == (byte)'O') || // OpenType/CFF
             (header[0] == (byte)'t' && header[1] == (byte)'r' && header[2] == (byte)'u' && header[3] == (byte)'e') || // TrueType (Apple)
             (header[0] == (byte)'t' && header[1] == (byte)'t' && header[2] == (byte)'c' && header[3] == (byte)'f'));  // TrueType collection

        if (!isSfnt)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{Path.GetFileName(src)}' does not look like a TrueType/OpenType font file.",
                "Point src at a real .ttf/.otf font file; aioffice embeds the font FILE verbatim and " +
                "cannot synthesize one from a name.");
        }
    }

    private static long FontPartSize(FontPart part)
    {
        using var stream = part.GetStream();
        return stream.Length;
    }

    private static bool AsBool(string key, JsonNode? node)
    {
        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var flag))
            {
                return flag;
            }

            if (value.TryGetValue<string>(out var text) && bool.TryParse(text, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Property '{key}' is not a boolean: {node?.ToJsonString() ?? "null"}",
            "Use true or false.");
    }

    private static AiofficeException NotFound(PresentationPart presentation, string name)
    {
        var fonts = EmbeddedFonts(presentation);
        return new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No embedded font named '{name}'; the deck has {fonts.Count} embedded font(s).",
            "Run 'aioffice get <file> /fonts' to list embedded fonts and their names.",
            candidates: [.. fonts.Take(10).Select(f => PptxAddress.CanonicalFontPath(f.Font?.Typeface?.Value ?? string.Empty))]);
    }

    private static AiofficeException Corrupt(string detail) => new(
        ErrorCodes.FormatCorrupt,
        $"The presentation is malformed: {detail}.",
        "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
}
