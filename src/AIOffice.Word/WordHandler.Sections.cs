using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    private static readonly string[] SectionProps =
        ["orientation", "pageSize", "marginTop", "marginBottom", "marginLeft", "marginRight"];

    /// <summary>Known page sizes as portrait (width, height) in twentieths of a point.</summary>
    private static readonly Dictionary<string, (uint Width, uint Height)> PageSizes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["A4"] = (11906, 16838),
        ["Letter"] = (12240, 15840),
        ["A3"] = (16838, 23811),
    };

    private const double TwipsPerCm = 1440.0 / 2.54;

    // ------------------------------------------------------------- enumerate

    /// <summary>
    /// Every section of the document in document order: paragraph-level sectPr
    /// elements (section breaks) first, the body-level sectPr (the final,
    /// possibly only, section) last. Word's implicit single section IS the
    /// body sectPr; when even that is absent the document still has one
    /// conceptual section that <c>set /section[1]</c> materializes on demand.
    /// </summary>
    private static List<SectionProperties> EnumerateSections(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.Document?.Body is { } body
            ? [.. body.Descendants<SectionProperties>()]
            : [];

    /// <summary>Resolves /section[i] (1-based) or throws invalid_path with candidates.</summary>
    private static SectionProperties ResolveSection(WordprocessingDocument doc, DocPath path, bool createImplicit)
    {
        var segment = path.Segments[0];
        if (path.Segments.Count != 1 || segment.Id is not null)
        {
            throw SectionNotFound(doc, $"'{path.ToCanonicalString()}' is not a section path; sections are positional.");
        }

        var index = segment.Index ?? 1;
        var sections = EnumerateSections(doc);

        if (sections.Count == 0 && index == 1 && createImplicit &&
            doc.MainDocumentPart?.Document?.Body is { } body)
        {
            // Materialize the implicit single section as the body-level sectPr.
            var sectPr = new SectionProperties();
            body.AppendChild(sectPr);
            return sectPr;
        }

        if (index > sections.Count)
        {
            throw SectionNotFound(
                doc,
                sections.Count == 0
                    ? "This document has no explicit section properties yet."
                    : $"/section[{index}] does not exist; the document has {sections.Count} section(s).");
        }

        return sections[index - 1];
    }

    private static AiofficeException SectionNotFound(WordprocessingDocument doc, string message)
    {
        var count = Math.Max(EnumerateSections(doc).Count, 1);
        return new AiofficeException(
            ErrorCodes.InvalidPath,
            message,
            "Address sections positionally; /section[1] always exists (set materializes it on demand).",
            candidates: [.. Enumerable.Range(1, Math.Min(count, 5)).Select(n => $"/section[{n}]")]);
    }

    // ------------------------------------------------------------------- set

    /// <summary>
    /// <c>{"op":"set","path":"/section[1]","props":{"orientation":…,"pageSize":…,"marginTop":…}}</c>:
    /// page setup on one section. pageSize/orientation manage w:pgSz, margins
    /// manage w:pgMar (created with Word's defaults so the element stays
    /// schema-complete).
    /// </summary>
    private static object ApplySetSection(WordprocessingDocument doc, EditOp op)
    {
        var path = DocPath.Parse(op.Path);
        var sectPr = ResolveSection(doc, path, createImplicit: true);
        var props = RequireProps(op);

        string? orientation = null;
        string? sizeName = null;
        var margins = new List<(string Key, uint Twips)>();

        foreach (var (key, node) in props)
        {
            var value = NodeToString(node);
            switch (key)
            {
                case "orientation":
                    orientation = value.ToLowerInvariant();
                    if (orientation is not ("portrait" or "landscape"))
                    {
                        throw new AiofficeException(
                            ErrorCodes.InvalidArgs,
                            $"Unknown orientation '{value}'.",
                            "Use orientation portrait or landscape.",
                            candidates: ["portrait", "landscape"]);
                    }

                    break;

                case "pageSize":
                    if (!PageSizes.ContainsKey(value))
                    {
                        throw new AiofficeException(
                            ErrorCodes.InvalidArgs,
                            $"Unknown page size '{value}'.",
                            "Use one of: A4, Letter, A3.",
                            candidates: [.. PageSizes.Keys]);
                    }

                    sizeName = value;
                    break;

                case "marginTop" or "marginBottom" or "marginLeft" or "marginRight":
                    margins.Add((key, (uint)Math.Max(1, Math.Round(ParseLengthEmu(key, value) / 635.0))));
                    break;

                default:
                    throw new AiofficeException(
                        ErrorCodes.UnsupportedFeature,
                        $"Section property '{key}' is not supported.",
                        $"Did you mean '{WordFormatting.Nearest(key, SectionProps)}'? Supported: {string.Join(", ", SectionProps)}.",
                        candidates: SectionProps);
            }
        }

        if (sizeName is not null || orientation is not null)
        {
            ApplyPageSize(sectPr, sizeName, orientation);
        }

        foreach (var (key, twips) in margins)
        {
            var pgMar = EnsurePageMargin(sectPr);
            switch (key)
            {
                case "marginTop":
                    pgMar.Top = (int)twips;
                    break;
                case "marginBottom":
                    pgMar.Bottom = (int)twips;
                    break;
                case "marginLeft":
                    pgMar.Left = twips;
                    break;
                default:
                    pgMar.Right = twips;
                    break;
            }
        }

        var canonical = path.ToCanonicalString();
        return new { op = "set", path = canonical, type = "section" };
    }

    /// <summary>Applies size and/or orientation, swapping dimensions for landscape like Word does.</summary>
    private static void ApplyPageSize(SectionProperties sectPr, string? sizeName, string? orientation)
    {
        var pgSz = sectPr.GetFirstChild<PageSize>();
        if (pgSz is null)
        {
            pgSz = new PageSize { Width = PageSizes["A4"].Width, Height = PageSizes["A4"].Height };
            InsertSectionChild(sectPr, pgSz);
        }

        if (sizeName is not null)
        {
            var (w, h) = PageSizes[sizeName];
            pgSz.Width = w;
            pgSz.Height = h;
            pgSz.Orient = null; // size names are portrait; orientation re-applies below
        }

        var landscape = orientation == "landscape"
            || (orientation is null && pgSz.Orient?.Value == PageOrientationValues.Landscape);

        var width = pgSz.Width?.Value ?? PageSizes["A4"].Width;
        var height = pgSz.Height?.Value ?? PageSizes["A4"].Height;
        if (landscape != width > height)
        {
            (pgSz.Width, pgSz.Height) = (height, width);
        }

        pgSz.Orient = landscape ? PageOrientationValues.Landscape : null;
    }

    private static PageMargin EnsurePageMargin(SectionProperties sectPr)
    {
        var pgMar = sectPr.GetFirstChild<PageMargin>();
        if (pgMar is null)
        {
            // w:pgMar requires every attribute; start from Word's defaults (1in margins, 0.5in header/footer).
            pgMar = new PageMargin
            {
                Top = 1440,
                Right = 1440U,
                Bottom = 1440,
                Left = 1440U,
                Header = 720U,
                Footer = 720U,
                Gutter = 0U,
            };
            InsertSectionChild(sectPr, pgMar);
        }

        return pgMar;
    }

    /// <summary>CT_SectPr child order, used to insert pgSz/pgMar at their schema position.</summary>
    private static readonly Type[] SectPrOrder =
    [
        typeof(HeaderReference), typeof(FooterReference), typeof(FootnoteProperties), typeof(EndnoteProperties),
        typeof(SectionType), typeof(PageSize), typeof(PageMargin), typeof(PaperSource), typeof(PageBorders),
        typeof(LineNumberType), typeof(PageNumberType), typeof(Columns),
    ];

    private static void InsertSectionChild(SectionProperties sectPr, OpenXmlElement child)
    {
        var rank = Array.IndexOf(SectPrOrder, child.GetType());
        var before = sectPr.ChildElements.FirstOrDefault(existing =>
        {
            var existingRank = Array.IndexOf(SectPrOrder, existing.GetType());
            return existingRank > rank; // unknown (-1) children sort first and never push us back
        });

        if (before is null)
        {
            sectPr.AppendChild(child);
        }
        else
        {
            sectPr.InsertBefore(child, before);
        }
    }

    // ------------------------------------------------------------------- get

    /// <summary>get /section[i] data: orientation, page size and margins (cm).</summary>
    private static Dictionary<string, object?> GetSectionProperties(WordprocessingDocument doc, DocPath path)
    {
        var sections = EnumerateSections(doc);
        var index = path.Segments[0].Index ?? 1;

        // The implicit single section reads as portrait defaults without mutating the file.
        SectionProperties? sectPr;
        if (sections.Count == 0 && index == 1 && path.Segments.Count == 1 && path.Segments[0].Id is null)
        {
            sectPr = null;
        }
        else
        {
            sectPr = ResolveSection(doc, path, createImplicit: false);
        }

        var pgSz = sectPr?.GetFirstChild<PageSize>();
        var pgMar = sectPr?.GetFirstChild<PageMargin>();

        var width = pgSz?.Width?.Value;
        var height = pgSz?.Height?.Value;
        var landscape = pgSz?.Orient?.Value == PageOrientationValues.Landscape;
        var sizeName = MatchPageSizeName(width, height);

        return new Dictionary<string, object?>
        {
            ["index"] = index,
            ["sections"] = Math.Max(sections.Count, 1),
            ["orientation"] = landscape ? "landscape" : "portrait",
            ["pageSize"] = sizeName,
            ["widthCm"] = width is { } w ? TwipsToCm(w) : null,
            ["heightCm"] = height is { } h ? TwipsToCm(h) : null,
            ["marginTopCm"] = pgMar?.Top?.Value is { } t ? TwipsToCm(t) : null,
            ["marginBottomCm"] = pgMar?.Bottom?.Value is { } b ? TwipsToCm(b) : null,
            ["marginLeftCm"] = pgMar?.Left?.Value is { } l ? TwipsToCm((long)l) : null,
            ["marginRightCm"] = pgMar?.Right?.Value is { } r ? TwipsToCm((long)r) : null,
        };
    }

    /// <summary>The size name whose portrait dimensions match (either orientation), else "custom"/null.</summary>
    private static string? MatchPageSizeName(uint? width, uint? height)
    {
        if (width is not { } w || height is not { } h)
        {
            return null;
        }

        foreach (var (name, (pw, ph)) in PageSizes)
        {
            if ((w == pw && h == ph) || (w == ph && h == pw))
            {
                return name;
            }
        }

        return "custom";
    }

    private static double TwipsToCm(long twips) =>
        Math.Round(twips / TwipsPerCm, 2);

    private static string SectionPath(int index) =>
        string.Create(CultureInfo.InvariantCulture, $"/section[{index}]");
}
