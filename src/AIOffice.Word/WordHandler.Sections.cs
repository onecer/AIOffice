using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    private static readonly string[] SectionProps =
        ["orientation", "pageSize", "marginTop", "marginBottom", "marginLeft", "marginRight", "columns", "columnGap", "pageBorder"];

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
        JsonNode? columnsNode = null;
        uint? columnGapTwips = null;
        JsonNode? pageBorderNode = null;
        var pageBorderSet = false;

        foreach (var (key, node) in props)
        {
            var value = NodeToString(node);
            switch (key)
            {
                case "pageBorder":
                    pageBorderNode = node;
                    pageBorderSet = true;
                    break;

                case "columns":
                    columnsNode = node;
                    break;

                case "columnGap":
                    columnGapTwips = (uint)Math.Max(0, Math.Round(ParseLengthEmu(key, value) * TwipsPerEmu));
                    break;

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

        if (columnsNode is not null || columnGapTwips is not null)
        {
            ApplyColumns(sectPr, columnsNode, columnGapTwips);
        }

        if (pageBorderSet)
        {
            ApplyPageBorder(sectPr, pageBorderNode);
        }

        var canonical = path.ToCanonicalString();
        return new { op = "set", path = canonical, type = "section" };
    }

    /// <summary>Default gap between equal-width columns (Word's 0.5"): 720 twips.</summary>
    private const uint DefaultColumnGapTwips = 720;

    /// <summary>
    /// Builds the section's <c>w:cols</c>. <c>columns</c> as an integer N gives N
    /// equal columns separated by <c>columnGap</c> (or Word's 0.5" default); as an
    /// array of <c>{width[,gap]}</c> objects it gives explicit per-column widths
    /// (<c>equalWidth=0</c> plus one <c>w:col</c> each). Passing only
    /// <c>columnGap</c> retunes the gap of the existing layout.
    /// </summary>
    private static void ApplyColumns(SectionProperties sectPr, JsonNode? columnsNode, uint? gapTwips)
    {
        var existing = sectPr.GetFirstChild<Columns>();

        // columnGap alone: adjust the running layout (default to 1 column if none).
        if (columnsNode is null)
        {
            var cols = existing ?? new Columns { ColumnCount = 1 };
            cols.Space = (gapTwips ?? DefaultColumnGapTwips).ToString(CultureInfo.InvariantCulture);
            if (existing is null)
            {
                InsertSectionChild(sectPr, cols);
            }

            return;
        }

        existing?.Remove();

        if (columnsNode is JsonArray widths)
        {
            ApplyExplicitColumns(sectPr, widths, gapTwips);
            return;
        }

        var count = int.TryParse(NodeToString(columnsNode), out var n) ? n : 0;
        if (count < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"columns must be a positive integer or an array of column widths, got '{NodeToString(columnsNode)}'.",
                "Use columns:2 for two equal columns, or columns:[{\"width\":\"6cm\"},{\"width\":\"4cm\"}] for explicit widths.");
        }

        var equalCols = new Columns
        {
            ColumnCount = (short)count,
            EqualWidth = OnOffValue.FromBoolean(true),
            Space = (gapTwips ?? DefaultColumnGapTwips).ToString(CultureInfo.InvariantCulture),
        };
        InsertSectionChild(sectPr, equalCols);
    }

    private static void ApplyExplicitColumns(SectionProperties sectPr, JsonArray widths, uint? gapTwips)
    {
        if (widths.Count < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "columns array needs at least one column with a width.",
                "Example: columns:[{\"width\":\"6cm\"},{\"width\":\"4cm\"}].");
        }

        var gap = (gapTwips ?? DefaultColumnGapTwips);
        var cols = new Columns
        {
            ColumnCount = (short)widths.Count,
            EqualWidth = OnOffValue.FromBoolean(false),
            Space = gap.ToString(CultureInfo.InvariantCulture),
        };

        for (var i = 0; i < widths.Count; i++)
        {
            var entry = widths[i] as JsonObject;
            var widthRaw = entry?["width"] is { } w ? NodeToString(w) : NodeToString(widths[i]);
            var widthTwips = (long)Math.Max(1, Math.Round(ParseLengthEmu($"columns[{i}].width", widthRaw) * TwipsPerEmu));
            var colGap = entry?["gap"] is { } g
                ? (uint)Math.Max(0, Math.Round(ParseLengthEmu($"columns[{i}].gap", NodeToString(g)) * TwipsPerEmu))
                : gap;

            var col = new Column { Width = widthTwips.ToString(CultureInfo.InvariantCulture) };

            // The last column carries no trailing space (Word omits it).
            if (i < widths.Count - 1)
            {
                col.Space = colGap.ToString(CultureInfo.InvariantCulture);
            }

            cols.AppendChild(col);
        }

        InsertSectionChild(sectPr, cols);
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

    /// <summary>CT_SectPr child order, used to insert pgSz/pgMar/titlePg at their schema position.</summary>
    private static readonly Type[] SectPrOrder =
    [
        typeof(HeaderReference), typeof(FooterReference), typeof(FootnoteProperties), typeof(EndnoteProperties),
        typeof(SectionType), typeof(PageSize), typeof(PageMargin), typeof(PaperSource), typeof(PageBorders),
        typeof(LineNumberType), typeof(PageNumberType), typeof(Columns), typeof(FormProtection),
        typeof(VerticalTextAlignmentOnPage), typeof(NoEndnote), typeof(TitlePage), typeof(TextDirection),
        typeof(BiDi), typeof(GutterOnRight), typeof(DocGrid), typeof(PrinterSettingsReference),
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

    // ----------------------------------------------------------- add / remove

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[5]","type":"sectionBreak","props":{"kind":"nextPage|continuous"}}</c>:
    /// ends a section at the target paragraph. In the OOXML section model a
    /// paragraph-level w:pPr/w:sectPr marks the END of a section, so the new
    /// sectPr lands on p[5] and clones the setup of the section that governed it
    /// (ultimately the trailing body-level sectPr, materialized on demand) —
    /// both halves start out looking identical, exactly like Word's own
    /// Insert ▸ Section Break.
    /// </summary>
    private static object ApplyAddSectionBreak(WordprocessingDocument doc, EditOp op)
    {
        var kind = op.Props?["kind"] is { } kindNode ? NodeToString(kindNode) : "nextPage";
        if (kind is not ("nextPage" or "continuous"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"sectionBreak kind '{kind}' is not supported.",
                "Use kind nextPage (the new section starts on a fresh page) or continuous (same page).",
                candidates: ["nextPage", "continuous"]);
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph || paragraph.Parent is not Body body)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Section breaks end a top-level body paragraph, not '{anchor.Type}' at {anchor.CanonicalPath}.",
                "Target a direct /body/p[n] paragraph; sections cannot break inside tables, headers or footers.");
        }

        if (paragraph.ParagraphProperties?.SectionProperties is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"{anchor.CanonicalPath} already ends a section.",
                "Pick a different paragraph, or set the existing /section[i] properties instead.");
        }

        // The final section must stay addressable, so the body-level sectPr is materialized first.
        var bodySectPr = body.Elements<SectionProperties>().FirstOrDefault();
        if (bodySectPr is null)
        {
            bodySectPr = new SectionProperties();
            body.AppendChild(bodySectPr);
        }

        var governing = GoverningSectionProperties(body, paragraph) ?? bodySectPr;
        var sectPr = (SectionProperties)governing.CloneNode(true);
        sectPr.RemoveAllChildren<SectionType>(); // the break kind belongs to the new break, not the clone source
        if (kind == "continuous")
        {
            InsertSectionChild(sectPr, new SectionType { Val = SectionMarkValues.Continuous });
        }

        var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();
        pPr.SectionProperties = sectPr;

        var sections = EnumerateSections(doc);
        var index = sections.FindIndex(s => ReferenceEquals(s, sectPr)) + 1;
        return new
        {
            op = "add",
            type = "sectionBreak",
            path = SectionPath(index),
            kind,
            sections = sections.Count,
        };
    }

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[3]","type":"columnBreak"}</c>: a hard column
    /// break (<c>w:br w:type="column"</c>). By default the break is appended to the
    /// end of the target paragraph's content (text flows to the next column there);
    /// position before puts the break at the paragraph's start instead. Honors the
    /// --track flag like other text-content mutations by inserting the break in an
    /// w:ins run when tracking is on.
    /// </summary>
    private static object ApplyAddColumnBreak(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A column break attaches to a paragraph, not '{anchor.Type}' at {anchor.CanonicalPath}.",
                "Target a paragraph (e.g. /body/p[3]); the break sends following text to the next column.");
        }

        var position = op.Position;
        if (position is not (null or "before" or "after"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"columnBreak position '{position}' is not valid.",
                "Use position before (break at the paragraph start) or after (the default, break at its end).",
                candidates: ["before", "after"]);
        }

        var breakRun = new Run(new Break { Type = BreakValues.Column });

        if (session.Track)
        {
            var author = session.ResolveAuthor(op.Props?.DeepClone().AsObject());
            var ins = NewTrackChange(new InsertedRun(), author, DateTime.UtcNow, NextRevisionId(doc));
            ins.AppendChild(breakRun);
            InsertBreak(paragraph, ins, position);
        }
        else
        {
            InsertBreak(paragraph, breakRun, position);
        }

        return new { op = "add", type = "columnBreak", path = anchor.CanonicalPath, placement = position ?? "after" };
    }

    /// <summary>Places a column-break run at the paragraph start (before) or end (after), after pPr.</summary>
    private static void InsertBreak(Paragraph paragraph, OpenXmlElement breakElement, string? position)
    {
        if (position == "before")
        {
            var pPr = paragraph.ParagraphProperties;
            if (pPr is not null)
            {
                pPr.InsertAfterSelf(breakElement);
            }
            else
            {
                paragraph.InsertAt(breakElement, 0);
            }
        }
        else
        {
            paragraph.AppendChild(breakElement);
        }
    }

    /// <summary>The sectPr that governs the target paragraph: the first one at or after it in document order.</summary>
    private static SectionProperties? GoverningSectionProperties(Body body, Paragraph target)
    {
        var seen = false;
        foreach (var child in body.ChildElements)
        {
            if (ReferenceEquals(child, target))
            {
                seen = true;
            }

            if (!seen)
            {
                continue;
            }

            if (child is Paragraph p && p.ParagraphProperties?.SectionProperties is { } sectPr)
            {
                return sectPr;
            }

            if (child is SectionProperties bodySectPr)
            {
                return bodySectPr;
            }
        }

        return body.Elements<SectionProperties>().FirstOrDefault();
    }

    /// <summary>
    /// <c>{"op":"remove","path":"/section[i]"}</c>: removing a paragraph-level
    /// break merges FORWARD — the removed section's content joins the following
    /// section and adopts ITS page setup (Word's delete-a-break behavior).
    /// Removing the final (body-level) section merges BACKWARD: the previous
    /// break's sectPr moves to the body, so the previous section's setup governs
    /// the trailing content. The single remaining section cannot be removed.
    /// </summary>
    private static object ApplyRemoveSection(WordprocessingDocument doc, EditOp op)
    {
        var path = DocPath.Parse(op.Path);
        var sectPr = ResolveSection(doc, path, createImplicit: false);
        var canonical = path.ToCanonicalString();

        if (sectPr.Parent is ParagraphProperties pPr)
        {
            sectPr.Remove();
            if (!pPr.HasChildren)
            {
                pPr.Remove();
            }

            return new
            {
                op = "remove",
                path = canonical,
                type = "section",
                merge = "forward",
                sections = Math.Max(EnumerateSections(doc).Count, 1),
            };
        }

        var sections = EnumerateSections(doc);
        if (sections.Count <= 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A document always has at least one section; /section[1] cannot be removed.",
                "Remove a section break instead (a /section[i] that ends mid-document), or change this section's setup with set.");
        }

        var body = (Body)sectPr.Parent!;
        var previous = sections[^2]; // the last paragraph-level break
        var previousPPr = (ParagraphProperties)previous.Parent!;

        sectPr.Remove();
        previous.Remove();
        if (!previousPPr.HasChildren)
        {
            previousPPr.Remove();
        }

        body.AppendChild(previous); // body-level sectPr must stay the last body child

        return new
        {
            op = "remove",
            path = canonical,
            type = "section",
            merge = "backward",
            sections = Math.Max(EnumerateSections(doc).Count, 1),
        };
    }

    // -------------------------------------------------------------- structure

    /// <summary>
    /// Sections with their block ranges for read --view structure: every
    /// paragraph-level sectPr closes a section at its paragraph; the body-level
    /// (or implicit) sectPr closes the final one at the last block.
    /// </summary>
    private static List<object> SectionsStructure(Body body)
    {
        var blocks = new List<string>();
        var breakAt = new List<(int BlockIndex, SectionProperties SectPr)>();
        int p = 0, table = 0;
        foreach (var child in body.ChildElements)
        {
            switch (child)
            {
                case Paragraph paragraph:
                    blocks.Add(SectionBlockPath("p", ++p));
                    if (paragraph.ParagraphProperties?.SectionProperties is { } sectPr)
                    {
                        breakAt.Add((blocks.Count - 1, sectPr));
                    }

                    break;

                case Table:
                    blocks.Add(SectionBlockPath("table", ++table));
                    break;

                default:
                    break;
            }
        }

        var sections = new List<object>();
        var start = 0;
        foreach (var (blockIndex, sectPr) in breakAt)
        {
            sections.Add(SectionRangeShape(sections.Count + 1, sectPr, blocks, start, blockIndex));
            start = blockIndex + 1;
        }

        var bodySectPr = body.Elements<SectionProperties>().FirstOrDefault();
        sections.Add(SectionRangeShape(sections.Count + 1, bodySectPr, blocks, start, blocks.Count - 1));
        return sections;
    }

    private static string SectionBlockPath(string name, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"/body/{name}[{index}]");

    private static object SectionRangeShape(int index, SectionProperties? sectPr, List<string> blocks, int from, int to) => new
    {
        path = SectionPath(index),
        kind = SectionKindName(sectPr),
        start = from <= to && from < blocks.Count ? blocks[from] : null,
        end = to >= from && to >= 0 && to < blocks.Count ? blocks[to] : null,
        blocks = Math.Max(0, to - from + 1),
    };

    /// <summary>How the section begins; absent w:type means nextPage.</summary>
    private static string SectionKindName(SectionProperties? sectPr)
    {
        var value = sectPr?.GetFirstChild<SectionType>()?.Val?.Value;
        if (value is null)
        {
            return "nextPage";
        }

        if (value == SectionMarkValues.Continuous)
        {
            return "continuous";
        }

        if (value == SectionMarkValues.EvenPage)
        {
            return "evenPage";
        }

        if (value == SectionMarkValues.OddPage)
        {
            return "oddPage";
        }

        if (value == SectionMarkValues.NextColumn)
        {
            return "nextColumn";
        }

        return "nextPage";
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
        var cols = sectPr?.GetFirstChild<Columns>();

        var width = pgSz?.Width?.Value;
        var height = pgSz?.Height?.Value;
        var landscape = pgSz?.Orient?.Value == PageOrientationValues.Landscape;
        var sizeName = MatchPageSizeName(width, height);

        return new Dictionary<string, object?>
        {
            ["index"] = index,
            ["sections"] = Math.Max(sections.Count, 1),
            ["kind"] = SectionKindName(sectPr),
            ["orientation"] = landscape ? "landscape" : "portrait",
            ["pageSize"] = sizeName,
            ["widthCm"] = width is { } w ? TwipsToCm(w) : null,
            ["heightCm"] = height is { } h ? TwipsToCm(h) : null,
            ["marginTopCm"] = pgMar?.Top?.Value is { } t ? TwipsToCm(t) : null,
            ["marginBottomCm"] = pgMar?.Bottom?.Value is { } b ? TwipsToCm(b) : null,
            ["marginLeftCm"] = pgMar?.Left?.Value is { } l ? TwipsToCm((long)l) : null,
            ["marginRightCm"] = pgMar?.Right?.Value is { } r ? TwipsToCm((long)r) : null,
            ["columns"] = ColumnCountOf(cols),
            ["columnGapCm"] = ColumnGapCm(cols),
            ["columnWidthsCm"] = ColumnWidthsCm(cols),
            ["pageBorder"] = PageBorderShape(sectPr),
        };
    }

    /// <summary>The number of columns the section declares (1 when absent — single column).</summary>
    private static int ColumnCountOf(Columns? cols)
    {
        if (cols is null)
        {
            return 1;
        }

        if (cols.ColumnCount?.Value is { } n && n > 0)
        {
            return n;
        }

        var explicitCols = cols.Elements<Column>().Count();
        return explicitCols > 0 ? explicitCols : 1;
    }

    private static double? ColumnGapCm(Columns? cols) =>
        cols?.Space?.Value is { } space &&
        long.TryParse(space, NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips)
            ? TwipsToCm(twips)
            : null;

    private static List<double>? ColumnWidthsCm(Columns? cols)
    {
        var explicitCols = cols?.Elements<Column>().ToList();
        if (explicitCols is not { Count: > 0 })
        {
            return null; // equal-width columns have no per-column widths
        }

        return [.. explicitCols.Select(c =>
            c.Width?.Value is { } w && long.TryParse(w, NumberStyles.Integer, CultureInfo.InvariantCulture, out var t)
                ? TwipsToCm(t)
                : 0.0)];
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
