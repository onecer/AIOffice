using System.Text;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using B = DocumentFormat.OpenXml.Bibliography;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    private static readonly string[] ReadViews =
        ["text", "outline", "stats", "structure", "revisions", "comments", "styles", "markdown", "properties", "fields", "embeds", "sources"];

    public Envelope Read(CommandContext ctx)
    {
        var file = RequireFile(ctx, mustExist: true);
        var view = StringArg(ctx.Args, "view") ?? "text";
        if (!ReadViews.Contains(view, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown view '{view}'.",
                "Use --view text, outline, stats, structure, revisions, comments, styles, markdown, properties, fields, embeds or sources.",
                candidates: ReadViews);
        }

        var (doc, ms, bytes) = OpenCopy(file, editable: false);
        using (doc)
        using (ms)
        {
            var body = GetBody(doc, file);
            var warnings = new List<Warning>();

            object data = view switch
            {
                "text" => TextView(doc, body, ctx.Args, warnings),
                "outline" => OutlineView(body),
                "stats" => StatsView(body),
                "revisions" => RevisionsView(doc),
                "comments" => CommentsView(doc),
                "styles" => StylesView(doc),
                "markdown" => MarkdownView(doc, body),
                "properties" => new { view = "properties", properties = PropertiesShape(doc) },
                "fields" => FieldsView(doc),
                "embeds" => EmbedsView(doc),
                "sources" => SourcesView(doc),
                _ => StructureView(doc, body, IntArg(ctx.Args, "depth") ?? 3),
            };

            return Envelope.Ok(data, MetaFor(file, Rev.OfBytes(bytes), warnings));
        }
    }

    // ----------------------------------------------------------------- text

    /// <summary>
    /// Paragraph-per-line with canonical path prefixes; honors --range a..b and
    /// --max-bytes. List items carry their "• " / "1. " markers, footnote
    /// references show as [^n] and the notes themselves follow as a footnote
    /// section.
    /// </summary>
    private static object TextView(
        DocumentFormat.OpenXml.Packaging.WordprocessingDocument doc,
        Body body,
        System.Text.Json.Nodes.JsonObject args,
        List<Warning> warnings)
    {
        var paragraphs = WordAddress.EnumerateBody(body)
            .Where(n => n.Type == "p")
            .Select(n => new { path = n.CanonicalPath, text = DisplayText(doc, n.Element) })
            .ToList();

        var footnotes = EnumerateFootnotes(doc)
            .Select(f => new { path = FootnotePath(f.Id), id = f.Id, text = FootnoteText(f.Footnote) })
            .ToList();
        var footnoteSection = footnotes.Count > 0 ? footnotes : null;

        var endnotes = EnumerateEndnotes(doc)
            .Select(e => new { path = EndnotePath(e.Id), id = e.Id, text = EndnoteText(e.Endnote) })
            .ToList();
        var endnoteSection = endnotes.Count > 0 ? endnotes : null;

        var total = paragraphs.Count;
        var (from, to) = ParseRange(StringArg(args, "range"), total);
        var window = paragraphs.Skip(from - 1).Take(to - from + 1).ToList();

        if (IntArg(args, "maxBytes") is { } maxBytes)
        {
            var budget = maxBytes;
            var kept = new List<object>();
            foreach (var line in window)
            {
                budget -= Encoding.UTF8.GetByteCount(line.text) + Encoding.UTF8.GetByteCount(line.path) + 1;
                if (budget < 0)
                {
                    warnings.Add(new Warning(
                        WarningCodes.ResultTruncated,
                        $"Output truncated at --max-bytes {maxBytes}; {total} paragraph(s) exist. Use --range to page."));
                    break;
                }

                kept.Add(line);
            }

            return new { view = "text", totalParagraphs = total, range = $"{from}..{to}", lines = kept, footnotes = footnoteSection, endnotes = endnoteSection };
        }

        return new { view = "text", totalParagraphs = total, range = $"{from}..{to}", lines = window, footnotes = footnoteSection, endnotes = endnoteSection };
    }

    private static (int From, int To) ParseRange(string? range, int total)
    {
        if (total == 0)
        {
            return (1, 0);
        }

        if (range is null)
        {
            return (1, total);
        }

        var m = Regex.Match(range, @"^([0-9]+)(?:\.\.([0-9]+))?$");
        if (!m.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Range '{range}' is not valid.",
                "Use --range a..b with 1-based inclusive indices, e.g. --range 2..5, or a single index like --range 3.");
        }

        var from = int.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        var to = m.Groups[2].Success ? int.Parse(m.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture) : from;
        if (from < 1 || to < from)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Range '{range}' must be 1-based and ascending.",
                "Use --range a..b where 1 <= a <= b.");
        }

        return (from, Math.Min(to, total));
    }

    // -------------------------------------------------------------- outline

    private sealed record OutlineNode(string Path, int Level, string Text)
    {
        public List<OutlineNode> Children { get; } = [];
    }

    /// <summary>Headings (style Heading1..9) as a nested tree.</summary>
    private static object OutlineView(Body body)
    {
        var roots = new List<OutlineNode>();
        var stack = new Stack<OutlineNode>();

        foreach (var node in WordAddress.EnumerateBody(body).Where(n => n.Type == "p"))
        {
            var style = (node.Element as Paragraph)?.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (HeadingLevel(style) is not { } level)
            {
                continue;
            }

            var item = new OutlineNode(node.CanonicalPath, level, node.Element.InnerText);
            while (stack.Count > 0 && stack.Peek().Level >= level)
            {
                stack.Pop();
            }

            if (stack.Count == 0)
            {
                roots.Add(item);
            }
            else
            {
                stack.Peek().Children.Add(item);
            }

            stack.Push(item);
        }

        return new { view = "outline", headings = roots.Select(ToShape).ToList() };

        static object ToShape(OutlineNode n) => new
        {
            path = n.Path,
            level = n.Level,
            text = n.Text,
            children = n.Children.Select(ToShape).ToList(),
        };
    }

    internal static int? HeadingLevel(string? styleId) =>
        styleId is not null &&
        styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) &&
        int.TryParse(styleId.AsSpan("Heading".Length), out var level) &&
        level is >= 1 and <= 9
            ? level
            : null;

    // ---------------------------------------------------------------- stats

    private static object StatsView(Body body)
    {
        var nodes = WordAddress.EnumerateBody(body).ToList();
        var paragraphs = nodes.Where(n => n.Type == "p").ToList();
        var text = string.Join('\n', paragraphs.Select(p => p.Element.InnerText));

        return new
        {
            view = "stats",
            paragraphs = paragraphs.Count,
            words = text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length,
            characters = text.Replace("\n", string.Empty, StringComparison.Ordinal).Length,
            tables = nodes.Count(n => n.Type == "table"),
            headings = paragraphs.Count(p =>
                HeadingLevel((p.Element as Paragraph)?.ParagraphProperties?.ParagraphStyleId?.Val?.Value) is not null),
        };
    }

    // ------------------------------------------------------------ structure

    /// <summary>Depth-limited element tree of addressable nodes; headers/footers listed beside the body.</summary>
    private static object StructureView(DocumentFormat.OpenXml.Packaging.WordprocessingDocument doc, Body body, int maxDepth)
    {
        var headers = WordAddress.HeaderFooterRoots(doc)
            .Where(r => r.Type == "header")
            .Select(DescribeHeaderFooter)
            .ToList();
        var footers = WordAddress.HeaderFooterRoots(doc)
            .Where(r => r.Type == "footer")
            .Select(DescribeHeaderFooter)
            .ToList();

        var bookmarks = EnumerateBookmarks(doc)
            .Select(b => BookmarkShape(doc, b))
            .ToList();

        var tocs = EnumerateTocs(doc)
            .Select((sdt, i) => new { path = $"/toc[{i + 1}]", properties = TocShape(sdt) })
            .ToList();

        var captions = CaptionsStructure(doc);

        var tablesOfFigures = TablesOfFiguresStructure(doc);

        var indexes = IndexesStructure(doc);

        var mergeFields = MergeFieldsView(doc);

        var ifFields = IfFieldsView(doc);

        var embeds = EnumerateEmbeds(doc)
            .Select(e => new
            {
                path = e.CanonicalPath,
                name = e.DisplayName,
                mediaType = e.MediaType,
                size = e.Size,
                container = e.Container,
            })
            .ToList();

        var sources = (ReadSourcesRoot(doc)?.Elements<B.Source>() ?? [])
            .Select(s => new { path = SourcePath(TagOf(s) ?? string.Empty), tag = TagOf(s) })
            .ToList();

        var bibliographies = EnumerateBibliographies(doc)
            .Select((_, i) => new { path = $"/bibliography[{i + 1}]" })
            .ToList();

        // v1.3.0 additions: body drawing shapes, text boxes, legacy form fields
        // and the document theme colors.
        var shapes = BodyShapesStructure(doc, textBoxes: false);
        var textBoxes = BodyShapesStructure(doc, textBoxes: true);
        var formFields = FormFieldsView(doc);
        var theme = doc.MainDocumentPart?.ThemePart is not null ? GetThemeProperties(doc) : null;

        return new
        {
            view = "structure",
            root = Describe(body, "/body", "body", 0),
            headers = headers.Count > 0 ? headers : null,
            footers = footers.Count > 0 ? footers : null,
            bookmarks = bookmarks.Count > 0 ? bookmarks : null,
            tocs = tocs.Count > 0 ? tocs : null,
            captions = captions.Count > 0 ? captions : null,
            tablesOfFigures = tablesOfFigures.Count > 0 ? tablesOfFigures : null,
            indexes = indexes.Count > 0 ? indexes : null,
            mergeFields = mergeFields.Count > 0 ? mergeFields : null,
            ifFields = ifFields.Count > 0 ? ifFields : null,
            embeds = embeds.Count > 0 ? embeds : null,
            sources = sources.Count > 0 ? sources : null,
            bibliographies = bibliographies.Count > 0 ? bibliographies : null,
            shapes = shapes.Count > 0 ? shapes : null,
            textBoxes = textBoxes.Count > 0 ? textBoxes : null,
            formFields = formFields.Count > 0 ? formFields : null,
            theme,
            sections = SectionsStructure(body),
        };

        object Describe(OpenXmlElement element, string path, string type, int depth)
        {
            var children = DirectAddressableChildren(element, path);
            var snippet = element.InnerText;
            return new
            {
                type,
                path,
                text = snippet.Length > 60 ? snippet[..60] + "…" : (snippet.Length > 0 ? snippet : null),
                childCount = children.Count,
                children = depth < maxDepth
                    ? children.Select(c => Describe(c.Element, c.CanonicalPath, c.Type, depth + 1)).ToList()
                    : null,
            };
        }

        // Header/footer roots additionally carry their variant (default | firstPage | even).
        object DescribeHeaderFooter(ResolvedNode root)
        {
            var part = OwningPart(doc, root.Element);
            var children = DirectAddressableChildren(root.Element, root.CanonicalPath);
            var snippet = root.Element.InnerText;
            return new
            {
                type = root.Type,
                path = root.CanonicalPath,
                variant = part is null ? null : WordAddress.HeaderFooterVariantOf(doc, part) ?? "unreferenced",
                text = snippet.Length > 60 ? snippet[..60] + "…" : (snippet.Length > 0 ? snippet : null),
                childCount = children.Count,
                children = children.Select(c => Describe(c.Element, c.CanonicalPath, c.Type, 1)).ToList(),
            };
        }
    }

    private static List<ResolvedNode> DirectAddressableChildren(OpenXmlElement element, string path)
    {
        var result = new List<ResolvedNode>();
        int p = 0, table = 0, run = 0, link = 0, tr = 0, tc = 0;
        foreach (var child in element.ChildElements)
        {
            switch (child)
            {
                case Paragraph para:
                    result.Add(new ResolvedNode(para, $"{path}/p[{++p}]", "p"));
                    break;
                case Table t:
                    result.Add(new ResolvedNode(t, $"{path}/table[{++table}]", "table"));
                    break;
                case Run r:
                    result.Add(new ResolvedNode(r, $"{path}/run[{++run}]", "run"));
                    break;
                case Hyperlink h:
                    result.Add(new ResolvedNode(h, $"{path}/link[{++link}]", "link"));
                    break;
                case TableRow row:
                    result.Add(new ResolvedNode(row, $"{path}/tr[{++tr}]", "tr"));
                    break;
                case TableCell cell:
                    result.Add(new ResolvedNode(cell, $"{path}/tc[{++tc}]", "tc"));
                    break;
                default:
                    break;
            }
        }

        return result;
    }
}
