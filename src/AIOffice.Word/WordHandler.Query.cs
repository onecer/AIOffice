using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    private static readonly string[] QueryableElements = ["p", "run", "link", "table", "tr", "tc", "*"];

    public Envelope Get(CommandContext ctx)
    {
        var file = RequireFile(ctx, mustExist: true);
        var pathArg = StringArg(ctx.Args, "path") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "get requires a path.",
            "Pass a canonical path, e.g. 'aioffice get report.docx /body/p[1]'. Run query to discover paths.");

        var (doc, ms, bytes) = OpenCopy(file, editable: false);
        using (doc)
        using (ms)
        {
            var docPath = DocPath.Parse(pathArg);
            var meta = MetaFor(file, Rev.OfBytes(bytes));

            // The id-addressed roots that live outside body content.
            switch (docPath.Segments[0].Name)
            {
                case "revision":
                {
                    var revision = ResolveRevision(doc, docPath);
                    return Envelope.Ok(
                        new { path = RevisionPath(revision.Id), type = "revision", properties = RevisionShape(revision) },
                        meta);
                }

                case "comment":
                {
                    var (comment, id) = ResolveComment(doc, docPath);
                    return Envelope.Ok(
                        new { path = CommentPath(id), type = "comment", properties = CommentShape(doc, comment, id) },
                        meta);
                }

                case "style":
                {
                    var properties = GetStyleProperties(doc, docPath);
                    return Envelope.Ok(
                        new { path = StylePath((string)properties["id"]!), type = "style", properties },
                        meta);
                }

                case "section":
                {
                    var properties = GetSectionProperties(doc, docPath);
                    return Envelope.Ok(
                        new { path = SectionPath((int)properties["index"]!), type = "section", properties },
                        meta);
                }

                case "bookmark":
                {
                    var properties = GetBookmarkProperties(doc, docPath);
                    return Envelope.Ok(
                        new { path = BookmarkPath((string)properties["name"]!), type = "bookmark", properties },
                        meta);
                }

                case "footnote":
                {
                    var properties = GetFootnoteProperties(doc, docPath);
                    return Envelope.Ok(
                        new { path = FootnotePath((int)properties["id"]!), type = "footnote", properties },
                        meta);
                }

                case "endnote":
                {
                    var properties = GetEndnoteProperties(doc, docPath);
                    return Envelope.Ok(
                        new { path = EndnotePath((int)properties["id"]!), type = "endnote", properties },
                        meta);
                }

                case "toc":
                {
                    var properties = GetTocProperties(doc, docPath);
                    return Envelope.Ok(
                        new { path = "/toc[1]", type = "toc", properties },
                        meta);
                }

                case "watermark":
                {
                    var properties = GetWatermarkProperties(doc, docPath);
                    return Envelope.Ok(
                        new { path = "/watermark[1]", type = "watermark", properties },
                        meta);
                }

                default:
                    break;
            }

            var node = WordAddress.Resolve(doc, docPath);
            return Envelope.Ok(
                new { path = node.CanonicalPath, type = node.Type, properties = NodeProperties(doc, node) },
                meta);
        }
    }

    public Envelope Query(CommandContext ctx)
    {
        var file = RequireFile(ctx, mustExist: true);
        var selectorArg = StringArg(ctx.Args, "selector") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "query requires a selector.",
            "Pass a selector like p[style=Heading1], run[bold=true] or p:contains('Q3').");

        var selector = Selector.Parse(selectorArg);
        if (!QueryableElements.Contains(selector.Element, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{selector.Element}' is not a docx element.",
                "Query one of: p, run, table, tr, tc, or * for any element.",
                candidates: QueryableElements);
        }

        var (doc, ms, bytes) = OpenCopy(file, editable: false);
        using (doc)
        using (ms)
        {
            _ = GetBody(doc, file); // a docx without a body is format_corrupt, even for header-only queries
            var matches = WordAddress.EnumerateAll(doc)
                .Where(n => selector.Element == "*" || n.Type == selector.Element)
                .Where(n => selector.Predicates.All(predicate => Matches(doc, n, predicate)))
                .Select(n => new { path = n.CanonicalPath, type = n.Type, snippet = Snippet(n.Element.InnerText) })
                .ToList();

            return Envelope.Ok(
                new { selector = selector.ToCanonicalString(), count = matches.Count, matches },
                MetaFor(file, Rev.OfBytes(bytes)));
        }
    }

    // ------------------------------------------------------------ predicates

    private static bool Matches(WordprocessingDocument doc, ResolvedNode node, SelectorPredicate predicate) => predicate switch
    {
        ContainsPredicate contains =>
            node.Element.InnerText.Contains(contains.Text, StringComparison.OrdinalIgnoreCase),
        AttributePredicate attr => MatchesAttribute(doc, node, attr),
        _ => false,
    };

    private static bool MatchesAttribute(WordprocessingDocument doc, ResolvedNode node, AttributePredicate attr)
    {
        var actual = AttributeValue(doc, node, attr.Attribute);

        switch (attr.Op)
        {
            case SelectorOperator.Equals:
                return actual is not null && string.Equals(actual, attr.Value, StringComparison.OrdinalIgnoreCase);

            case SelectorOperator.NotEquals:
                return !string.Equals(actual, attr.Value, StringComparison.OrdinalIgnoreCase);

            case SelectorOperator.ContainsText:
                return actual?.Contains(attr.Value, StringComparison.OrdinalIgnoreCase) ?? false;

            case SelectorOperator.GreaterThan:
            case SelectorOperator.GreaterOrEqual:
            case SelectorOperator.LessThan:
            case SelectorOperator.LessOrEqual:
            {
                if (attr.NumericValue is not { } wanted ||
                    actual is null ||
                    !double.TryParse(actual, NumberStyles.Float, CultureInfo.InvariantCulture, out var have))
                {
                    return false;
                }

                return attr.Op switch
                {
                    SelectorOperator.GreaterThan => have > wanted,
                    SelectorOperator.GreaterOrEqual => have >= wanted,
                    SelectorOperator.LessThan => have < wanted,
                    _ => have <= wanted,
                };
            }

            default:
                return false;
        }
    }

    /// <summary>The comparable value of a selector attribute on one node.</summary>
    private static string? AttributeValue(WordprocessingDocument doc, ResolvedNode node, string attribute)
    {
        var element = node.Element;
        var run = element as Run ?? (element as Paragraph)?.ChildElements.OfType<Run>().FirstOrDefault();

        return attribute switch
        {
            "text" => element.InnerText,
            "url" => element is Hyperlink hyperlink
                ? ResolveLinkUrl(doc, hyperlink) ?? (hyperlink.Anchor?.Value is { } anchor ? "#" + anchor : null)
                : null,
            "style" => element switch
            {
                Paragraph p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value,
                Run r => r.RunProperties?.RunStyle?.Val?.Value,
                _ => null,
            },
            "bold" => BoolString(WordFormatting.IsOn(run?.RunProperties?.Bold)),
            "italic" => BoolString(WordFormatting.IsOn(run?.RunProperties?.Italic)),
            "underline" => BoolString(WordFormatting.IsUnderlined(run?.RunProperties)),
            "fontSize" => WordFormatting.FontSizePoints(run?.RunProperties)?.ToString(CultureInfo.InvariantCulture),
            "color" => run?.RunProperties?.Color?.Val?.Value,
            "alignment" => WordFormatting.AlignmentName((element as Paragraph)?.ParagraphProperties?.Justification),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown selector attribute '{attribute}'.",
                $"Did you mean '{WordFormatting.Nearest(attribute, KnownAttributes)}'? " +
                $"Known attributes: {string.Join(", ", KnownAttributes)}.",
                candidates: KnownAttributes),
        };

        static string? BoolString(bool? value) => value switch
        {
            true => "true",
            false => "false",
            null => null,
        };
    }

    private static readonly string[] KnownAttributes =
        ["text", "url", "style", "bold", "italic", "underline", "fontSize", "color", "alignment"];

    // -------------------------------------------------------------- get data

    private static object NodeProperties(WordprocessingDocument doc, ResolvedNode node) => node.Element switch
    {
        // Inline-image carriers answer as images (dimensions from the extent).
        Paragraph ip when ip.Descendants<Drawing>().Any() => ImageProperties(ip),
        Run ir when ir.Descendants<Drawing>().Any() => ImageProperties(ir),
        Paragraph p => ParagraphPropertiesShape(doc, p),
        Run r => WordFormatting.ReadRunProps(r),
        Hyperlink => LinkProperties(doc, node),
        Table t => new Dictionary<string, object?>
        {
            ["rows"] = t.ChildElements.OfType<TableRow>().Count(),
            ["columns"] = t.ChildElements.OfType<TableRow>().FirstOrDefault()?.ChildElements.OfType<TableCell>().Count() ?? 0,
        },
        TableRow row => new Dictionary<string, object?>
        {
            ["cells"] = row.ChildElements.OfType<TableCell>().Select(c => c.InnerText).ToList(),
        },
        TableCell cell => new Dictionary<string, object?>
        {
            ["text"] = cell.InnerText,
            ["paragraphs"] = cell.ChildElements.OfType<Paragraph>().Count(),
        },
        _ => new Dictionary<string, object?> { ["text"] = node.Element.InnerText },
    };

    /// <summary>Paragraph get data: formatting props plus {list, level, number?} for list items.</summary>
    private static Dictionary<string, object?> ParagraphPropertiesShape(WordprocessingDocument doc, Paragraph p)
    {
        var properties = WordFormatting.ReadParagraphProps(p);
        if (ListInfoOf(doc, p) is { } info)
        {
            properties["list"] = info.Kind;
            properties["level"] = info.Level;
            if (info.Kind == "number")
            {
                properties["number"] = ComputeListNumber(doc, p, info);
            }
        }

        return properties;
    }

    private static string Snippet(string text) =>
        text.Length <= 80 ? text : text[..80] + "…";
}
