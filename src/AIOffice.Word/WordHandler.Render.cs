using System.Text;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    public Envelope Render(CommandContext ctx)
    {
        var file = RequireFile(ctx, mustExist: true);
        var to = StringArg(ctx.Args, "to") ?? "html";

        if (to is not ("html" or "text"))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"docx render --to {to} is not implemented yet.",
                "Use --to html (semantic HTML with data-aio-path attributes) or --to text.",
                candidates: ["html", "text"]);
        }

        var (doc, ms, bytes) = OpenCopy(file, editable: false);
        using (doc)
        using (ms)
        {
            string? scopePath = StringArg(ctx.Args, "scope");
            string content;
            if (scopePath is not null)
            {
                var node = WordAddress.Resolve(doc, DocPath.Parse(scopePath));
                content = to == "html" ? RenderHtmlNode(node) : RenderText(node.Element);
            }
            else
            {
                content = to == "html" ? RenderHtmlDocument(doc, file) : RenderText(GetBody(doc, file));
            }

            string? outFile = null;
            if (StringArg(ctx.Args, "output") is { } outArg)
            {
                outFile = ctx.Workspace.Resolve(outArg);
                var dir = Path.GetDirectoryName(outFile);
                if (dir is { Length: > 0 })
                {
                    Directory.CreateDirectory(dir);
                }

                File.WriteAllText(outFile, content);
            }

            return Envelope.Ok(
                new { format = to, scope = scopePath, content, written = outFile },
                MetaFor(file, Rev.OfBytes(bytes)));
        }
    }

    // ----------------------------------------------------------------- html

    /// <summary>
    /// Clean semantic HTML under the data-aio-path render contract: every block
    /// element that maps to an addressable node (p/headings/li, table, tr, td,
    /// header/footer wrappers) carries its canonical document path, so a browser
    /// click maps straight back to an address. Headers render first inside
    /// &lt;header&gt; tags, footers last inside &lt;footer&gt; tags.
    /// </summary>
    internal static string RenderHtmlDocument(WordprocessingDocument doc, string file)
    {
        var sb = new StringBuilder();
        var roots = WordAddress.HeaderFooterRoots(doc).ToList();

        foreach (var header in roots.Where(r => r.Type == "header"))
        {
            AppendHeaderFooter(sb, header);
        }

        AppendContainer(sb, GetBody(doc, file), "/body");

        foreach (var footer in roots.Where(r => r.Type == "footer"))
        {
            AppendHeaderFooter(sb, footer);
        }

        return sb.ToString().TrimEnd('\n');
    }

    /// <summary>Scoped render of one resolved node, blocks tagged with canonical paths.</summary>
    internal static string RenderHtmlNode(ResolvedNode node)
    {
        var sb = new StringBuilder();
        switch (node.Element)
        {
            case Paragraph p:
                AppendBlock(sb, p, listOpen: false, node.CanonicalPath);
                break;

            case Table t:
                AppendTable(sb, t, node.CanonicalPath);
                break;

            case Header or Footer:
                AppendHeaderFooter(sb, node);
                break;

            default:
                AppendContainer(sb, node.Element, node.CanonicalPath);
                break;
        }

        return sb.ToString().TrimEnd('\n');
    }

    /// <summary>The html tag for a header/footer root is its type name, path on the wrapper.</summary>
    private static void AppendHeaderFooter(StringBuilder sb, ResolvedNode root)
    {
        var tag = root.Type; // "header" | "footer"
        sb.Append('<').Append(tag).Append(" data-aio-path=\"").Append(root.CanonicalPath).Append("\">\n");
        AppendContainer(sb, root.Element, root.CanonicalPath);
        sb.Append("</").Append(tag).Append(">\n");
    }

    private static void AppendContainer(StringBuilder sb, OpenXmlElement container, string path)
    {
        var listOpen = false;
        var pIndex = 0;
        var tableIndex = 0;
        foreach (var child in container.ChildElements)
        {
            switch (child)
            {
                case Paragraph p:
                    pIndex++;
                    var isListItem = p.ParagraphProperties?.NumberingProperties is not null;
                    if (isListItem && !listOpen)
                    {
                        sb.Append("<ul>\n");
                        listOpen = true;
                    }
                    else if (!isListItem && listOpen)
                    {
                        sb.Append("</ul>\n");
                        listOpen = false;
                    }

                    AppendBlock(sb, p, listOpen, $"{path}/p[{pIndex}]");
                    break;

                case Table t:
                    tableIndex++;
                    if (listOpen)
                    {
                        sb.Append("</ul>\n");
                        listOpen = false;
                    }

                    AppendTable(sb, t, $"{path}/table[{tableIndex}]");
                    break;

                default:
                    break;
            }
        }

        if (listOpen)
        {
            sb.Append("</ul>\n");
        }
    }

    private static void AppendBlock(StringBuilder sb, Paragraph p, bool listOpen, string path)
    {
        var style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var level = HeadingLevel(style);
        var tag = listOpen
            ? "li"
            : level is { } h
                ? "h" + Math.Min(h, 6)
                : "p";

        sb.Append('<').Append(tag).Append(" data-aio-path=\"").Append(path).Append("\">");
        AppendInline(sb, p);
        sb.Append("</").Append(tag).Append(">\n");
    }

    private static void AppendInline(StringBuilder sb, Paragraph p)
    {
        foreach (var run in p.ChildElements.OfType<Run>())
        {
            var open = new List<string>(3);
            if (WordFormatting.IsOn(run.RunProperties?.Bold) == true)
            {
                open.Add("strong");
            }

            if (WordFormatting.IsOn(run.RunProperties?.Italic) == true)
            {
                open.Add("em");
            }

            if (WordFormatting.IsUnderlined(run.RunProperties) == true)
            {
                open.Add("u");
            }

            foreach (var tag in open)
            {
                sb.Append('<').Append(tag).Append('>');
            }

            foreach (var piece in run.ChildElements)
            {
                switch (piece)
                {
                    case Text t:
                        sb.Append(Escape(t.Text));
                        break;
                    case Break:
                        sb.Append("<br/>");
                        break;
                    case TabChar:
                        sb.Append('\t');
                        break;
                    default:
                        break;
                }
            }

            for (var i = open.Count - 1; i >= 0; i--)
            {
                sb.Append("</").Append(open[i]).Append('>');
            }
        }
    }

    private static void AppendTable(StringBuilder sb, Table table, string path)
    {
        sb.Append("<table data-aio-path=\"").Append(path).Append("\">\n");
        var rowIndex = 0;
        foreach (var row in table.ChildElements.OfType<TableRow>())
        {
            rowIndex++;
            var rowPath = $"{path}/tr[{rowIndex}]";
            sb.Append("<tr data-aio-path=\"").Append(rowPath).Append("\">");
            var cellIndex = 0;
            foreach (var cell in row.ChildElements.OfType<TableCell>())
            {
                cellIndex++;
                sb.Append("<td data-aio-path=\"").Append(rowPath).Append("/tc[").Append(cellIndex).Append("]\">");
                var first = true;
                foreach (var p in cell.ChildElements.OfType<Paragraph>())
                {
                    if (!first)
                    {
                        sb.Append("<br/>");
                    }

                    AppendInline(sb, p);
                    first = false;
                }

                sb.Append("</td>");
            }

            sb.Append("</tr>\n");
        }

        sb.Append("</table>\n");
    }

    private static string Escape(string text) => text
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal);

    // ----------------------------------------------------------------- text

    private static string RenderText(OpenXmlElement root)
    {
        var sb = new StringBuilder();
        Append(root);
        return sb.ToString().TrimEnd('\n');

        void Append(OpenXmlElement element)
        {
            switch (element)
            {
                case Paragraph p:
                    sb.Append(p.InnerText).Append('\n');
                    break;

                case Table t:
                    foreach (var row in t.ChildElements.OfType<TableRow>())
                    {
                        sb.Append(string.Join('\t', row.ChildElements.OfType<TableCell>().Select(c => c.InnerText)))
                          .Append('\n');
                    }

                    break;

                case Run r:
                    sb.Append(r.InnerText).Append('\n');
                    break;

                default:
                    foreach (var child in element.ChildElements)
                    {
                        Append(child);
                    }

                    break;
            }
        }
    }
}
