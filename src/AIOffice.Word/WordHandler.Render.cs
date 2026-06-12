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
                content = to == "html" ? RenderHtmlNode(doc, node) : RenderText(node.Element);
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
            AppendHeaderFooter(sb, header, OwningPart(doc, header.Element));
        }

        AppendContainer(sb, GetBody(doc, file), "/body", doc.MainDocumentPart);

        foreach (var footer in roots.Where(r => r.Type == "footer"))
        {
            AppendHeaderFooter(sb, footer, OwningPart(doc, footer.Element));
        }

        return sb.ToString().TrimEnd('\n');
    }

    /// <summary>Scoped render of one resolved node, blocks tagged with canonical paths.</summary>
    internal static string RenderHtmlNode(WordprocessingDocument doc, ResolvedNode node)
    {
        var sb = new StringBuilder();
        var part = OwningPart(doc, node.Element);
        switch (node.Element)
        {
            case Paragraph p:
                AppendBlock(sb, p, listOpen: false, node.CanonicalPath, part);
                break;

            case Table t:
                AppendTable(sb, t, node.CanonicalPath, part);
                break;

            case Header or Footer:
                AppendHeaderFooter(sb, node, part);
                break;

            default:
                AppendContainer(sb, node.Element, node.CanonicalPath, part);
                break;
        }

        return sb.ToString().TrimEnd('\n');
    }

    /// <summary>The part whose relationships resolve r:embed ids under this element.</summary>
    private static OpenXmlPart? OwningPart(WordprocessingDocument doc, OpenXmlElement element)
    {
        var root = element;
        while (root.Parent is { } parent)
        {
            root = parent;
        }

        return root switch
        {
            Header header => doc.MainDocumentPart?.HeaderParts.FirstOrDefault(p => ReferenceEquals(p.Header, header)),
            Footer footer => doc.MainDocumentPart?.FooterParts.FirstOrDefault(p => ReferenceEquals(p.Footer, footer)),
            _ => doc.MainDocumentPart,
        };
    }

    /// <summary>The html tag for a header/footer root is its type name, path on the wrapper.</summary>
    private static void AppendHeaderFooter(StringBuilder sb, ResolvedNode root, OpenXmlPart? part)
    {
        var tag = root.Type; // "header" | "footer"
        sb.Append('<').Append(tag).Append(" data-aio-path=\"").Append(root.CanonicalPath).Append("\">\n");
        AppendContainer(sb, root.Element, root.CanonicalPath, part);
        sb.Append("</").Append(tag).Append(">\n");
    }

    private static void AppendContainer(StringBuilder sb, OpenXmlElement container, string path, OpenXmlPart? part)
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

                    AppendBlock(sb, p, listOpen, $"{path}/p[{pIndex}]", part);
                    break;

                case Table t:
                    tableIndex++;
                    if (listOpen)
                    {
                        sb.Append("</ul>\n");
                        listOpen = false;
                    }

                    AppendTable(sb, t, $"{path}/table[{tableIndex}]", part);
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

    private static void AppendBlock(StringBuilder sb, Paragraph p, bool listOpen, string path, OpenXmlPart? part)
    {
        var style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var level = HeadingLevel(style);
        var tag = listOpen
            ? "li"
            : level is { } h
                ? "h" + Math.Min(h, 6)
                : "p";

        sb.Append('<').Append(tag).Append(" data-aio-path=\"").Append(path).Append("\">");
        AppendInline(sb, p, path, part);
        sb.Append("</").Append(tag).Append(">\n");
    }

    /// <summary>
    /// Inline content with tracked changes rendered at their end state:
    /// w:ins content shows, w:del content is hidden. Inline images become
    /// data-URI &lt;img&gt; tags carrying the run's canonical path.
    /// </summary>
    private static void AppendInline(StringBuilder sb, Paragraph p, string path, OpenXmlPart? part)
    {
        var runIndex = 0;
        foreach (var child in p.ChildElements)
        {
            switch (child)
            {
                case Run run:
                    runIndex++;
                    AppendRun(sb, run, $"{path}/run[{runIndex}]", part);
                    break;

                case InsertedRun ins: // pending insertions are document content
                    foreach (var run in ins.ChildElements.OfType<Run>())
                    {
                        AppendRun(sb, run, path, part);
                    }

                    break;

                default: // DeletedRun, comment markers, pPr, …
                    break;
            }
        }
    }

    private static void AppendRun(StringBuilder sb, Run run, string pathForImages, OpenXmlPart? part)
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
                case Drawing drawing:
                    AppendImage(sb, drawing, pathForImages, part);
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

    /// <summary>Embed budget for data URIs; bigger images render as a sized placeholder.</summary>
    private const int MaxInlineImageBytes = 512 * 1024;

    private static void AppendImage(StringBuilder sb, Drawing drawing, string path, OpenXmlPart? part)
    {
        var extent = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().FirstOrDefault();
        var widthPx = extent?.Cx?.Value is { } cx ? (int)Math.Round(cx / 9_525.0) : 0;
        var heightPx = extent?.Cy?.Value is { } cy ? (int)Math.Round(cy / 9_525.0) : 0;

        sb.Append("<img data-aio-path=\"").Append(path).Append('"');
        if (widthPx > 0 && heightPx > 0)
        {
            sb.Append(" width=\"").Append(widthPx).Append("\" height=\"").Append(heightPx).Append('"');
        }

        sb.Append(" alt=\"image\"");

        var embedId = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault()?.Embed?.Value;
        if (embedId is not null && part?.GetPartById(embedId) is ImagePart imagePart)
        {
            using var stream = imagePart.GetStream();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            if (buffer.Length <= MaxInlineImageBytes)
            {
                sb.Append(" src=\"data:").Append(imagePart.ContentType).Append(";base64,")
                  .Append(Convert.ToBase64String(buffer.ToArray())).Append('"');
            }
        }

        sb.Append("/>");
    }

    private static void AppendTable(StringBuilder sb, Table table, string path, OpenXmlPart? part)
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
                var cellPath = $"{rowPath}/tc[{cellIndex}]";
                sb.Append("<td data-aio-path=\"").Append(cellPath).Append("\">");
                var pIndex = 0;
                foreach (var p in cell.ChildElements.OfType<Paragraph>())
                {
                    pIndex++;
                    if (pIndex > 1)
                    {
                        sb.Append("<br/>");
                    }

                    AppendInline(sb, p, $"{cellPath}/p[{pIndex}]", part);
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
