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
            AppendHeaderFooter(sb, header, doc, OwningPart(doc, header.Element));
        }

        AppendContainer(sb, GetBody(doc, file), "/body", doc, doc.MainDocumentPart);
        AppendFootnoteSection(sb, doc);
        AppendEndnoteSection(sb, doc);

        foreach (var footer in roots.Where(r => r.Type == "footer"))
        {
            AppendHeaderFooter(sb, footer, doc, OwningPart(doc, footer.Element));
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
                AppendBlock(sb, p, node.CanonicalPath, doc, part);
                break;

            case Table t:
                AppendTable(sb, t, node.CanonicalPath, doc, part);
                break;

            case Hyperlink:
            {
                var paragraphPath = node.CanonicalPath[..node.CanonicalPath.LastIndexOf('/')];
                AppendInlineChild(sb, node.Element, node.CanonicalPath, paragraphPath, doc, part);
                break;
            }

            case Header or Footer:
                AppendHeaderFooter(sb, node, doc, part);
                break;

            default:
                AppendContainer(sb, node.Element, node.CanonicalPath, doc, part);
                break;
        }

        return sb.ToString().TrimEnd('\n');
    }

    /// <summary>Footnotes render once, after the body, as an ordered list of notes.</summary>
    private static void AppendFootnoteSection(StringBuilder sb, WordprocessingDocument doc)
    {
        var footnotes = EnumerateFootnotes(doc).ToList();
        if (footnotes.Count == 0)
        {
            return;
        }

        sb.Append("<section class=\"footnotes\">\n<ol>\n");
        foreach (var (footnote, id) in footnotes)
        {
            sb.Append("<li data-aio-path=\"").Append(FootnotePath(id)).Append("\">")
              .Append(Escape(FootnoteText(footnote)))
              .Append("</li>\n");
        }

        sb.Append("</ol>\n</section>\n");
    }

    /// <summary>Endnotes render once, after the body (and footnotes), as an ordered list of notes.</summary>
    private static void AppendEndnoteSection(StringBuilder sb, WordprocessingDocument doc)
    {
        var endnotes = EnumerateEndnotes(doc).ToList();
        if (endnotes.Count == 0)
        {
            return;
        }

        sb.Append("<section class=\"endnotes\">\n<ol>\n");
        foreach (var (endnote, id) in endnotes)
        {
            sb.Append("<li data-aio-path=\"").Append(EndnotePath(id)).Append("\">")
              .Append(Escape(EndnoteText(endnote)))
              .Append("</li>\n");
        }

        sb.Append("</ol>\n</section>\n");
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
    private static void AppendHeaderFooter(StringBuilder sb, ResolvedNode root, WordprocessingDocument doc, OpenXmlPart? part)
    {
        var tag = root.Type; // "header" | "footer"
        sb.Append('<').Append(tag).Append(" data-aio-path=\"").Append(root.CanonicalPath).Append("\">\n");
        AppendContainer(sb, root.Element, root.CanonicalPath, doc, part);
        sb.Append("</").Append(tag).Append(">\n");
    }

    /// <summary>
    /// Blocks in document order; consecutive list paragraphs become real
    /// &lt;ul&gt;/&lt;ol&gt; structures, nested by their numbering level
    /// (sub-lists open inside the parent &lt;li&gt;).
    /// </summary>
    private static void AppendContainer(StringBuilder sb, OpenXmlElement container, string path, WordprocessingDocument doc, OpenXmlPart? part)
    {
        var lists = new List<(string Tag, bool LiOpen)>(); // open list stack, outermost first
        var pIndex = 0;
        var tableIndex = 0;
        foreach (var child in container.ChildElements)
        {
            switch (child)
            {
                case Paragraph p:
                    pIndex++;
                    var pPath = $"{path}/p[{pIndex}]";
                    if (ListInfoOf(doc, p) is { } info)
                    {
                        AppendListItem(sb, p, pPath, info, lists, doc, part);
                    }
                    else
                    {
                        CloseLists(sb, lists, 0);
                        AppendBlock(sb, p, pPath, doc, part);
                    }

                    break;

                case Table t:
                    tableIndex++;
                    CloseLists(sb, lists, 0);
                    AppendTable(sb, t, $"{path}/table[{tableIndex}]", doc, part);
                    break;

                default:
                    break;
            }
        }

        CloseLists(sb, lists, 0);
    }

    private static void AppendListItem(
        StringBuilder sb,
        Paragraph p,
        string path,
        ListInfo info,
        List<(string Tag, bool LiOpen)> lists,
        WordprocessingDocument doc,
        OpenXmlPart? part)
    {
        var tag = info.Kind == "bullet" ? "ul" : "ol";
        var depth = info.Level + 1;

        CloseLists(sb, lists, depth);
        if (lists.Count == depth && lists[^1].Tag != tag)
        {
            CloseLists(sb, lists, depth - 1);
        }

        while (lists.Count < depth)
        {
            sb.Append('<').Append(tag).Append(">\n");
            lists.Add((tag, false));
        }

        if (lists[^1].LiOpen)
        {
            sb.Append("</li>\n");
        }

        sb.Append("<li data-aio-path=\"").Append(path).Append("\">");
        AppendInline(sb, p, path, doc, part);
        lists[^1] = (lists[^1].Tag, true);
    }

    /// <summary>
    /// Closes open lists down to the given depth. Each level's open li closes
    /// before its list; the parent li hosting a sublist stays open (it closes
    /// when its own level does).
    /// </summary>
    private static void CloseLists(StringBuilder sb, List<(string Tag, bool LiOpen)> lists, int depth)
    {
        while (lists.Count > depth)
        {
            var (tag, liOpen) = lists[^1];
            lists.RemoveAt(lists.Count - 1);
            if (liOpen)
            {
                sb.Append("</li>\n");
            }

            sb.Append("</").Append(tag).Append(">\n");
        }
    }

    private static void AppendBlock(StringBuilder sb, Paragraph p, string path, WordprocessingDocument doc, OpenXmlPart? part)
    {
        var style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var level = HeadingLevel(style);
        var tag = level is { } h ? "h" + Math.Min(h, 6) : "p";

        sb.Append('<').Append(tag).Append(" data-aio-path=\"").Append(path).Append("\">");
        AppendInline(sb, p, path, doc, part);
        sb.Append("</").Append(tag).Append(">\n");
    }

    /// <summary>
    /// Inline content with tracked changes rendered at their end state:
    /// w:ins content shows, w:del content is hidden. Inline images become
    /// data-URI &lt;img&gt; tags carrying the run's canonical path; hyperlinks
    /// become &lt;a href&gt; carrying theirs.
    /// </summary>
    private static void AppendInline(StringBuilder sb, Paragraph p, string path, WordprocessingDocument doc, OpenXmlPart? part)
    {
        var runIndex = 0;
        var linkIndex = 0;
        foreach (var child in p.ChildElements)
        {
            switch (child)
            {
                case Run run:
                    runIndex++;
                    AppendRun(sb, run, $"{path}/run[{runIndex}]", part);
                    break;

                case Hyperlink link:
                    linkIndex++;
                    AppendInlineChild(sb, link, $"{path}/link[{linkIndex}]", path, doc, part);
                    break;

                case InsertedRun ins: // pending insertions are document content
                    foreach (var run in ins.ChildElements.OfType<Run>())
                    {
                        AppendRun(sb, run, path, part);
                    }

                    break;

                case SimpleField field: // w:fldSimple renders its cached result
                    foreach (var run in field.ChildElements.OfType<Run>())
                    {
                        AppendRun(sb, run, path, part);
                    }

                    break;

                case DocumentFormat.OpenXml.Math.OfficeMath inlineMath:
                    // A numbered equation is a display equation laid out inline with a
                    // number; render its math in display style (the number run follows).
                    AppendEquation(sb, inlineMath, display: NumberedEquationInParagraph(p) is not null);
                    break;

                case DocumentFormat.OpenXml.Math.Paragraph displayMath: // m:oMathPara
                    foreach (var oMath in displayMath.ChildElements.OfType<DocumentFormat.OpenXml.Math.OfficeMath>())
                    {
                        AppendEquation(sb, oMath, display: true);
                    }

                    break;

                default: // DeletedRun, comment markers, pPr, …
                    break;
            }
        }
    }

    /// <summary>
    /// One equation as MathML (<c>&lt;math&gt;</c>), the native browser shape. The
    /// element carries the equation's canonical path via a data-aio-path wrapper
    /// span and keeps the source LaTeX in a title for accessibility/fallback.
    /// </summary>
    private static void AppendEquation(StringBuilder sb, DocumentFormat.OpenXml.Math.OfficeMath oMath, bool display)
    {
        var latex = EquationLatex(oMath);
        var mathml = Equations.OmmlToMathml.ToMathml(oMath, display);
        sb.Append("<span class=\"equation\" title=\"")
          .Append(Escape(latex).Replace("\"", "&quot;", StringComparison.Ordinal))
          .Append("\">")
          .Append(mathml)
          .Append("</span>");
    }

    /// <summary>One hyperlink as &lt;a href&gt;: external relationship url or #bookmark anchor.</summary>
    private static void AppendInlineChild(StringBuilder sb, OpenXmlElement element, string linkPath, string paragraphPath, WordprocessingDocument doc, OpenXmlPart? part)
    {
        var hyperlink = (Hyperlink)element;
        var href = ResolveLinkUrl(doc, hyperlink)
            ?? (hyperlink.Anchor?.Value is { Length: > 0 } anchor ? "#" + anchor : null);

        sb.Append("<a data-aio-path=\"").Append(linkPath).Append('"');
        if (href is not null)
        {
            sb.Append(" href=\"").Append(Escape(href).Replace("\"", "&quot;", StringComparison.Ordinal)).Append('"');
        }

        sb.Append('>');
        foreach (var run in hyperlink.ChildElements.OfType<Run>())
        {
            AppendRun(sb, run, paragraphPath, part);
        }

        sb.Append("</a>");
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
                case FootnoteReference reference:
                    sb.Append("<sup data-aio-path=\"").Append(FootnotePath((int)(reference.Id?.Value ?? 0)))
                      .Append("\">").Append(reference.Id?.Value ?? 0).Append("</sup>");
                    break;
                case EndnoteReference endnoteReference:
                    sb.Append("<sup data-aio-path=\"").Append(EndnotePath((int)(endnoteReference.Id?.Value ?? 0)))
                      .Append("\">e").Append(endnoteReference.Id?.Value ?? 0).Append("</sup>");
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

    private static void AppendTable(StringBuilder sb, Table table, string path, WordprocessingDocument doc, OpenXmlPart? part)
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
                var (colspan, rowspan) = CellSpans(cell);
                if (rowspan == 0)
                {
                    continue; // vMerge continuation slot: the restart cell above renders it
                }

                var cellPath = $"{rowPath}/tc[{cellIndex}]";
                sb.Append("<td data-aio-path=\"").Append(cellPath).Append('"');
                if (colspan > 1)
                {
                    sb.Append(" colspan=\"").Append(colspan).Append('"');
                }

                if (rowspan > 1)
                {
                    sb.Append(" rowspan=\"").Append(rowspan).Append('"');
                }

                sb.Append('>');
                var pIndex = 0;
                foreach (var p in cell.ChildElements.OfType<Paragraph>())
                {
                    pIndex++;
                    if (pIndex > 1)
                    {
                        sb.Append("<br/>");
                    }

                    AppendInline(sb, p, $"{cellPath}/p[{pIndex}]", doc, part);
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
