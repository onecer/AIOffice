using System.Text;
using AIOffice.Core;
using DocumentFormat.OpenXml;
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
                "Use --to html (semantic HTML) or --to text; png/svg rendering arrives in M1.",
                candidates: ["html", "text"]);
        }

        var (doc, ms, bytes) = OpenCopy(file, editable: false);
        using (doc)
        using (ms)
        {
            OpenXmlElement scopeRoot = GetBody(doc, file);
            string? scopePath = StringArg(ctx.Args, "scope");
            if (scopePath is not null)
            {
                scopeRoot = WordAddress.Resolve(doc, DocPath.Parse(scopePath)).Element;
            }

            var content = to == "html" ? RenderHtml(scopeRoot) : RenderText(scopeRoot);

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

    /// <summary>Clean semantic HTML: headings, paragraphs, bold/italic/underline, tables, basic lists.</summary>
    internal static string RenderHtml(OpenXmlElement root)
    {
        var sb = new StringBuilder();
        switch (root)
        {
            case Paragraph p:
                AppendBlock(sb, p, listOpen: false);
                break;

            case Table t:
                AppendTable(sb, t);
                break;

            default:
                AppendContainer(sb, root);
                break;
        }

        return sb.ToString().TrimEnd('\n');
    }

    private static void AppendContainer(StringBuilder sb, OpenXmlElement container)
    {
        var listOpen = false;
        foreach (var child in container.ChildElements)
        {
            switch (child)
            {
                case Paragraph p:
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

                    AppendBlock(sb, p, listOpen);
                    break;

                case Table t:
                    if (listOpen)
                    {
                        sb.Append("</ul>\n");
                        listOpen = false;
                    }

                    AppendTable(sb, t);
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

    private static void AppendBlock(StringBuilder sb, Paragraph p, bool listOpen)
    {
        var style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var level = HeadingLevel(style);
        var tag = listOpen
            ? "li"
            : level is { } h
                ? "h" + Math.Min(h, 6)
                : "p";

        sb.Append('<').Append(tag).Append('>');
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

    private static void AppendTable(StringBuilder sb, Table table)
    {
        sb.Append("<table>\n");
        foreach (var row in table.ChildElements.OfType<TableRow>())
        {
            sb.Append("<tr>");
            foreach (var cell in row.ChildElements.OfType<TableCell>())
            {
                sb.Append("<td>");
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
