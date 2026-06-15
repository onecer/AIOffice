using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>Fonts that mark a run as code when exporting to markdown.</summary>
    private static readonly string[] MonospaceFonts =
        ["Consolas", "Courier New", "Courier", "Menlo", "Monaco", "SF Mono", "Cascadia Code"];

    /// <summary>
    /// The markdown bridge, export side: <c>read --view markdown</c> walks the
    /// body and emits clean GFM — headings by style, bold/italic/strike, nested
    /// lists with real numbering, pipe tables (pipes escaped), hyperlinks,
    /// fenced code blocks, blockquotes, horizontal rules, [^n] footnotes.
    /// Headers, footers and watermarks have no markdown shape; a leading HTML
    /// comment names what was omitted.
    /// </summary>
    private static object MarkdownView(WordprocessingDocument doc, Body body)
    {
        var sb = new StringBuilder();

        var omitted = new List<string>();
        if (doc.MainDocumentPart?.HeaderParts.Any() == true)
        {
            omitted.Add("headers");
        }

        if (doc.MainDocumentPart?.FooterParts.Any() == true)
        {
            omitted.Add("footers");
        }

        if (FindWatermarkShapes(doc).Count > 0)
        {
            omitted.Add("watermark");
        }

        if (omitted.Count > 0)
        {
            sb.Append("<!-- omitted: ").Append(string.Join(", ", omitted)).Append(" -->");
        }

        var lastWasListItem = false;
        void Emit(string block, bool listItem = false)
        {
            if (block.Length == 0)
            {
                return;
            }

            if (sb.Length > 0)
            {
                sb.Append(listItem && lastWasListItem ? "\n" : "\n\n");
            }

            sb.Append(block);
            lastWasListItem = listItem;
        }

        foreach (var child in body.ChildElements)
        {
            switch (child)
            {
                case Paragraph paragraph:
                {
                    var (block, isListItem) = ParagraphMarkdown(doc, paragraph);
                    Emit(block, isListItem);
                    break;
                }

                case Table table:
                    Emit(TableMarkdown(doc, table));
                    break;

                default:
                    break;
            }
        }

        foreach (var (footnote, id) in EnumerateFootnotes(doc))
        {
            Emit($"[^{id}]: {EscapeMd(FootnoteText(footnote), inTable: false)}");
        }

        foreach (var (endnote, id) in EnumerateEndnotes(doc))
        {
            Emit($"[^e{id}]: {EscapeMd(EndnoteText(endnote), inTable: false)}");
        }

        return new { view = "markdown", markdown = sb.ToString() };
    }

    // ---------------------------------------------------------------- blocks

    private static (string Block, bool ListItem) ParagraphMarkdown(WordprocessingDocument doc, Paragraph paragraph)
    {
        var pPr = paragraph.ParagraphProperties;
        var style = pPr?.ParagraphStyleId?.Val?.Value;

        // A numbered display equation lives inline in a tab-aligned paragraph with a
        // trailing number run; render it as block math plus its label, skipping the
        // tabs/number run so the markdown reads "$$latex$$ (1.1)".
        if (NumberedEquationInParagraph(paragraph) is { } numbered)
        {
            return ("$$" + EquationLatex(numbered) + "$$ " + ReadStoredEquationNumber(numbered), false);
        }

        // Code paragraphs (the import's "Code" style) become fenced blocks.
        if (string.Equals(style, "Code", StringComparison.Ordinal))
        {
            return ("```\n" + CodeBlockText(paragraph) + "\n```", false);
        }

        // An empty paragraph whose only feature is a bottom border is a horizontal rule.
        var borders = pPr?.ParagraphBorders;
        if (paragraph.InnerText.Length == 0 && borders?.BottomBorder is not null && borders.LeftBorder is null)
        {
            return ("---", false);
        }

        if (HeadingLevel(style) is { } level)
        {
            return (new string('#', Math.Min(level, 6)) + " " + InlineMarkdown(doc, paragraph, inTable: false), false);
        }

        var quoteDepth = QuoteDepth(pPr);
        var quotePrefix = string.Concat(Enumerable.Repeat("> ", quoteDepth));

        if (ListInfoOf(doc, paragraph) is { } info)
        {
            // 4 spaces per level: inside list context that nests for both bullet
            // ("- ", content col 2) and ordered ("1. ", content col 3) parents
            // without crossing the +4 indented-code threshold.
            var indent = new string(' ', info.Level * 4);
            var marker = info.Kind == "bullet"
                ? "- "
                : string.Create(CultureInfo.InvariantCulture, $"{ComputeListNumber(doc, paragraph, info)}. ");
            return (quotePrefix + indent + marker + InlineMarkdown(doc, paragraph, inTable: false), true);
        }

        var text = InlineMarkdown(doc, paragraph, inTable: false);
        if (text.Trim().Length == 0)
        {
            return (string.Empty, false);
        }

        return (quotePrefix + EscapeLineStart(text), false);
    }

    private static int QuoteDepth(ParagraphProperties? pPr)
    {
        if (pPr?.ParagraphBorders?.LeftBorder is null)
        {
            return 0;
        }

        var leftTwips = pPr.Indentation?.Left?.Value is { } left &&
            int.TryParse(left, NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips)
                ? twips
                : 360;
        return Math.Max(1, leftTwips / 360);
    }

    private static string CodeBlockText(Paragraph paragraph)
    {
        var sb = new StringBuilder();
        foreach (var piece in paragraph.Descendants())
        {
            switch (piece)
            {
                case Text text:
                    sb.Append(text.Text);
                    break;

                case Break:
                    sb.Append('\n');
                    break;

                default:
                    break;
            }
        }

        return sb.ToString();
    }

    /// <summary>Pipe table: first row is the GFM header, alignments read off its cell paragraphs.</summary>
    private static string TableMarkdown(WordprocessingDocument doc, Table table)
    {
        var rows = table.ChildElements.OfType<TableRow>().ToList();
        if (rows.Count == 0)
        {
            return string.Empty;
        }

        var columns = Math.Max(1, GridColumnCount(table));
        var lines = new List<string>();
        string?[] alignments = new string?[columns];

        for (var r = 0; r < rows.Count; r++)
        {
            var cells = new List<string>();
            foreach (var cell in rows[r].ChildElements.OfType<TableCell>())
            {
                var span = CellSpan(cell);
                var content = VerticalMergeState(cell) == "continue"
                    ? string.Empty // GFM has no rowspan; continuation slots read as empty
                    : string.Join("<br>", cell.ChildElements.OfType<Paragraph>()
                        // The GFM header row is implicitly bold; re-emitting ** would be noise.
                        .Select(p => InlineMarkdown(doc, p, inTable: true, suppressBold: r == 0))
                        .Where(t => t.Length > 0));

                if (r == 0)
                {
                    var alignment = WordFormatting.AlignmentName(
                        cell.ChildElements.OfType<Paragraph>().FirstOrDefault()?.ParagraphProperties?.Justification);
                    for (var c = cells.Count; c < Math.Min(columns, cells.Count + span); c++)
                    {
                        alignments[c] = alignment;
                    }
                }

                cells.Add(content);
                for (var s = 1; s < span; s++)
                {
                    cells.Add(string.Empty); // GFM has no colspan either; keep the grid width
                }
            }

            while (cells.Count < columns)
            {
                cells.Add(string.Empty);
            }

            lines.Add("| " + string.Join(" | ", cells.Take(columns)) + " |");

            if (r == 0)
            {
                lines.Add("| " + string.Join(" | ", alignments.Select(a => a switch
                {
                    "center" => ":---:",
                    "right" => "---:",
                    "left" => ":---",
                    _ => "---",
                })) + " |");
            }
        }

        return string.Join('\n', lines);
    }

    // --------------------------------------------------------------- inlines

    private static string InlineMarkdown(WordprocessingDocument doc, Paragraph paragraph, bool inTable, bool suppressBold = false)
    {
        var sb = new StringBuilder();
        foreach (var child in paragraph.ChildElements)
        {
            AppendInlineMarkdown(sb, child, doc, inTable, suppressBold);
        }

        return sb.ToString();
    }

    private static void AppendInlineMarkdown(StringBuilder sb, OpenXmlElement child, WordprocessingDocument doc, bool inTable, bool suppressBold)
    {
        switch (child)
        {
            case Run run:
                sb.Append(RunMarkdown(run, inTable, suppressBold));
                break;

            case Hyperlink hyperlink:
            {
                var href = ResolveLinkUrl(doc, hyperlink)
                    ?? (hyperlink.Anchor?.Value is { Length: > 0 } anchor ? "#" + anchor : null);
                var inner = new StringBuilder();
                foreach (var run in hyperlink.ChildElements.OfType<Run>())
                {
                    inner.Append(RunMarkdown(run, inTable, suppressBold));
                }

                sb.Append(href is null
                    ? inner.ToString()
                    : $"[{inner}]({href.Replace(")", "%29", StringComparison.Ordinal)})");
                break;
            }

            case SimpleField field: // cached field result is the honest text-shape
                foreach (var run in field.ChildElements.OfType<Run>())
                {
                    sb.Append(RunMarkdown(run, inTable, suppressBold));
                }

                break;

            case InsertedRun ins: // tracked changes export at their end state
                foreach (var run in ins.ChildElements.OfType<Run>())
                {
                    sb.Append(RunMarkdown(run, inTable, suppressBold));
                }

                break;

            case DocumentFormat.OpenXml.Math.OfficeMath inlineMath: // GFM inline math
                sb.Append('$').Append(EquationLatex(inlineMath)).Append('$');
                break;

            case DocumentFormat.OpenXml.Math.Paragraph displayMath: // m:oMathPara → block math
                foreach (var oMath in displayMath.ChildElements.OfType<DocumentFormat.OpenXml.Math.OfficeMath>())
                {
                    sb.Append("$$").Append(EquationLatex(oMath)).Append("$$");
                }

                break;

            default: // DeletedRun, comment markers, pPr, bookmarks, …
                break;
        }
    }

    private static string RunMarkdown(Run run, bool inTable, bool suppressBold = false)
    {
        var rPr = run.RunProperties;
        var code = IsMonospaceRun(rPr);

        var sb = new StringBuilder();
        foreach (var piece in run.ChildElements)
        {
            switch (piece)
            {
                case Text text:
                    sb.Append(code ? text.Text.Replace("`", "ˋ", StringComparison.Ordinal) : EscapeMd(text.Text, inTable));
                    break;

                case Break:
                    sb.Append(inTable ? "<br>" : "  \n");
                    break;

                case TabChar:
                    sb.Append('\t');
                    break;

                case FootnoteReference footnoteRef:
                    sb.Append("[^").Append(footnoteRef.Id?.Value ?? 0).Append(']');
                    break;

                case EndnoteReference endnoteRef:
                    sb.Append("[^e").Append(endnoteRef.Id?.Value ?? 0).Append(']');
                    break;

                case Drawing drawing:
                {
                    var name = drawing
                        .Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
                        .FirstOrDefault()?.Name?.Value ?? "image";
                    // The bytes live in the package, not on disk; the target is a media note.
                    sb.Append("![").Append(EscapeMd(name, inTable)).Append("](media:")
                      .Append(Uri.EscapeDataString(name)).Append(')');
                    break;
                }

                default:
                    break;
            }
        }

        var content = sb.ToString();
        if (content.Trim().Length == 0)
        {
            return content;
        }

        if (code)
        {
            return WrapMd(content, "`");
        }

        if (WordFormatting.IsOn(rPr?.Strike) == true)
        {
            content = WrapMd(content, "~~");
        }

        if (!suppressBold && WordFormatting.IsOn(rPr?.Bold) == true)
        {
            content = WrapMd(content, "**");
        }

        if (WordFormatting.IsOn(rPr?.Italic) == true)
        {
            content = WrapMd(content, "*");
        }

        return content;
    }

    private static bool IsMonospaceRun(RunProperties? rPr) =>
        rPr?.RunFonts?.Ascii?.Value is { } font &&
        MonospaceFonts.Contains(font, StringComparer.OrdinalIgnoreCase);

    /// <summary>Wraps in emphasis delimiters, keeping surrounding whitespace outside (GFM requirement).</summary>
    private static string WrapMd(string text, string mark)
    {
        var trimmed = text.Trim();
        if (trimmed.Length == 0)
        {
            return text;
        }

        var leadingLength = text.Length - text.TrimStart().Length;
        var leading = text[..leadingLength];
        var trailing = text[(leadingLength + trimmed.Length)..];
        return leading + mark + trimmed + mark + trailing;
    }

    /// <summary>Backslash-escapes the characters markdown would re-interpret.</summary>
    private static string EscapeMd(string text, bool inTable)
    {
        var sb = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            switch (c)
            {
                case '\\' or '`' or '*' or '_' or '[' or ']' or '<':
                    sb.Append('\\').Append(c);
                    break;

                case '|' when inTable:
                    sb.Append("\\|");
                    break;

                default:
                    sb.Append(c);
                    break;
            }
        }

        return sb.ToString();
    }

    /// <summary>Defuses block-construct lookalikes ("# x", "- x", "1. x", "> x") at line start.</summary>
    private static string EscapeLineStart(string text)
    {
        if (text.Length == 0)
        {
            return text;
        }

        if (text[0] is '#' or '>' or '-' or '+')
        {
            return "\\" + text;
        }

        var dot = text.IndexOf('.', StringComparison.Ordinal);
        if (dot > 0 && dot <= 9 && text[..dot].All(char.IsAsciiDigit))
        {
            return text[..dot] + "\\." + text[(dot + 1)..];
        }

        return text;
    }
}
