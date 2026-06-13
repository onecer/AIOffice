using System.Text;
using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// The command-layer serializer between the format-neutral document model
/// (<see cref="NeutralDoc"/>) and GitHub-flavoured Markdown — the bridge that
/// lets <c>convert</c> move ANY format to/from <c>.md</c> without each handler
/// needing its own markdown path. A docx/xlsx/pptx exports to <see cref="NeutralDoc"/>
/// via <see cref="INeutralConvertible"/>; this turns that into markdown text
/// (<see cref="Write"/>) and parses markdown back into a <see cref="NeutralDoc"/>
/// (<see cref="Parse"/>) the destination handler can import.
///
/// <para>The grammar is deliberately small — exactly the neutral block kinds:
/// ATX headings (<c># … ######</c>), paragraphs, bullet/ordered list items
/// (indent = two spaces per level), GFM pipe tables (a <c>---</c> separator row
/// marks <see cref="NeutralBlock.HeaderRow"/>) and <c>![alt](src)</c> images.
/// Inline runs carry <c>**bold**</c>, <c>*italic*</c>, <c>&lt;u&gt;underline&lt;/u&gt;</c>
/// and <c>[text](href)</c> links. Anything markdown cannot represent (colors,
/// nested formatting) is flattened, matching the neutral model's own lossiness.</para>
/// </summary>
public static class NeutralMarkdown
{
    // ─────────────────────────────────────────────────────────── write (model → md)

    /// <summary>Serializes a <see cref="NeutralDoc"/> to GFM markdown text (LF line endings).</summary>
    public static string Write(NeutralDoc doc)
    {
        var sb = new StringBuilder();

        void Blank()
        {
            if (sb.Length > 0)
            {
                sb.Append('\n');
            }
        }

        // The document Title is carried as a leading H1 so it survives the trip and
        // is recovered by Parse (which reads the first H1 back into Title). We emit
        // it only when it is not ALREADY the first heading block — otherwise the
        // existing leading "# Title" already carries it and a second one would
        // duplicate the title on every round trip. Keeping it an H1 (not bespoke
        // front matter) keeps Write/Parse a stable fixed point.
        if (doc.Title is { Length: > 0 } title && !TitleIsLeadingHeading(doc, title))
        {
            sb.Append("# ").Append(Inline([new NeutralRun(title)])).Append('\n');
        }

        foreach (var block in doc.Blocks)
        {
            switch (block.Kind)
            {
                case NeutralBlockKind.Heading:
                    Blank();
                    sb.Append('#', Math.Clamp(block.Level == 0 ? 1 : block.Level, 1, 6));
                    sb.Append(' ').Append(Inline(block.Runs)).Append('\n');
                    break;

                case NeutralBlockKind.Paragraph:
                    Blank();
                    sb.Append(Inline(block.Runs)).Append('\n');
                    break;

                case NeutralBlockKind.ListItem:
                {
                    // List items pack tight (one blank line before the first only).
                    if (sb.Length > 0 && !sb.ToString().EndsWith("\n\n", StringComparison.Ordinal) &&
                        !EndsWithListLine(sb))
                    {
                        sb.Append('\n');
                    }

                    sb.Append(new string(' ', Math.Clamp(block.Level, 0, 8) * 2));
                    sb.Append(block.Ordered ? "1. " : "- ");
                    sb.Append(Inline(block.Runs)).Append('\n');
                    break;
                }

                case NeutralBlockKind.Table:
                    Blank();
                    WriteTable(sb, block);
                    break;

                case NeutralBlockKind.Image:
                    Blank();
                    sb.Append("![").Append(EscapeInline(block.Alt ?? string.Empty)).Append("](")
                        .Append(block.Source ?? string.Empty).Append(")\n");
                    break;

                default:
                    break;
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// True when the doc's first block is a level-1 heading whose text already
    /// equals <paramref name="title"/> — so emitting the Title separately would
    /// duplicate it. (Parse derives Title from exactly this first H1, so an
    /// already-leading title needs no extra line.)
    /// </summary>
    private static bool TitleIsLeadingHeading(NeutralDoc doc, string title)
    {
        var first = doc.Blocks.Count > 0 ? doc.Blocks[0] : null;
        if (first is not { Kind: NeutralBlockKind.Heading })
        {
            return false;
        }

        var level = first.Level == 0 ? 1 : first.Level;
        if (level != 1)
        {
            return false;
        }

        var text = first.Runs is { } runs ? string.Concat(runs.Select(r => r.Text)) : string.Empty;
        return string.Equals(text, title, StringComparison.Ordinal);
    }

    private static bool EndsWithListLine(StringBuilder sb)
    {
        var text = sb.ToString();
        var lastBreak = text.TrimEnd('\n').LastIndexOf('\n');
        var line = (lastBreak < 0 ? text : text[(lastBreak + 1)..]).TrimEnd('\n').TrimStart();
        return line.StartsWith("- ", StringComparison.Ordinal) || line.StartsWith("1. ", StringComparison.Ordinal);
    }

    private static void WriteTable(StringBuilder sb, NeutralBlock block)
    {
        var rows = block.Rows ?? [];
        if (rows.Count == 0)
        {
            return;
        }

        var columns = Math.Max(1, rows.Max(r => r.Count));

        // A headerless table still needs a header row in GFM; synthesize a blank
        // one so the body rows render, and let HeaderRow=false survive the trip.
        var start = 0;
        IReadOnlyList<string> header;
        if (block.HeaderRow)
        {
            header = rows[0];
            start = 1;
        }
        else
        {
            header = Enumerable.Repeat(string.Empty, columns).ToList();
        }

        WriteRow(sb, header, columns);
        sb.Append('|');
        for (var c = 0; c < columns; c++)
        {
            sb.Append(" --- |");
        }

        sb.Append('\n');

        for (var r = start; r < rows.Count; r++)
        {
            WriteRow(sb, rows[r], columns);
        }
    }

    private static void WriteRow(StringBuilder sb, IReadOnlyList<string> cells, int columns)
    {
        sb.Append('|');
        for (var c = 0; c < columns; c++)
        {
            var text = c < cells.Count ? cells[c] : string.Empty;
            sb.Append(' ').Append(EscapeCell(text)).Append(" |");
        }

        sb.Append('\n');
    }

    /// <summary>Renders a run list to inline markdown, grouping bold/italic/underline/link markers.</summary>
    private static string Inline(IReadOnlyList<NeutralRun>? runs)
    {
        if (runs is null || runs.Count == 0)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        foreach (var run in runs)
        {
            var text = EscapeInline(run.Text);
            if (run.Bold)
            {
                text = "**" + text + "**";
            }

            if (run.Italic)
            {
                text = "*" + text + "*";
            }

            if (run.Underline)
            {
                text = "<u>" + text + "</u>";
            }

            if (run.Href is { Length: > 0 } href)
            {
                text = "[" + text + "](" + href + ")";
            }

            sb.Append(text);
        }

        return sb.ToString().Replace("\n", " ", StringComparison.Ordinal);
    }

    private static string EscapeInline(string text) =>
        text.Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace("\r", "\n", StringComparison.Ordinal);

    private static string EscapeCell(string text) =>
        EscapeInline(text).Replace("|", "\\|", StringComparison.Ordinal).Replace("\n", "<br>", StringComparison.Ordinal);

    // ─────────────────────────────────────────────────────────── parse (md → model)

    /// <summary>Parses GFM markdown text into a <see cref="NeutralDoc"/>; the title is the first H1, if any.</summary>
    public static NeutralDoc Parse(string markdown)
    {
        var text = markdown.Replace("\r\n", "\n", StringComparison.Ordinal).Replace("\r", "\n", StringComparison.Ordinal);
        var lines = text.Split('\n');
        var blocks = new List<NeutralBlock>();
        string? title = null;

        var i = 0;
        while (i < lines.Length)
        {
            var line = lines[i];
            var trimmed = line.TrimStart();

            // Blank line: paragraph separator.
            if (trimmed.Length == 0)
            {
                i++;
                continue;
            }

            // ATX heading.
            if (HeadingLevel(trimmed) is { } level)
            {
                var runs = ParseInline(trimmed[level..].Trim());
                blocks.Add(new NeutralBlock(NeutralBlockKind.Heading, Level: level, Runs: runs));
                title ??= level == 1 ? string.Concat(runs.Select(r => r.Text)) : null;
                i++;
                continue;
            }

            // Image line: ![alt](src), standing alone.
            if (TryParseImageLine(trimmed) is { } image)
            {
                blocks.Add(image);
                i++;
                continue;
            }

            // GFM pipe table: a pipe row followed by a separator row.
            if (IsTableRow(line) && i + 1 < lines.Length && IsSeparatorRow(lines[i + 1]))
            {
                var (table, consumed) = ParseTable(lines, i);
                blocks.Add(table);
                i += consumed;
                continue;
            }

            // List item.
            if (TryParseListItem(line) is { } item)
            {
                blocks.Add(item);
                i++;
                continue;
            }

            // Plain paragraph: gather the contiguous non-blank, non-structural lines.
            var paragraph = new StringBuilder(trimmed);
            i++;
            while (i < lines.Length)
            {
                var next = lines[i];
                var nextTrim = next.TrimStart();
                if (nextTrim.Length == 0 || HeadingLevel(nextTrim) is not null ||
                    TryParseListItem(next) is not null ||
                    (IsTableRow(next) && i + 1 < lines.Length && IsSeparatorRow(lines[i + 1])) ||
                    TryParseImageLine(nextTrim) is not null)
                {
                    break;
                }

                paragraph.Append(' ').Append(nextTrim);
                i++;
            }

            blocks.Add(new NeutralBlock(NeutralBlockKind.Paragraph, Runs: ParseInline(paragraph.ToString())));
        }

        return new NeutralDoc(title, blocks);
    }

    private static int? HeadingLevel(string trimmed)
    {
        var hashes = 0;
        while (hashes < trimmed.Length && trimmed[hashes] == '#')
        {
            hashes++;
        }

        return hashes is >= 1 and <= 6 && hashes < trimmed.Length && trimmed[hashes] == ' ' ? hashes : null;
    }

    private static NeutralBlock? TryParseImageLine(string trimmed)
    {
        if (!trimmed.StartsWith("![", StringComparison.Ordinal))
        {
            return null;
        }

        var altEnd = trimmed.IndexOf("](", StringComparison.Ordinal);
        if (altEnd < 0 || !trimmed.EndsWith(")", StringComparison.Ordinal))
        {
            return null;
        }

        var alt = trimmed[2..altEnd];
        var src = trimmed[(altEnd + 2)..^1];
        return new NeutralBlock(
            NeutralBlockKind.Image,
            Source: src.Length > 0 ? src : null,
            Alt: alt.Length > 0 ? Unescape(alt) : null);
    }

    private static NeutralBlock? TryParseListItem(string line)
    {
        var indent = 0;
        while (indent < line.Length && line[indent] == ' ')
        {
            indent++;
        }

        var rest = line[indent..];
        var level = indent / 2;

        if (rest.StartsWith("- ", StringComparison.Ordinal) || rest.StartsWith("* ", StringComparison.Ordinal) ||
            rest.StartsWith("+ ", StringComparison.Ordinal))
        {
            return new NeutralBlock(
                NeutralBlockKind.ListItem,
                Level: Math.Clamp(level, 0, 8),
                Ordered: false,
                Runs: ParseInline(rest[2..].Trim()));
        }

        // Ordered: "1. ", "12. ", etc.
        var dot = rest.IndexOf(". ", StringComparison.Ordinal);
        if (dot > 0 && rest[..dot].All(char.IsAsciiDigit))
        {
            return new NeutralBlock(
                NeutralBlockKind.ListItem,
                Level: Math.Clamp(level, 0, 8),
                Ordered: true,
                Runs: ParseInline(rest[(dot + 2)..].Trim()));
        }

        return null;
    }

    private static bool IsTableRow(string line) => line.TrimStart().StartsWith('|');

    private static bool IsSeparatorRow(string line)
    {
        var cells = SplitRow(line);
        return cells.Count > 0 && cells.All(c =>
        {
            var t = c.Trim().Replace(":", string.Empty, StringComparison.Ordinal);
            return t.Length > 0 && t.All(ch => ch == '-');
        });
    }

    private static (NeutralBlock Table, int Consumed) ParseTable(string[] lines, int start)
    {
        var headerCells = SplitRow(lines[start]).Select(c => Unescape(c.Trim())).ToList();
        var rows = new List<IReadOnlyList<string>> { headerCells };

        // The header row may be all-empty (the synthesized headerless case).
        var hasHeader = headerCells.Any(c => c.Length > 0);

        var i = start + 2; // skip header + separator
        while (i < lines.Length && IsTableRow(lines[i]))
        {
            rows.Add(SplitRow(lines[i]).Select(c => Unescape(c.Trim()).Replace("<br>", "\n", StringComparison.Ordinal)).ToList());
            i++;
        }

        if (!hasHeader)
        {
            rows.RemoveAt(0); // drop the synthetic blank header row
        }

        return (new NeutralBlock(NeutralBlockKind.Table, Rows: rows, HeaderRow: hasHeader), i - start);
    }

    /// <summary>Splits a GFM pipe row into its cells, honouring <c>\|</c> escapes and the leading/trailing pipes.</summary>
    private static List<string> SplitRow(string line)
    {
        var trimmed = line.Trim();
        if (trimmed.StartsWith('|'))
        {
            trimmed = trimmed[1..];
        }

        if (trimmed.EndsWith('|'))
        {
            trimmed = trimmed[..^1];
        }

        var cells = new List<string>();
        var current = new StringBuilder();
        for (var i = 0; i < trimmed.Length; i++)
        {
            if (trimmed[i] == '\\' && i + 1 < trimmed.Length && trimmed[i + 1] == '|')
            {
                current.Append('|');
                i++;
            }
            else if (trimmed[i] == '|')
            {
                cells.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(trimmed[i]);
            }
        }

        cells.Add(current.ToString());
        return cells;
    }

    /// <summary>Parses inline markdown into runs: links, bold, italic, underline. Unknown markup stays literal text.</summary>
    private static IReadOnlyList<NeutralRun> ParseInline(string text)
    {
        var runs = new List<NeutralRun>();
        AppendInline(runs, text, bold: false, italic: false, underline: false, href: null);
        return runs.Count == 0 && text.Length > 0
            ? [new NeutralRun(Unescape(text))]
            : runs;
    }

    private static void AppendInline(List<NeutralRun> runs, string text, bool bold, bool italic, bool underline, string? href)
    {
        var i = 0;
        var literal = new StringBuilder();

        void Flush()
        {
            if (literal.Length > 0)
            {
                runs.Add(new NeutralRun(Unescape(literal.ToString()), bold, italic, underline, Href: href));
                literal.Clear();
            }
        }

        while (i < text.Length)
        {
            // Escaped char.
            if (text[i] == '\\' && i + 1 < text.Length)
            {
                literal.Append(text[i + 1]);
                i += 2;
                continue;
            }

            // Link: [label](href)
            if (text[i] == '[' && href is null && TryReadLink(text, i, out var label, out var linkHref, out var linkEnd))
            {
                Flush();
                AppendInline(runs, label, bold, italic, underline, linkHref);
                i = linkEnd;
                continue;
            }

            // Underline: <u>…</u>
            if (!underline && text.AsSpan(i).StartsWith("<u>"))
            {
                var close = text.IndexOf("</u>", i + 3, StringComparison.Ordinal);
                if (close >= 0)
                {
                    Flush();
                    AppendInline(runs, text[(i + 3)..close], bold, italic, underline: true, href);
                    i = close + 4;
                    continue;
                }
            }

            // Bold: **…**
            if (!bold && text.AsSpan(i).StartsWith("**"))
            {
                var close = text.IndexOf("**", i + 2, StringComparison.Ordinal);
                if (close >= 0)
                {
                    Flush();
                    AppendInline(runs, text[(i + 2)..close], bold: true, italic, underline, href);
                    i = close + 2;
                    continue;
                }
            }

            // Italic: *…* (single star, not part of **)
            if (!italic && text[i] == '*')
            {
                var close = text.IndexOf('*', i + 1);
                if (close > i + 1)
                {
                    Flush();
                    AppendInline(runs, text[(i + 1)..close], bold, italic: true, underline, href);
                    i = close + 1;
                    continue;
                }
            }

            literal.Append(text[i]);
            i++;
        }

        Flush();
    }

    private static bool TryReadLink(string text, int open, out string label, out string href, out int end)
    {
        label = string.Empty;
        href = string.Empty;
        end = open;

        var close = text.IndexOf(']', open + 1);
        if (close < 0 || close + 1 >= text.Length || text[close + 1] != '(')
        {
            return false;
        }

        var hrefEnd = text.IndexOf(')', close + 2);
        if (hrefEnd < 0)
        {
            return false;
        }

        label = text[(open + 1)..close];
        href = text[(close + 2)..hrefEnd];
        end = hrefEnd + 1;
        return true;
    }

    private static string Unescape(string text)
    {
        if (!text.Contains('\\', StringComparison.Ordinal))
        {
            return text;
        }

        var sb = new StringBuilder(text.Length);
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] == '\\' && i + 1 < text.Length)
            {
                sb.Append(text[i + 1]);
                i++;
            }
            else
            {
                sb.Append(text[i]);
            }
        }

        return sb.ToString();
    }
}
