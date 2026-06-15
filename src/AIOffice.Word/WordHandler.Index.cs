using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Index (v1.2.0). An index entry is an <c>XE</c> field marked on a paragraph
/// ("AI" or "AI:agents" for a sub-entry); the index itself is an <c>INDEX</c>
/// field wrapped in an Index gallery sdt, plus pre-rendered alphabetized entries
/// built from the marked XE fields. Page numbers cannot be known headlessly, so
/// they are cached as "?" and an <c>index_cached</c> warning says Word computes
/// them on open (F9), mirroring the TOC/TOF contract.
/// </summary>
public sealed partial class WordHandler
{
    // -------------------------------------------------------------- index entry

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[3]","type":"indexEntry","props":{"text":"AI","subEntry"?,"find"?}}</c>:
    /// places an XE field on the paragraph. With props.find the field lands before
    /// the matched text; otherwise it is appended. A sub-entry encodes as Word's
    /// "main:sub" XE form.
    /// </summary>
    private static object ApplyAddIndexEntry(WordprocessingDocument doc, EditOp op)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author");

        var text = props["text"] is { } textNode ? NodeToString(textNode).Trim() : null;
        if (string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type indexEntry needs props.text (the term to index).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[3]\",\"type\":\"indexEntry\",\"props\":{\"text\":\"AI\"}}.");
        }

        var subEntry = props["subEntry"] is { } subNode ? NodeToString(subNode).Trim() : null;
        if (text.Contains('"', StringComparison.Ordinal) || subEntry?.Contains('"', StringComparison.Ordinal) == true)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Index entry text must not contain double quotes.",
                "Quotes would corrupt the XE field instruction; remove them from props.text/props.subEntry.");
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Index entries are marked on paragraphs, not '{anchor.Type}'.",
                anchor.Type is "tc" or "header" or "footer"
                    ? $"Target the paragraph inside it: {anchor.CanonicalPath}/p[1]."
                    : "Pick a paragraph path, e.g. /body/p[3].");
        }

        var entryText = subEntry is { Length: > 0 } ? $"{text}:{subEntry}" : text;
        var instruction = $" XE \"{entryText}\" ";
        var find = props["find"] is { } findNode ? NodeToString(findNode) : null;

        // An XE field has no visible result — it is a hidden marker — so its
        // field frame is begin / instr / end with no separate/cached run.
        if (find is { Length: > 0 })
        {
            PlaceXeFieldAtText(paragraph, find, instruction);
        }
        else
        {
            AppendXeField(paragraph, instruction);
        }

        return new
        {
            op = "add",
            type = "indexEntry",
            path = anchor.CanonicalPath,
            text,
            subEntry,
            entry = entryText,
            note = "The XE field is a hidden marker; build the index with add --type index at /body.",
        };
    }

    /// <summary>Appends an XE marker field (begin / instr / end, no result) to a paragraph.</summary>
    private static void AppendXeField(Paragraph paragraph, string instruction)
    {
        foreach (var run in XeFieldRuns(instruction))
        {
            paragraph.AppendChild(run);
        }
    }

    /// <summary>Inserts an XE marker before the first occurrence of the find text, splitting at the boundary; else appends.</summary>
    private static void PlaceXeFieldAtText(Paragraph paragraph, string find, string instruction) =>
        PlaceFieldAtText(paragraph, find, () => XeFieldRuns(instruction), () => AppendXeField(paragraph, instruction));

    private static List<Run> XeFieldRuns(string instruction) =>
    [
        new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
        new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
        new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
    ];

    // -------------------------------------------------------------------- index

    /// <summary>
    /// <c>{"op":"add","path":"/body","type":"index","props":{"columns"?:2}}</c>:
    /// an INDEX gallery sdt holding the <c>INDEX</c> field plus pre-rendered,
    /// alphabetized entries built from every marked XE field. Page numbers are
    /// cached as "?" (an index_cached warning explains). Re-running it refreshes
    /// the entries in place.
    /// </summary>
    private static object ApplyAddIndex(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var path = DocPath.Parse(op.Path);
        var root = path.Segments[0];
        if (path.Segments.Count != 1 || root.Name != "body" || root.Index is not null || root.Id is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add --type index targets /body, not '{op.Path}'.",
                "Use {\"op\":\"add\",\"path\":\"/body\",\"type\":\"index\",\"props\":{\"columns\":2}}.",
                candidates: ["/body"]);
        }

        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author");

        var columns = props["columns"] is { } columnsNode && int.TryParse(NodeToString(columnsNode), out var c) && c >= 1
            ? c
            : 2;

        var body = GetBody(doc, file: string.Empty);
        var existing = EnumerateIndexes(doc).FirstOrDefault();

        var entries = AlphabetizedIndexEntries(doc);

        EnsureIndexStyles(doc);
        var sdt = BuildIndexBlock(entries, columns);

        // The index is conventionally placed at the end (before sectPr); a fresh
        // add appends, a refresh replaces the existing block in place.
        if (existing is not null)
        {
            existing.InsertBeforeSelf(sdt);
            existing.Remove();
        }
        else if (body.Elements<SectionProperties>().FirstOrDefault() is { } sectPr)
        {
            sectPr.InsertBeforeSelf(sdt);
        }
        else
        {
            body.AppendChild(sdt);
        }

        if (!session.Warnings.Any(w => w.Code == WarningCodes.IndexCached))
        {
            session.Warnings.Add(new Warning(
                WarningCodes.IndexCached,
                "Index entries are pre-rendered and alphabetized from the marked XE fields; page numbers are cached " +
                "as \"?\" because pagination only exists once the document is laid out. Word computes the numbers when " +
                "it opens the file (or on field refresh, F9)."));
        }

        var refreshed = existing is not null;
        return new
        {
            op = "add",
            type = "index",
            path = "/index[1]",
            columns,
            entries = entries.Count,
            refreshed = refreshed ? true : (bool?)null,
        };
    }

    /// <summary>remove /index[1]: drops the sdt block (the XE markers stay in the body).</summary>
    private static object ApplyRemoveIndex(WordprocessingDocument doc, EditOp op)
    {
        var sdt = ResolveIndex(doc, DocPath.Parse(op.Path));
        sdt.Remove();
        return new { op = "remove", path = "/index[1]", type = "index" };
    }

    // ------------------------------------------------------------------ build

    /// <summary>The distinct, alphabetized "main[:sub]" terms of every XE field, body order broken by sort.</summary>
    private static List<string> AlphabetizedIndexEntries(WordprocessingDocument doc)
    {
        var entries = EnumerateIndexEntries(doc).Select(e => e.Entry);
        return [.. entries
            .Distinct(StringComparer.Ordinal)
            .OrderBy(e => e, StringComparer.OrdinalIgnoreCase)
            .ThenBy(e => e, StringComparer.Ordinal)];
    }

    /// <summary>
    /// The INDEX gallery sdt: the complex field (begin/instr/separate ... end)
    /// with one pre-rendered, alphabetized entry paragraph per term between the
    /// separate and end markers, page numbers cached as "?".
    /// </summary>
    private static SdtBlock BuildIndexBlock(List<string> entries, int columns)
    {
        var content = new SdtContentBlock();

        var instruction = $" INDEX \\c \"{columns}\" \\z \"1033\" ";
        var fieldOpeners = new OpenXmlElement[]
        {
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        };

        if (entries.Count == 0)
        {
            var lone = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "IndexHeading" }));
            lone.Append(fieldOpeners);
            lone.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            content.AppendChild(lone);
        }

        for (var i = 0; i < entries.Count; i++)
        {
            var entry = entries[i];
            // A sub-entry ("main:sub") renders one level deeper, like Word.
            var isSub = entry.Contains(':', StringComparison.Ordinal);
            var display = isSub ? entry[(entry.IndexOf(':', StringComparison.Ordinal) + 1)..] : entry;
            var paragraph = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = isSub ? "Index2" : "Index1" }));

            if (i == 0)
            {
                paragraph.Append(fieldOpeners);
            }

            // "term, ?" — the cached page number is unknown until Word paginates.
            paragraph.AppendChild(new Run(NewText(display + ", ?")));

            if (i == entries.Count - 1)
            {
                paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            }

            content.AppendChild(paragraph);
        }

        return new SdtBlock(
            new SdtProperties(
                new SdtContentDocPartObject(
                    new DocPartGallery { Val = "Index" },
                    new DocPartUnique())),
            content);
    }

    /// <summary>Index1/Index2 entry styles, defined on demand (indent ladder).</summary>
    private static void EnsureIndexStyles(WordprocessingDocument doc)
    {
        var styles = EnsureStylesRoot(doc);
        foreach (var (id, name, indent) in new[]
        {
            ("Index1", "index 1", 0),
            ("Index2", "index 2", 220),
            ("IndexHeading", "Index Heading", 0),
        })
        {
            if (FindStyle(styles, id) is not null)
            {
                continue;
            }

            var pPr = indent > 0
                ? new StyleParagraphProperties(new Indentation { Left = indent.ToString(CultureInfo.InvariantCulture) })
                : new StyleParagraphProperties();

            styles.AppendChild(new Style(
                new StyleName { Val = name },
                new BasedOn { Val = "Normal" },
                pPr)
            {
                Type = StyleValues.Paragraph,
                StyleId = id,
            });
        }
    }

    // ------------------------------------------------------------------- read

    /// <summary>An XE field resolved from the document: its full "main[:sub]" entry and the field code.</summary>
    internal sealed record IndexEntryInfo(string Entry, FieldCode Code);

    /// <summary>Every XE field across the body, in document order.</summary>
    private static List<IndexEntryInfo> EnumerateIndexEntries(WordprocessingDocument doc)
    {
        var entries = new List<IndexEntryInfo>();
        if (doc.MainDocumentPart?.Document?.Body is not { } body)
        {
            return entries;
        }

        foreach (var code in body.Descendants<FieldCode>())
        {
            if (XeFieldEntry(code.Text) is { } entry)
            {
                entries.Add(new IndexEntryInfo(entry, code));
            }
        }

        return entries;
    }

    /// <summary>The indexed term of an XE field ("XE \"AI:agents\"" → "AI:agents"), or null if not an XE field.</summary>
    internal static string? XeFieldEntry(string? instruction)
    {
        if (instruction is null)
        {
            return null;
        }

        var match = Regex.Match(instruction.Trim(), "^XE\\s+\"([^\"]*)\"");
        return match.Success ? match.Groups[1].Value : null;
    }

    /// <summary>get /index[i] data: {columns, entryCount, markedEntries}.</summary>
    private static Dictionary<string, object?> GetIndexProperties(WordprocessingDocument doc, DocPath path)
    {
        var sdt = ResolveIndex(doc, path);
        return IndexShape(doc, sdt);
    }

    private static Dictionary<string, object?> IndexShape(WordprocessingDocument doc, SdtBlock sdt) => new()
    {
        ["columns"] = IndexColumns(sdt),
        ["entryCount"] = sdt.Descendants<Paragraph>()
            .Count(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value is "Index1" or "Index2"),
        ["markedEntries"] = EnumerateIndexEntries(doc).Select(e => e.Code).Count(),
    };

    /// <summary>The column count from an INDEX field's <c>\c "n"</c> switch, defaulting to 1.</summary>
    private static int IndexColumns(SdtBlock sdt)
    {
        foreach (var instr in sdt.Descendants<FieldCode>())
        {
            var match = Regex.Match(instr.Text, "INDEX.*\\\\c\\s+\"([0-9]+)\"");
            if (match.Success && int.TryParse(match.Groups[1].Value, out var n))
            {
                return n;
            }
        }

        return 1;
    }

    /// <summary>Top-level Index gallery sdt blocks in document order.</summary>
    private static List<SdtBlock> EnumerateIndexes(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.Document?.Body is { } body
            ? [.. body.Elements<SdtBlock>().Where(IsIndexBlock)]
            : [];

    private static bool IsIndexBlock(SdtBlock sdt) =>
        sdt.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value
            == "Index";

    /// <summary>Resolves /index[i] or throws invalid_path with candidates.</summary>
    private static SdtBlock ResolveIndex(WordprocessingDocument doc, DocPath path)
    {
        var indexes = EnumerateIndexes(doc);
        var segment = path.Segments[0];
        var index = segment.Index ?? 1;

        if (path.Segments.Count != 1 || segment.Id is not null || segment.Name != "index")
        {
            throw IndexNotFound($"'{path.ToCanonicalString()}' is not an index path; use /index[1].", indexes);
        }

        if (index > indexes.Count)
        {
            throw IndexNotFound(
                indexes.Count == 0
                    ? "This document has no index."
                    : $"/index[{index}] does not exist; the document has {indexes.Count} index block(s).",
                indexes);
        }

        return indexes[index - 1];
    }

    private static AiofficeException IndexNotFound(string message, List<SdtBlock> indexes) => new(
        ErrorCodes.InvalidPath,
        message,
        "Mark entries with add --type indexEntry, then generate it with {\"op\":\"add\",\"path\":\"/body\",\"type\":\"index\"}.",
        candidates: [.. Enumerable.Range(1, Math.Max(indexes.Count, 0)).Select(n => $"/index[{n}]")]);

    /// <summary>Index shapes for read --view structure: {path, columns, entryCount, markedEntries}.</summary>
    private static List<object> IndexesStructure(WordprocessingDocument doc) =>
        [.. EnumerateIndexes(doc).Select((sdt, i) => (object)new
        {
            path = $"/index[{i + 1}]",
            properties = IndexShape(doc, sdt),
        })];
}
