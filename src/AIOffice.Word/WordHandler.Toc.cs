using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>
    /// <c>{"op":"add","path":"/body","type":"toc","props":{"levels":"1-3","title":…?,"position":"before /body/p[1]"|"inside"}}</c>:
    /// a real Word TOC — one w:sdt gallery block holding the complex TOC field
    /// (begin / instrText / separate … end) PLUS pre-hydrated static entries
    /// generated from the current headings. Each entry is styled TOC1..TOC9 and
    /// hyperlinks to a _Toc bookmark created on its heading. Page numbers are
    /// intentionally absent (toc_pages_unknown warning): Word — and our render —
    /// only know pagination when the document is laid out, and Word recomputes
    /// them on open/refresh (F9). Re-running add refreshes the entries in place.
    /// </summary>
    private static object ApplyAddToc(WordprocessingDocument doc, string file, EditOp op, EditSession session)
    {
        var path = DocPath.Parse(op.Path);
        var root = path.Segments[0];
        if (path.Segments.Count != 1 || root.Name != "body" || root.Index is not null || root.Id is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add --type toc targets /body, not '{op.Path}'.",
                "Use {\"op\":\"add\",\"path\":\"/body\",\"type\":\"toc\",\"props\":{\"levels\":\"1-3\",\"position\":\"before /body/p[1]\"}}.",
                candidates: ["/body"]);
        }

        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author"); // batch attribution metadata, not TOC content
        var (fromLevel, toLevel) = ParseTocLevels(props["levels"] is { } levelsNode ? NodeToString(levelsNode) : "1-3");
        var title = props["title"] is { } titleNode ? NodeToString(titleNode) : null;
        var position = props["position"] is { } positionNode ? NodeToString(positionNode) : op.Position;

        var body = GetBody(doc, file);
        var existing = EnumerateTocs(doc).FirstOrDefault();
        var anchorBefore = ResolveTocAnchor(doc, position) ?? existing; // refresh keeps the old spot

        // Static entries come from the current headings; every heading gets a
        // _Toc bookmark so its entry can hyperlink to it (reused on refresh).
        var headings = body.Descendants<Paragraph>()
            .Select(p => (Paragraph: p, Level: HeadingLevel(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value)))
            .Where(h => h.Level is { } level && level >= fromLevel && level <= toLevel)
            .Select(h => (h.Paragraph, Level: h.Level!.Value))
            .ToList();

        var entries = new List<(int Level, string Text, string Bookmark)>();
        var nextBookmarkId = NextBookmarkId(doc);
        var existingNames = new HashSet<string>(
            EnumerateBookmarks(doc).Select(b => b.Name?.Value).OfType<string>(),
            StringComparer.OrdinalIgnoreCase);
        var nameSeed = 1;
        foreach (var (heading, level) in headings)
        {
            var bookmark = EnsureTocBookmark(heading, ref nextBookmarkId, ref nameSeed, existingNames);
            entries.Add((level, heading.InnerText, bookmark));
        }

        EnsureTocStyles(doc, fromLevel, toLevel, withTitle: title is { Length: > 0 });
        var sdt = BuildTocBlock(entries, fromLevel, toLevel, title);

        if (anchorBefore is not null)
        {
            anchorBefore.InsertBeforeSelf(sdt);
        }
        else if (body.Elements<SectionProperties>().FirstOrDefault() is { } sectPr)
        {
            sectPr.InsertBeforeSelf(sdt);
        }
        else
        {
            body.AppendChild(sdt);
        }

        var refreshed = existing is not null;
        existing?.Remove(); // replace semantics: one TOC, regenerated in place

        if (!session.Warnings.Any(w => w.Code == "toc_pages_unknown"))
        {
            session.Warnings.Add(new Warning(
                "toc_pages_unknown",
                "TOC entries carry no page numbers: pagination only exists once the document is laid out. " +
                "Word computes the numbers when it opens the file (or on field refresh, F9); aioffice render does not paginate."));
        }

        if (refreshed && !session.Warnings.Any(w => w.Code == "toc_refreshed"))
        {
            session.Warnings.Add(new Warning(
                "toc_refreshed",
                "An existing TOC was replaced: entries were regenerated from the current headings."));
        }

        return new
        {
            op = "add",
            type = "toc",
            path = "/toc[1]",
            levels = $"{fromLevel}-{toLevel}",
            entries = entries.Count,
            refreshed = refreshed ? true : (bool?)null,
        };
    }

    /// <summary>remove /toc[1]: drops the sdt block. The _Toc bookmarks stay on the headings (Word leaves them too).</summary>
    private static object ApplyRemoveToc(WordprocessingDocument doc, EditOp op)
    {
        var sdt = ResolveToc(doc, DocPath.Parse(op.Path));
        sdt.Remove();
        return new { op = "remove", path = "/toc[1]", type = "toc" };
    }

    // ------------------------------------------------------------------ read

    /// <summary>get /toc[1] data.</summary>
    private static Dictionary<string, object?> GetTocProperties(WordprocessingDocument doc, DocPath path)
    {
        var sdt = ResolveToc(doc, path);
        return TocShape(sdt);
    }

    private static Dictionary<string, object?> TocShape(SdtBlock sdt) => new()
    {
        ["levels"] = ParseTocFieldLevels(sdt),
        ["entryCount"] = sdt.Descendants<Hyperlink>().Count(),
        ["title"] = sdt.Descendants<Paragraph>()
            .FirstOrDefault(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "TOCHeading")?.InnerText,
    };

    /// <summary>The "a-b" range stored in the TOC field's \o switch, if present.</summary>
    private static string? ParseTocFieldLevels(SdtBlock sdt)
    {
        foreach (var instr in sdt.Descendants<FieldCode>())
        {
            var match = Regex.Match(instr.Text, "\\\\o\\s+\"([1-9])-([1-9])\"");
            if (match.Success)
            {
                return $"{match.Groups[1].Value}-{match.Groups[2].Value}";
            }
        }

        return null;
    }

    /// <summary>Top-level TOC gallery sdt blocks in document order.</summary>
    private static List<SdtBlock> EnumerateTocs(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.Document?.Body is { } body
            ? [.. body.Elements<SdtBlock>().Where(IsTocBlock)]
            : [];

    private static bool IsTocBlock(SdtBlock sdt) =>
        sdt.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value
            == "Table of Contents";

    /// <summary>Resolves /toc[1] or throws invalid_path.</summary>
    private static SdtBlock ResolveToc(WordprocessingDocument doc, DocPath path)
    {
        var tocs = EnumerateTocs(doc);
        var segment = path.Segments[0];
        var index = segment.Index ?? 1;

        if (path.Segments.Count != 1 || segment.Id is not null)
        {
            throw TocNotFound($"'{path.ToCanonicalString()}' is not a TOC path; use /toc[1].", tocs);
        }

        if (index > tocs.Count)
        {
            throw TocNotFound(
                tocs.Count == 0
                    ? "This document has no table of contents."
                    : $"/toc[{index}] does not exist; the document has {tocs.Count} TOC block(s).",
                tocs);
        }

        return tocs[index - 1];
    }

    private static AiofficeException TocNotFound(string message, List<SdtBlock> tocs) => new(
        ErrorCodes.InvalidPath,
        message,
        "Add one with {\"op\":\"add\",\"path\":\"/body\",\"type\":\"toc\",\"props\":{\"levels\":\"1-3\"}}.",
        candidates: [.. Enumerable.Range(1, Math.Max(tocs.Count, 0)).Select(n => $"/toc[{n}]")]);

    // ----------------------------------------------------------------- build

    private static (int From, int To) ParseTocLevels(string levels)
    {
        var match = Regex.Match(levels, "^([1-9])(?:-([1-9]))?$");
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.levels '{levels}' is not a heading-level range.",
                "Use \"a-b\" with 1 <= a <= b <= 9 (e.g. \"1-3\"), or a single number n meaning \"1-n\".");
        }

        var second = match.Groups[2];
        var from = second.Success ? int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture) : 1;
        var to = int.Parse((second.Success ? second : match.Groups[1]).Value, CultureInfo.InvariantCulture);
        if (from > to)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.levels '{levels}' is descending.",
                "Use an ascending range like \"1-3\".");
        }

        return (from, to);
    }

    /// <summary>Parses toc position "inside" | "before &lt;path&gt;" into the top-level body block to insert before (null = append).</summary>
    private static OpenXmlElement? ResolveTocAnchor(WordprocessingDocument doc, string? position)
    {
        if (position is null or "inside")
        {
            return null;
        }

        var parts = position.Split([' ', ':'], 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts is ["before", var targetPath] && targetPath.StartsWith('/'))
        {
            var target = WordAddress.Resolve(doc, DocPath.Parse(targetPath));
            var element = target.Element;
            while (element.Parent is not null and not Body)
            {
                element = element.Parent;
            }

            if (element.Parent is Body)
            {
                return element;
            }

            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"toc position targets a /body block, not {target.CanonicalPath}.",
                "Use \"before /body/p[n]\" (or a table path) — the TOC is a body-level block.");
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"toc position '{position}' is not valid.",
            "Use \"inside\" (append to the body) or \"before <path>\", e.g. \"before /body/p[1]\".",
            candidates: ["inside", "before /body/p[1]"]);
    }

    /// <summary>The heading's _Toc bookmark name, creating bookmarkStart/End when missing (refresh reuses it).</summary>
    private static string EnsureTocBookmark(
        Paragraph heading, ref int nextBookmarkId, ref int nameSeed, HashSet<string> existingNames)
    {
        var existing = heading.Descendants<BookmarkStart>()
            .FirstOrDefault(b => b.Name?.Value?.StartsWith("_Toc", StringComparison.Ordinal) == true);
        if (existing?.Name?.Value is { } name)
        {
            return name;
        }

        string newName;
        do
        {
            newName = "_Toc" + nameSeed.ToString("000000000", CultureInfo.InvariantCulture);
            nameSeed++;
        }
        while (existingNames.Contains(newName));

        existingNames.Add(newName);
        var id = nextBookmarkId.ToString(CultureInfo.InvariantCulture);
        nextBookmarkId++;

        var start = new BookmarkStart { Id = id, Name = newName };
        if (heading.ParagraphProperties is { } pPr)
        {
            pPr.InsertAfterSelf(start);
        }
        else
        {
            heading.PrependChild(start);
        }

        heading.AppendChild(new BookmarkEnd { Id = id });
        return newName;
    }

    /// <summary>
    /// The sdt gallery block: optional TOCHeading title, then entry paragraphs
    /// hosting the complex field — begin/instrText/separate in the first entry,
    /// end in the last. Entries between separate and end are the field's cached
    /// result, exactly where Word keeps its own.
    /// </summary>
    private static SdtBlock BuildTocBlock(List<(int Level, string Text, string Bookmark)> entries, int fromLevel, int toLevel, string? title)
    {
        var content = new SdtContentBlock();
        if (title is { Length: > 0 })
        {
            content.AppendChild(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                new Run(NewText(title))));
        }

        var instruction = $" TOC \\o \"{fromLevel}-{toLevel}\" \\h \\z \\u ";
        var fieldOpeners = new OpenXmlElement[]
        {
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        };

        if (entries.Count == 0)
        {
            var lone = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOC1" }));
            lone.Append(fieldOpeners);
            lone.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            content.AppendChild(lone);
        }

        for (var i = 0; i < entries.Count; i++)
        {
            var (level, text, bookmark) = entries[i];
            // Word maps the highest included heading level to TOC1.
            var styleIndex = Math.Clamp(level - fromLevel + 1, 1, 9);
            var paragraph = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOC" + styleIndex }));

            if (i == 0)
            {
                paragraph.Append(fieldOpeners);
            }

            paragraph.AppendChild(new Hyperlink(new Run(NewText(text)))
            {
                Anchor = bookmark,
                History = true,
            });

            if (i == entries.Count - 1)
            {
                paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            }

            content.AppendChild(paragraph);
        }

        return new SdtBlock(
            new SdtProperties(
                new SdtContentDocPartObject(
                    new DocPartGallery { Val = "Table of Contents" },
                    new DocPartUnique())),
            content);
    }

    /// <summary>TOC1..TOCn (indent ladder) and the TOCHeading title style, defined on demand.</summary>
    private static void EnsureTocStyles(WordprocessingDocument doc, int fromLevel, int toLevel, bool withTitle)
    {
        var styles = EnsureStylesRoot(doc);

        for (var i = 1; i <= toLevel - fromLevel + 1; i++)
        {
            var id = "TOC" + i;
            if (FindStyle(styles, id) is null)
            {
                styles.AppendChild(new Style(
                    new StyleName { Val = "toc " + i },
                    new BasedOn { Val = "Normal" },
                    new StyleParagraphProperties(new Indentation
                    {
                        Left = ((i - 1) * 220).ToString(CultureInfo.InvariantCulture),
                    }))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = id,
                });
            }
        }

        if (withTitle && FindStyle(styles, "TOCHeading") is null)
        {
            EnsureStyleDefined(doc, "Heading1");
            styles.AppendChild(new Style(
                new StyleName { Val = "TOC Heading" },
                new BasedOn { Val = "Heading1" })
            {
                Type = StyleValues.Paragraph,
                StyleId = "TOCHeading",
            });
        }
    }
}
