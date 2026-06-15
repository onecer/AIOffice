using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Table of figures (v1.2.0). A TOF is a complex field
/// (<c>TOC \h \z \c "Figure"</c>) that lists every caption of one label, wrapped
/// in a Table-of-Figures gallery sdt block — the same machinery as the M4 TOC,
/// but driven by the M8 captions rather than headings. Each caption gets a _Toc
/// bookmark so its pre-rendered entry hyperlinks back to it, and a
/// <c>figures_cached</c> warning notes that Word repaginates the page numbers on
/// open (F9), exactly like the TOC.
/// </summary>
public sealed partial class WordHandler
{
    /// <summary>The caption labels a table of figures can be built for.</summary>
    private static readonly string[] TableOfFiguresLabels = ["Figure", "Table", "Equation"];

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body","type":"tableOfFigures","props":{"label":"Figure","position":"before /body/p[1]","title"?}}</c>:
    /// a TOF gallery sdt holding the <c>TOC \h \z \c "Figure"</c> field plus a
    /// pre-rendered, hyperlinked static entry per matching caption. Re-running it
    /// for the same label refreshes that table's entries in place.
    /// </summary>
    private static object ApplyAddTableOfFigures(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var path = DocPath.Parse(op.Path);
        var root = path.Segments[0];
        if (path.Segments.Count != 1 || root.Name != "body" || root.Index is not null || root.Id is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add --type tableOfFigures targets /body, not '{op.Path}'.",
                "Use {\"op\":\"add\",\"path\":\"/body\",\"type\":\"tableOfFigures\",\"props\":{\"label\":\"Figure\",\"position\":\"before /body/p[1]\"}}.",
                candidates: ["/body"]);
        }

        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author");

        var label = props["label"] is { } labelNode ? NodeToString(labelNode) : "Figure";
        if (!TableOfFiguresLabels.Contains(label, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown table-of-figures label '{label}'.",
                "Use label Figure (default), Table or Equation — the same labels captions carry.",
                candidates: TableOfFiguresLabels);
        }

        var title = props["title"] is { } titleNode ? NodeToString(titleNode) : null;
        var position = props["position"] is { } positionNode ? NodeToString(positionNode) : op.Position;

        var body = GetBody(doc, file: string.Empty);
        var existing = EnumerateTablesOfFigures(doc).FirstOrDefault(t => TableOfFiguresLabel(t) == label);
        var anchorBefore = ResolveTocAnchor(doc, position) ?? existing;

        // Entries are this label's captions in body order; each is bookmarked so
        // its entry can hyperlink to it (reusing the _Toc bookmark convention).
        var captions = EnumerateCaptions(doc).Where(c => c.Label == label).ToList();

        var nextBookmarkId = NextBookmarkId(doc);
        var existingNames = new HashSet<string>(
            EnumerateBookmarks(doc).Select(b => b.Name?.Value).OfType<string>(),
            StringComparer.OrdinalIgnoreCase);
        var nameSeed = 1;

        var entries = new List<(string Text, string Bookmark)>();
        foreach (var caption in captions)
        {
            var bookmark = EnsureTocBookmark(caption.Paragraph, ref nextBookmarkId, ref nameSeed, existingNames);
            entries.Add((CaptionEntryText(caption), bookmark));
        }

        EnsureTableOfFiguresStyles(doc, withTitle: title is { Length: > 0 });
        var sdt = BuildTableOfFiguresBlock(entries, label, title);

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
        existing?.Remove();

        if (!session.Warnings.Any(w => w.Code == WarningCodes.FiguresCached))
        {
            session.Warnings.Add(new Warning(
                WarningCodes.FiguresCached,
                "Table-of-figures entries are pre-rendered from the current captions and carry no page numbers: " +
                "pagination only exists once the document is laid out. Word computes the numbers when it opens the " +
                "file (or on field refresh, F9); aioffice does not paginate."));
        }

        return new
        {
            op = "add",
            type = "tableOfFigures",
            path = "/tableOfFigures[1]",
            label,
            entries = entries.Count,
            refreshed = refreshed ? true : (bool?)null,
        };
    }

    /// <summary>remove /tableOfFigures[1]: drops the sdt block (caption bookmarks stay on the captions).</summary>
    private static object ApplyRemoveTableOfFigures(WordprocessingDocument doc, EditOp op)
    {
        var sdt = ResolveTableOfFigures(doc, DocPath.Parse(op.Path));
        sdt.Remove();
        return new { op = "remove", path = "/tableOfFigures[1]", type = "tableOfFigures" };
    }

    // ------------------------------------------------------------------ build

    /// <summary>"Figure 1: Quarterly trend" — the caption's full rendered text for a TOF entry.</summary>
    private static string CaptionEntryText(CaptionInfo caption) =>
        caption.Text is { Length: > 0 }
            ? string.Create(CultureInfo.InvariantCulture, $"{caption.Label} {caption.Number}: {caption.Text}")
            : string.Create(CultureInfo.InvariantCulture, $"{caption.Label} {caption.Number}");

    /// <summary>
    /// The TOF gallery sdt: optional title, then entry paragraphs hosting the
    /// complex field — begin/instr/separate in the first entry, end in the last —
    /// each entry hyperlinked to its caption bookmark, exactly like the TOC.
    /// </summary>
    private static SdtBlock BuildTableOfFiguresBlock(List<(string Text, string Bookmark)> entries, string label, string? title)
    {
        var content = new SdtContentBlock();
        if (title is { Length: > 0 })
        {
            content.AppendChild(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                new Run(NewText(title))));
        }

        var instruction = $" TOC \\h \\z \\c \"{label}\" ";
        var fieldOpeners = new OpenXmlElement[]
        {
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
        };

        if (entries.Count == 0)
        {
            var lone = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "TableofFigures" }));
            lone.Append(fieldOpeners);
            lone.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            content.AppendChild(lone);
        }

        for (var i = 0; i < entries.Count; i++)
        {
            var (text, bookmark) = entries[i];
            var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "TableofFigures" }));

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
                    new DocPartGallery { Val = "Table of Figures" },
                    new DocPartUnique())),
            content);
    }

    /// <summary>The TableofFigures entry style (and the TOCHeading title style when used), on demand.</summary>
    private static void EnsureTableOfFiguresStyles(WordprocessingDocument doc, bool withTitle)
    {
        var styles = EnsureStylesRoot(doc);
        if (FindStyle(styles, "TableofFigures") is null)
        {
            styles.AppendChild(new Style(
                new StyleName { Val = "table of figures" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new SpacingBetweenLines { After = "100" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "TableofFigures",
            });
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

    // ------------------------------------------------------------------- read

    /// <summary>get /tableOfFigures[i] data.</summary>
    private static Dictionary<string, object?> GetTableOfFiguresProperties(WordprocessingDocument doc, DocPath path)
    {
        var sdt = ResolveTableOfFigures(doc, path);
        return TableOfFiguresShape(sdt);
    }

    private static Dictionary<string, object?> TableOfFiguresShape(SdtBlock sdt) => new()
    {
        ["label"] = TableOfFiguresLabel(sdt),
        ["entryCount"] = sdt.Descendants<Hyperlink>().Count(),
        ["title"] = sdt.Descendants<Paragraph>()
            .FirstOrDefault(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "TOCHeading")?.InnerText,
    };

    /// <summary>The label a TOF lists, from its <c>\c "Figure"</c> field switch.</summary>
    private static string? TableOfFiguresLabel(SdtBlock sdt)
    {
        foreach (var instr in sdt.Descendants<FieldCode>())
        {
            var match = Regex.Match(instr.Text, "\\\\c\\s+\"([A-Za-z]+)\"");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
        }

        return null;
    }

    /// <summary>Top-level Table-of-Figures gallery sdt blocks in document order.</summary>
    private static List<SdtBlock> EnumerateTablesOfFigures(WordprocessingDocument doc) =>
        doc.MainDocumentPart?.Document?.Body is { } body
            ? [.. body.Elements<SdtBlock>().Where(IsTableOfFiguresBlock)]
            : [];

    private static bool IsTableOfFiguresBlock(SdtBlock sdt) =>
        sdt.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value
            == "Table of Figures";

    /// <summary>Resolves /tableOfFigures[i] or throws invalid_path with candidates.</summary>
    private static SdtBlock ResolveTableOfFigures(WordprocessingDocument doc, DocPath path)
    {
        var tofs = EnumerateTablesOfFigures(doc);
        var segment = path.Segments[0];
        var index = segment.Index ?? 1;

        if (path.Segments.Count != 1 || segment.Id is not null || segment.Name != "tableOfFigures")
        {
            throw TableOfFiguresNotFound($"'{path.ToCanonicalString()}' is not a table-of-figures path; use /tableOfFigures[1].", tofs);
        }

        if (index > tofs.Count)
        {
            throw TableOfFiguresNotFound(
                tofs.Count == 0
                    ? "This document has no table of figures."
                    : $"/tableOfFigures[{index}] does not exist; the document has {tofs.Count} table-of-figures block(s).",
                tofs);
        }

        return tofs[index - 1];
    }

    private static AiofficeException TableOfFiguresNotFound(string message, List<SdtBlock> tofs) => new(
        ErrorCodes.InvalidPath,
        message,
        "Add one with {\"op\":\"add\",\"path\":\"/body\",\"type\":\"tableOfFigures\",\"props\":{\"label\":\"Figure\"}}.",
        candidates: [.. Enumerable.Range(1, Math.Max(tofs.Count, 0)).Select(n => $"/tableOfFigures[{n}]")]);

    /// <summary>Table-of-figures shapes for read --view structure: {path, label, entryCount}.</summary>
    private static List<object> TablesOfFiguresStructure(WordprocessingDocument doc) =>
        [.. EnumerateTablesOfFigures(doc).Select((sdt, i) => (object)new
        {
            path = $"/tableOfFigures[{i + 1}]",
            properties = TableOfFiguresShape(sdt),
        })];
}
