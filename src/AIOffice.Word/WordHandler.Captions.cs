using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Captions and cross-references (M8). A caption is a <c>Caption</c>-styled
/// paragraph that pairs a SEQ field ("Figure { SEQ Figure }") with descriptive
/// text and is wrapped in a <c>_Ref…</c> bookmark so it can be referenced; the
/// SEQ field auto-numbers so Word renumbers on open. A cross-reference is a REF
/// field pointing at that bookmark with cached display text ("Figure 1"); like
/// the TOC, cached numbers reflect insert order and Word recomputes them on
/// open (a <c>caption_numbers_cached</c> warning says so).
/// </summary>
public sealed partial class WordHandler
{
    /// <summary>The caption labels Word ships a SEQ counter for.</summary>
    private static readonly string[] CaptionLabels = ["Figure", "Table", "Equation"];

    /// <summary>How a cross-reference renders the target.</summary>
    private static readonly string[] CrossRefShows = ["labelAndNumber", "numberOnly", "text", "page"];

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[3]","type":"caption","props":{"label":"Figure","text":"Quarterly trend","position":"after"}}</c>:
    /// inserts a Caption-styled paragraph — "Figure " + SEQ field + ": text" —
    /// wrapped in a _Ref bookmark, before/after the anchor block.
    /// </summary>
    private static object ApplyAddCaption(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author");

        var label = props["label"] is { } labelNode ? NodeToString(labelNode) : null;
        if (label is null || !CaptionLabels.Contains(label, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                label is null ? "add --type caption needs props.label." : $"Unknown caption label '{label}'.",
                "Use label Figure, Table or Equation, e.g. " +
                "{\"op\":\"add\",\"path\":\"/body/p[3]\",\"type\":\"caption\",\"props\":{\"label\":\"Figure\",\"text\":\"Quarterly trend\"}}.",
                candidates: CaptionLabels);
        }

        var text = props["text"] is { } textNode ? NodeToString(textNode) : string.Empty;
        var position = props["position"] is { } positionNode ? NodeToString(positionNode) : op.Position ?? "after";
        if (position is not ("before" or "after"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"caption position '{position}' is not valid.",
                "Use position \"before\" or \"after\" to place the caption relative to the anchor block.",
                candidates: ["before", "after"]);
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not (Paragraph or Table) || !anchor.CanonicalPath.StartsWith("/body", StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Captions attach to body paragraphs or tables, not '{anchor.Type}' at {anchor.CanonicalPath}.",
                anchor.Type is "tc"
                    ? $"Target the paragraph inside the cell: {anchor.CanonicalPath}/p[1]."
                    : "Pick a body block path, e.g. /body/p[3] (the figure) or /body/table[1].");
        }

        EnsureCaptionStyle(doc);

        // The cached number is this caption's 1-based ordinal among same-label
        // captions, in the position it will occupy — Word recomputes on open.
        var existingSameLabel = EnumerateCaptions(doc).Count(c => c.Label == label);
        var number = existingSameLabel + 1;

        var bookmarkName = NextRefBookmarkName(doc);
        var bookmarkId = NextBookmarkId(doc).ToString(CultureInfo.InvariantCulture);

        var caption = BuildCaptionParagraph(label, number, text, bookmarkName, bookmarkId);

        if (position == "before")
        {
            anchor.Element.InsertBeforeSelf(caption);
        }
        else
        {
            anchor.Element.InsertAfterSelf(caption);
        }

        if (!session.Warnings.Any(w => w.Code == "caption_numbers_cached"))
        {
            session.Warnings.Add(new Warning(
                "caption_numbers_cached",
                "Caption numbers are cached in insert order: the SEQ field stores the computed value, and Word " +
                "renumbers all captions when it opens the file or on field refresh (F9). Cross-reference numbers " +
                "reflect the order at insert time until then."));
        }

        var captionPath = CaptionPath(label, number);
        return new
        {
            op = "add",
            type = "caption",
            path = captionPath,
            label,
            number,
            text,
            bookmark = bookmarkName,
            anchor = anchor.CanonicalPath,
            note = "The SEQ field auto-numbers; Word renumbers captions on open or field refresh (F9).",
        };
    }

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[5]","type":"crossRef","props":{"to":"/caption[@label=Figure][1]","show":"labelAndNumber"}}</c>:
    /// appends a REF field to the target paragraph pointing at the caption's
    /// bookmark, with cached display text ("Figure 1").
    /// </summary>
    private static object ApplyAddCrossRef(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        var author = session.ResolveAuthor(props);

        var to = props["to"] is { } toNode ? NodeToString(toNode) : null;
        if (string.IsNullOrEmpty(to))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type crossRef needs props.to (a caption path or label+index).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[5]\",\"type\":\"crossRef\",\"props\":{\"to\":\"/caption[@label=Figure][1]\",\"show\":\"labelAndNumber\"}}.");
        }

        var show = props["show"] is { } showNode ? NodeToString(showNode) : "labelAndNumber";
        if (!CrossRefShows.Contains(show, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown crossRef show '{show}'.",
                "Use show labelAndNumber (default), numberOnly, text or page.",
                candidates: CrossRefShows);
        }

        var target = ResolveCaption(doc, to);
        var leadingText = props["leadingText"] is { } leadingNode ? NodeToString(leadingNode) : null;

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cross-references are appended to paragraphs, not '{anchor.Type}'.",
                anchor.Type is "tc"
                    ? $"Target the paragraph inside the cell: {anchor.CanonicalPath}/p[1]."
                    : "Pick a paragraph path, e.g. /body/p[5].");
        }

        // Remember the anchor's last pre-existing child so a tracked add can wrap
        // exactly the runs we are about to append (and nothing the author wrote).
        var lastExisting = paragraph.LastChild;

        if (leadingText is { Length: > 0 })
        {
            paragraph.AppendChild(new Run(NewText(leadingText)));
        }

        var displayText = CrossRefDisplayText(target, show);
        var instruction = show switch
        {
            "page" => $" PAGEREF {target.Bookmark} \\h ",
            _ => $" REF {target.Bookmark} \\h ",
        };

        AppendComplexField(paragraph, instruction, displayText);

        // Tracked: wrap ONLY the newly-appended runs (optional leadingText plus the
        // complex field's begin/instr/separate/cached/end) in a single w:ins. As with
        // the footnote ref we deliberately skip MarkParagraphInserted: the anchor's
        // pre-existing runs and paragraph mark are the author's, not part of this add.
        if (session.Track)
        {
            WrapAppendedFieldRuns(doc, paragraph, lastExisting, author);
        }

        if (!session.Warnings.Any(w => w.Code == "caption_numbers_cached"))
        {
            session.Warnings.Add(new Warning(
                "caption_numbers_cached",
                "The cross-reference's displayed number is cached: Word recomputes it (and renumbers captions) " +
                "when it opens the file or on field refresh (F9)."));
        }

        return new
        {
            op = "add",
            type = "crossRef",
            path = anchor.CanonicalPath,
            to = CaptionPath(target.Label, target.Number),
            target = target.Bookmark,
            show,
            cached = displayText,
            note = "Word recomputes the reference when it opens the file or on field refresh (F9).",
        };
    }

    // ------------------------------------------------------------------ build

    /// <summary>A Caption-styled paragraph: bookmark + "Label " + SEQ field + ": text".</summary>
    private static Paragraph BuildCaptionParagraph(string label, int number, string text, string bookmarkName, string bookmarkId)
    {
        var paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Caption" }));

        // The bookmark wraps the whole caption so a REF can show its full content.
        paragraph.AppendChild(new BookmarkStart { Id = bookmarkId, Name = bookmarkName });
        paragraph.AppendChild(new Run(NewText(label + " ")));

        // The SEQ complex field, cached with the computed number.
        AppendComplexField(paragraph, $" SEQ {label} \\* ARABIC ", number.ToString(CultureInfo.InvariantCulture));

        if (text is { Length: > 0 })
        {
            paragraph.AppendChild(new Run(NewText(": " + text)));
        }

        paragraph.AppendChild(new BookmarkEnd { Id = bookmarkId });
        return paragraph;
    }

    /// <summary>Appends a complex field (begin / instr / separate / cached result / end) to a paragraph.</summary>
    private static void AppendComplexField(Paragraph paragraph, string instruction, string cachedResult)
    {
        paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        paragraph.AppendChild(new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }));
        paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        paragraph.AppendChild(new Run(NewText(cachedResult)));
        paragraph.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
    }

    /// <summary>
    /// Tracked-add closure shared by crossRef / mergeField / ifField appends: wraps
    /// the runs appended after <paramref name="lastExisting"/> (the field's
    /// begin/instr/separate/cached/end) in a single w:ins, leaving the anchor's
    /// pre-existing runs and paragraph mark untouched (they are the author's). Like
    /// the footnote ref it deliberately skips MarkParagraphInserted.
    /// </summary>
    private static void WrapAppendedFieldRuns(
        WordprocessingDocument doc, Paragraph paragraph, OpenXmlElement? lastExisting, string author)
    {
        var newRuns = (lastExisting is null
            ? paragraph.Elements<Run>()
            : lastExisting.ElementsAfter().OfType<Run>()).ToList();
        var ins = NewTrackChange(new InsertedRun(), author, DateTime.UtcNow, NextRevisionId(doc));
        newRuns[0].InsertBeforeSelf(ins);
        foreach (var run in newRuns)
        {
            run.Remove();
            ins.AppendChild(run);
        }
    }

    /// <summary>The cached display text a cross-reference shows for its target and mode.</summary>
    private static string CrossRefDisplayText(CaptionInfo target, string show) => show switch
    {
        "numberOnly" => target.Number.ToString(CultureInfo.InvariantCulture),
        "text" => target.Text is { Length: > 0 } ? target.Text : $"{target.Label} {target.Number}",
        "page" => "1", // pagination only exists once Word lays the document out
        _ => $"{target.Label} {target.Number}",
    };

    // ------------------------------------------------------------------- read

    /// <summary>A caption resolved from the document: its label, cached number, text and bookmark.</summary>
    internal sealed record CaptionInfo(Paragraph Paragraph, string Label, int Number, string Text, string Bookmark, string AnchorPath);

    /// <summary>
    /// All captions in body order: Caption-styled paragraphs carrying a SEQ field.
    /// The number is taken from the SEQ field's cached result; ties fall back to
    /// the paragraph's ordinal among same-label captions.
    /// </summary>
    private static List<CaptionInfo> EnumerateCaptions(WordprocessingDocument doc)
    {
        var captions = new List<CaptionInfo>();
        if (doc.MainDocumentPart?.Document?.Body is not { } body)
        {
            return captions;
        }

        var perLabelOrdinal = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var node in WordAddress.EnumerateBody(body).Where(n => n.Type == "p"))
        {
            var paragraph = (Paragraph)node.Element;
            if (paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value != "Caption")
            {
                continue;
            }

            var label = SeqFieldLabel(paragraph);
            if (label is null || !CaptionLabels.Contains(label, StringComparer.Ordinal))
            {
                continue;
            }

            perLabelOrdinal.TryGetValue(label, out var ordinal);
            ordinal++;
            perLabelOrdinal[label] = ordinal;

            var bookmark = paragraph.Descendants<BookmarkStart>()
                .FirstOrDefault(b => b.Name?.Value?.StartsWith("_Ref", StringComparison.Ordinal) == true)?.Name?.Value
                ?? string.Empty;

            captions.Add(new CaptionInfo(
                paragraph, label, ordinal, CaptionText(paragraph, label), bookmark, node.CanonicalPath));
        }

        return captions;
    }

    /// <summary>The label of the first SEQ field in a paragraph ("SEQ Figure \* ARABIC" -> "Figure"), or null.</summary>
    private static string? SeqFieldLabel(Paragraph paragraph)
    {
        foreach (var code in paragraph.Descendants<FieldCode>())
        {
            var match = Regex.Match(code.Text.Trim(), @"^SEQ\s+(\S+)");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
        }

        return null;
    }

    /// <summary>The descriptive text of a caption: everything after the "Label N: " prefix.</summary>
    private static string CaptionText(Paragraph paragraph, string label)
    {
        // Innertext is "Label " + cachedNumber + ": text". Strip the label, the
        // cached number run and the ": " separator to recover the text.
        var inner = paragraph.InnerText;
        var colon = inner.IndexOf(": ", StringComparison.Ordinal);
        return colon >= 0 ? inner[(colon + 2)..] : string.Empty;
    }

    /// <summary>get /caption[@label=Figure][i]: {label, number, text, bookmark, anchorPath}.</summary>
    private static (string Path, Dictionary<string, object?> Properties) GetCaptionProperties(
        WordprocessingDocument doc, string pathArg)
    {
        var caption = ResolveCaption(doc, pathArg);
        return (CaptionPath(caption.Label, caption.Number), new Dictionary<string, object?>
        {
            ["label"] = caption.Label,
            ["number"] = caption.Number,
            ["text"] = caption.Text,
            ["bookmark"] = caption.Bookmark,
            ["anchorPath"] = caption.AnchorPath,
        });
    }

    /// <summary>get /crossRef[i]: {to, show, cached, target}.</summary>
    private static (string Path, Dictionary<string, object?> Properties) GetCrossRefProperties(
        WordprocessingDocument doc, string pathArg)
    {
        var refs = EnumerateCrossRefs(doc);
        var index = ParseSingleIndex(pathArg, "crossRef");
        if (index < 1 || index > refs.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                refs.Count == 0
                    ? "This document has no cross-references."
                    : $"/crossRef[{index}] does not exist; there are {refs.Count} cross-reference(s).",
                "Run 'aioffice read <file> --view structure' to list captions, or add a crossRef first.",
                candidates: [.. Enumerable.Range(1, refs.Count).Select(n => $"/crossRef[{n}]")]);
        }

        var (bookmark, show, cached) = refs[index - 1];
        var target = EnumerateCaptions(doc).FirstOrDefault(c => c.Bookmark == bookmark);
        return ($"/crossRef[{index}]", new Dictionary<string, object?>
        {
            ["to"] = target is null ? bookmark : CaptionPath(target.Label, target.Number),
            ["target"] = bookmark,
            ["show"] = show,
            ["cached"] = cached,
        });
    }

    /// <summary>Every REF/PAGEREF field whose target is a _Ref caption bookmark, in document order.</summary>
    private static List<(string Bookmark, string Show, string Cached)> EnumerateCrossRefs(WordprocessingDocument doc)
    {
        var refs = new List<(string, string, string)>();
        var captionBookmarks = EnumerateCaptions(doc).Select(c => c.Bookmark).ToHashSet(StringComparer.Ordinal);

        foreach (var code in doc.MainDocumentPart?.Document?.Descendants<FieldCode>() ?? [])
        {
            var tokens = code.Text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length < 2 || tokens[0] is not ("REF" or "PAGEREF"))
            {
                continue;
            }

            var bookmark = tokens[1];
            if (!captionBookmarks.Contains(bookmark))
            {
                continue;
            }

            var show = tokens[0] == "PAGEREF" ? "page" : "labelAndNumber";
            var cached = ComplexFieldResult(code);
            refs.Add((bookmark, show, cached));
        }

        return refs;
    }

    /// <summary>The cached result run text that follows a complex field's separate marker.</summary>
    private static string ComplexFieldResult(FieldCode code)
    {
        // Walk forward from the FieldCode's run to the next FieldChar=end, collecting text after separate.
        var startRun = code.Ancestors<Run>().FirstOrDefault() ?? code.Parent as Run;
        if (startRun?.Parent is not { } parent)
        {
            return string.Empty;
        }

        var siblings = parent.ChildElements.ToList();
        var startIndex = startRun is null ? -1 : siblings.IndexOf(startRun);
        var result = string.Empty;
        var afterSeparate = false;
        for (var i = startIndex + 1; i < siblings.Count; i++)
        {
            if (siblings[i] is Run run)
            {
                var fieldChar = run.GetFirstChild<FieldChar>();
                if (fieldChar?.FieldCharType?.Value == FieldCharValues.Separate)
                {
                    afterSeparate = true;
                    continue;
                }

                if (fieldChar?.FieldCharType?.Value == FieldCharValues.End)
                {
                    break;
                }

                if (afterSeparate)
                {
                    result += run.InnerText;
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Resolves a caption reference: either a "/caption[@label=Figure][i]" path
    /// or a bare "label+index" form. Throws invalid_path with candidates.
    /// </summary>
    private static CaptionInfo ResolveCaption(WordprocessingDocument doc, string reference)
    {
        var captions = EnumerateCaptions(doc);

        var match = Regex.Match(reference.Trim(), @"^/?caption\[@label=([A-Za-z]+)\]\[([0-9]+)\]$");
        if (match.Success)
        {
            var label = match.Groups[1].Value;
            var number = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
            var found = captions.FirstOrDefault(c => c.Label == label && c.Number == number);
            return found ?? throw CaptionNotFound(
                $"No {label} caption number {number} exists.", captions);
        }

        // "Figure 1" shorthand.
        var shorthand = Regex.Match(reference.Trim(), @"^([A-Za-z]+)\s+([0-9]+)$");
        if (shorthand.Success)
        {
            var label = shorthand.Groups[1].Value;
            var number = int.Parse(shorthand.Groups[2].Value, CultureInfo.InvariantCulture);
            var found = captions.FirstOrDefault(c => c.Label == label && c.Number == number);
            return found ?? throw CaptionNotFound($"No {label} caption number {number} exists.", captions);
        }

        throw CaptionNotFound($"'{reference}' is not a caption reference.", captions);
    }

    private static int ParseSingleIndex(string pathArg, string root)
    {
        var match = Regex.Match(pathArg.Trim(), @"^/?" + Regex.Escape(root) + @"\[([0-9]+)\]$");
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"'{pathArg}' is not a /{root}[i] path.",
                $"Address the {root} by 1-based index, e.g. /{root}[1].");
        }

        return int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
    }

    private static AiofficeException CaptionNotFound(string message, List<CaptionInfo> captions) => new(
        ErrorCodes.InvalidPath,
        message,
        "Run 'aioffice read <file> --view structure' to list captions with their labels and numbers.",
        candidates: [.. captions.Take(5).Select(c => CaptionPath(c.Label, c.Number))]);

    private static string CaptionPath(string label, int number) =>
        string.Create(CultureInfo.InvariantCulture, $"/caption[@label={label}][{number}]");

    /// <summary>The next free _Ref bookmark name (Word's auto-reference bookmark convention).</summary>
    private static string NextRefBookmarkName(WordprocessingDocument doc)
    {
        var existing = EnumerateBookmarks(doc).Select(b => b.Name?.Value).OfType<string>().ToHashSet(StringComparer.Ordinal);
        var seed = 100000000; // _RefNNNNNNNNN, like Word
        string name;
        do
        {
            name = "_Ref" + seed.ToString(CultureInfo.InvariantCulture);
            seed++;
        }
        while (existing.Contains(name));

        return name;
    }

    /// <summary>The Caption paragraph style, defined on demand (based on Normal, italic, smaller).</summary>
    private static void EnsureCaptionStyle(WordprocessingDocument doc)
    {
        var styles = EnsureStylesRoot(doc);
        if (FindStyle(styles, "Caption") is not null)
        {
            return;
        }

        styles.AppendChild(new Style(
            new StyleName { Val = "caption" },
            new BasedOn { Val = "Normal" },
            new StyleParagraphProperties(new SpacingBetweenLines { After = "200" }),
            new StyleRunProperties(
                new Italic(),
                new Color { Val = "44546A" },
                new FontSize { Val = "18" }))
        {
            Type = StyleValues.Paragraph,
            StyleId = "Caption",
        });
    }

    // ------------------------------------------------------------ structure

    /// <summary>Caption shapes for read --view structure: {path, label, number, text}.</summary>
    private static List<object> CaptionsStructure(WordprocessingDocument doc) =>
        [.. EnumerateCaptions(doc).Select(c => (object)new
        {
            path = CaptionPath(c.Label, c.Number),
            label = c.Label,
            number = c.Number,
            text = c.Text,
            anchorPath = c.AnchorPath,
        })];
}
