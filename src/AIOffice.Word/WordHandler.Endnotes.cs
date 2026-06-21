using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    // M4 resolves the M3 debt: endnotes were an unsupported_feature refusal
    // (naming footnotes as the workaround) — they now share the footnote surface.

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[2]","type":"endnote","props":{"text":…}}</c>:
    /// writes the note into the endnotes part (separator/continuation defaults
    /// created on first use) and appends a superscript reference run to the
    /// target paragraph — the same surface as footnotes, rendered at the end of
    /// the document instead of the bottom of the page.
    /// </summary>
    private static object ApplyAddEndnote(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props?.DeepClone().AsObject();
        var author = session.ResolveAuthor(props);

        var text = props?["text"] is { } textNode ? NodeToString(textNode) : null;
        if (string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type endnote needs props.text.",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[2]\",\"type\":\"endnote\",\"props\":{\"text\":\"See the appendix for the full data set.\"}}.");
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph || !anchor.CanonicalPath.StartsWith("/body", StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Endnote references attach to body paragraphs, not '{anchor.Type}' at {anchor.CanonicalPath}.",
                anchor.Type is "tc"
                    ? $"Target the paragraph inside the cell: {anchor.CanonicalPath}/p[1]."
                    : "Pick a body paragraph path, e.g. /body/p[2]. Word does not allow endnotes in headers or footers.");
        }

        var endnotes = EnsureEndnotesRoot(doc);
        var id = EnumerateEndnotes(doc).Select(e => e.Id).DefaultIfEmpty(0).Max() + 1;

        endnotes.AppendChild(new Endnote(
            new Paragraph(
                new Run(
                    new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                    new EndnoteReferenceMark()),
                new Run(NewText(" " + text))))
        {
            Id = id,
        });

        var referenceRun = new Run(
            new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new EndnoteReference { Id = id });
        paragraph.AppendChild(referenceRun);

        // Tracked: wrap ONLY the newly-appended reference run in a w:ins (whose @w:id
        // comes from NextRevisionId — a separate id space from the endnote's @w:id).
        // We deliberately do NOT MarkParagraphInserted: the anchor's pre-existing runs
        // and paragraph mark are the author's, not part of this insertion.
        if (session.Track)
        {
            var ins = NewTrackChange(new InsertedRun(), author, DateTime.UtcNow, NextRevisionId(doc));
            referenceRun.InsertBeforeSelf(ins);
            referenceRun.Remove();
            ins.AppendChild(referenceRun);
        }

        return new { op = "add", type = "endnote", path = EndnotePath(id), anchor = anchor.CanonicalPath };
    }

    // ---------------------------------------------------------------- remove

    /// <summary>remove /endnote[@id=N]: drops the part entry and every reference run.</summary>
    private static object ApplyRemoveEndnote(WordprocessingDocument doc, EditOp op)
    {
        var (endnote, id) = ResolveEndnote(doc, DocPath.Parse(op.Path));

        var body = doc.MainDocumentPart?.Document?.Body;
        foreach (var reference in body?.Descendants<EndnoteReference>().Where(r => r.Id?.Value == id).ToList() ?? [])
        {
            if (reference.Parent is Run run && run.ChildElements.All(c => c is EndnoteReference or RunProperties))
            {
                run.Remove();
            }
            else
            {
                reference.Remove();
            }
        }

        endnote.Remove();
        return new { op = "remove", path = EndnotePath(id), type = "endnote" };
    }

    // ------------------------------------------------------------------ part

    /// <summary>The endnotes root with Word's separator/continuation defaults in place.</summary>
    private static Endnotes EnsureEndnotesRoot(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The document has no main part.",
            "Re-export the file from Word.");

        var part = main.EndnotesPart;
        if (part is null)
        {
            part = main.AddNewPart<EndnotesPart>();
            part.Endnotes = new Endnotes(
                new Endnote(new Paragraph(new Run(new SeparatorMark())))
                {
                    Type = FootnoteEndnoteValues.Separator,
                    Id = -1,
                },
                new Endnote(new Paragraph(new Run(new ContinuationSeparatorMark())))
                {
                    Type = FootnoteEndnoteValues.ContinuationSeparator,
                    Id = 0,
                });
        }

        return part.Endnotes ??= new Endnotes();
    }

    // ------------------------------------------------------------------ read

    /// <summary>Real (content) endnotes with their ids; separators are plumbing, not content.</summary>
    private static IEnumerable<(Endnote Endnote, int Id)> EnumerateEndnotes(WordprocessingDocument doc) =>
        (doc.MainDocumentPart?.EndnotesPart?.Endnotes?.Elements<Endnote>() ?? [])
        .Where(e => e.Type is null || e.Type.Value == FootnoteEndnoteValues.Normal)
        .Select(e => (e, (int)(e.Id?.Value ?? 0)));

    private static Dictionary<string, object?> GetEndnoteProperties(WordprocessingDocument doc, DocPath path)
    {
        var (endnote, id) = ResolveEndnote(doc, path);

        var body = doc.MainDocumentPart?.Document?.Body;
        var reference = body?.Descendants<EndnoteReference>().FirstOrDefault(r => r.Id?.Value == id);
        var paragraph = reference?.Ancestors<Paragraph>().FirstOrDefault();
        var anchorPath = paragraph is null
            ? null
            : WordAddress.EnumerateAll(doc).FirstOrDefault(n => ReferenceEquals(n.Element, paragraph))?.CanonicalPath;

        return new Dictionary<string, object?>
        {
            ["id"] = id,
            ["text"] = EndnoteText(endnote),
            ["anchorPath"] = anchorPath,
        };
    }

    private static string EndnoteText(Endnote endnote) => endnote.InnerText.TrimStart();

    /// <summary>Resolves /endnote[@id=N] (or positional /endnote[i]) or throws invalid_path with candidates.</summary>
    private static (Endnote Endnote, int Id) ResolveEndnote(WordprocessingDocument doc, DocPath path)
    {
        var endnotes = EnumerateEndnotes(doc).ToList();
        var segment = path.Segments[0];

        if (path.Segments.Count == 1 && segment.Id is { } idValue &&
            int.TryParse(idValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id))
        {
            var byId = endnotes.FirstOrDefault(e => e.Id == id);
            if (byId.Endnote is not null)
            {
                return byId;
            }

            throw EndnoteNotFound($"No endnote has id {id}.", endnotes);
        }

        if (path.Segments.Count == 1 && segment.Id is null)
        {
            var index = segment.Index ?? 1;
            if (index <= endnotes.Count)
            {
                return endnotes[index - 1];
            }

            throw EndnoteNotFound(
                endnotes.Count == 0
                    ? "This document has no endnotes."
                    : $"/endnote[{index}] does not exist; there are {endnotes.Count} endnote(s).",
                endnotes);
        }

        throw EndnoteNotFound($"'{path.ToCanonicalString()}' is not an endnote path.", endnotes);
    }

    private static AiofficeException EndnoteNotFound(string message, List<(Endnote Endnote, int Id)> endnotes) => new(
        ErrorCodes.InvalidPath,
        message,
        "Run 'aioffice read <file> --view text' to see [^en] endnote markers with their ids.",
        candidates: [.. endnotes.Take(5).Select(e => EndnotePath(e.Id))]);

    private static string EndnotePath(int id) =>
        string.Create(CultureInfo.InvariantCulture, $"/endnote[@id={id}]");
}
