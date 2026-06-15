using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[2]","type":"footnote","props":{"text":…}}</c>:
    /// writes the note into the footnotes part (separator/continuation defaults
    /// created on first use) and appends a superscript reference run to the
    /// target paragraph.
    /// </summary>
    private static object ApplyAddFootnote(WordprocessingDocument doc, EditOp op)
    {
        var text = op.Props?["text"] is { } textNode ? NodeToString(textNode) : null;
        if (string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type footnote needs props.text.",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[2]\",\"type\":\"footnote\",\"props\":{\"text\":\"Source: 2025 annual report.\"}}.");
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph || !anchor.CanonicalPath.StartsWith("/body", StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Footnote references attach to body paragraphs, not '{anchor.Type}' at {anchor.CanonicalPath}.",
                anchor.Type is "tc"
                    ? $"Target the paragraph inside the cell: {anchor.CanonicalPath}/p[1]."
                    : "Pick a body paragraph path, e.g. /body/p[2]. Word does not allow footnotes in headers or footers.");
        }

        var footnotes = EnsureFootnotesRoot(doc);
        var id = EnumerateFootnotes(doc).Select(f => f.Id).DefaultIfEmpty(0).Max() + 1;

        footnotes.AppendChild(new Footnote(
            new Paragraph(
                new Run(
                    new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                    new FootnoteReferenceMark()),
                new Run(NewText(" " + text))))
        {
            Id = id,
        });

        paragraph.AppendChild(new Run(
            new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new FootnoteReference { Id = id }));

        return new { op = "add", type = "footnote", path = FootnotePath(id), anchor = anchor.CanonicalPath };
    }

    // The EndnotesUnsupported() refusal that lived here through M3 is gone:
    // endnotes are a real feature since M4 (see WordHandler.Endnotes.cs).

    // ---------------------------------------------------------------- remove

    /// <summary>remove /footnote[@id=N]: drops the part entry and every reference run.</summary>
    private static object ApplyRemoveFootnote(WordprocessingDocument doc, EditOp op)
    {
        var (footnote, id) = ResolveFootnote(doc, DocPath.Parse(op.Path));

        var body = doc.MainDocumentPart?.Document?.Body;
        foreach (var reference in body?.Descendants<FootnoteReference>().Where(r => r.Id?.Value == id).ToList() ?? [])
        {
            if (reference.Parent is Run run && run.ChildElements.All(c => c is FootnoteReference or RunProperties))
            {
                run.Remove();
            }
            else
            {
                reference.Remove();
            }
        }

        footnote.Remove();
        return new { op = "remove", path = FootnotePath(id), type = "footnote" };
    }

    // ------------------------------------------------------------------ part

    /// <summary>The footnotes root with Word's separator/continuation defaults in place.</summary>
    private static Footnotes EnsureFootnotesRoot(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The document has no main part.",
            "Re-export the file from Word.");

        var part = main.FootnotesPart;
        if (part is null)
        {
            part = main.AddNewPart<FootnotesPart>();
            part.Footnotes = new Footnotes(
                new Footnote(new Paragraph(new Run(new SeparatorMark())))
                {
                    Type = FootnoteEndnoteValues.Separator,
                    Id = -1,
                },
                new Footnote(new Paragraph(new Run(new ContinuationSeparatorMark())))
                {
                    Type = FootnoteEndnoteValues.ContinuationSeparator,
                    Id = 0,
                });
        }

        return part.Footnotes ??= new Footnotes();
    }

    // ------------------------------------------------------------------ read

    /// <summary>Real (content) footnotes with their ids; separators are plumbing, not content.</summary>
    private static IEnumerable<(Footnote Footnote, int Id)> EnumerateFootnotes(WordprocessingDocument doc) =>
        (doc.MainDocumentPart?.FootnotesPart?.Footnotes?.Elements<Footnote>() ?? [])
        .Where(f => f.Type is null || f.Type.Value == FootnoteEndnoteValues.Normal)
        .Select(f => (f, (int)(f.Id?.Value ?? 0)));

    private static Dictionary<string, object?> GetFootnoteProperties(WordprocessingDocument doc, DocPath path)
    {
        var (footnote, id) = ResolveFootnote(doc, path);
        return FootnoteShape(doc, footnote, id);
    }

    private static Dictionary<string, object?> FootnoteShape(WordprocessingDocument doc, Footnote footnote, int id)
    {
        var idValue = id;
        var body = doc.MainDocumentPart?.Document?.Body;
        var reference = body?.Descendants<FootnoteReference>().FirstOrDefault(r => r.Id?.Value == idValue);
        var paragraph = reference?.Ancestors<Paragraph>().FirstOrDefault();
        var anchorPath = paragraph is null
            ? null
            : WordAddress.EnumerateAll(doc).FirstOrDefault(n => ReferenceEquals(n.Element, paragraph))?.CanonicalPath;

        return new Dictionary<string, object?>
        {
            ["id"] = id,
            ["text"] = FootnoteText(footnote),
            ["anchorPath"] = anchorPath,
        };
    }

    private static string FootnoteText(Footnote footnote) => footnote.InnerText.TrimStart();

    /// <summary>Resolves /footnote[@id=N] (or positional /footnote[i]) or throws invalid_path with candidates.</summary>
    private static (Footnote Footnote, int Id) ResolveFootnote(WordprocessingDocument doc, DocPath path)
    {
        var footnotes = EnumerateFootnotes(doc).ToList();
        var segment = path.Segments[0];

        if (path.Segments.Count == 1 && segment.Id is { } idValue &&
            int.TryParse(idValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id))
        {
            var byId = footnotes.FirstOrDefault(f => f.Id == id);
            if (byId.Footnote is not null)
            {
                return byId;
            }

            throw FootnoteNotFound($"No footnote has id {id}.", footnotes);
        }

        if (path.Segments.Count == 1 && segment.Id is null)
        {
            var index = segment.Index ?? 1;
            if (index <= footnotes.Count)
            {
                return footnotes[index - 1];
            }

            throw FootnoteNotFound(
                footnotes.Count == 0
                    ? "This document has no footnotes."
                    : $"/footnote[{index}] does not exist; there are {footnotes.Count} footnote(s).",
                footnotes);
        }

        throw FootnoteNotFound($"'{path.ToCanonicalString()}' is not a footnote path.", footnotes);
    }

    private static AiofficeException FootnoteNotFound(string message, List<(Footnote Footnote, int Id)> footnotes) => new(
        ErrorCodes.InvalidPath,
        message,
        "Run 'aioffice read <file> --view text' to see [^n] footnote markers with their ids.",
        candidates: [.. footnotes.Take(5).Select(f => FootnotePath(f.Id))]);

    private static string FootnotePath(int id) =>
        string.Create(CultureInfo.InvariantCulture, $"/footnote[@id={id}]");

    // ------------------------------------------------------------- text view

    /// <summary>
    /// InnerText-compatible paragraph text with list markers prefixed, [^n]
    /// footnote markers, [^eN] endnote markers, and equations rendered as their
    /// stored LaTeX (<c>$latex$</c> inline, <c>$$latex$$</c> for a display block).
    /// </summary>
    private static string DisplayText(WordprocessingDocument doc, OpenXmlElement element)
    {
        var sb = new System.Text.StringBuilder();
        if (element is Paragraph paragraph)
        {
            sb.Append(ListMarker(doc, paragraph));
        }

        // Equations carry their own LaTeX leaves; render them as $…$/$$…$$ and skip
        // their descendant math text so the glyphs don't leak in twice.
        AppendDisplayText(sb, element);
        return sb.ToString();
    }

    private static void AppendDisplayText(System.Text.StringBuilder sb, OpenXmlElement element)
    {
        foreach (var child in element.ChildElements)
        {
            switch (child)
            {
                case DocumentFormat.OpenXml.Math.OfficeMath inlineMath:
                    sb.Append('$').Append(EquationLatex(inlineMath)).Append('$');
                    break;

                case DocumentFormat.OpenXml.Math.Paragraph displayMath:
                    foreach (var oMath in displayMath.ChildElements.OfType<DocumentFormat.OpenXml.Math.OfficeMath>())
                    {
                        sb.Append("$$").Append(EquationLatex(oMath)).Append("$$");
                    }

                    break;

                case FootnoteReference reference:
                    sb.Append("[^").Append(reference.Id?.Value ?? 0).Append(']');
                    break;

                case EndnoteReference reference:
                    sb.Append("[^e").Append(reference.Id?.Value ?? 0).Append(']');
                    break;

                // A complex field's instruction code (w:instrText) is never the
                // displayed text — only its cached result (between separate and end)
                // shows. Skipping it keeps e.g. an IF/MERGEFIELD instruction out of
                // the text view, matching what Word renders.
                case FieldCode:
                    break;

                case OpenXmlLeafTextElement leaf:
                    sb.Append(leaf.Text);
                    break;

                default:
                    AppendDisplayText(sb, child);
                    break;
            }
        }
    }
}
