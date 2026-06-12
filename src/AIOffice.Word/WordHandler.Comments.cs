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
    /// <c>{"op":"add","path":"&lt;content path&gt;","type":"comment","props":{"text":…,"author"?}}</c>:
    /// anchors commentRangeStart/End around the target paragraph or run, adds a
    /// commentReference run, and writes the entry into the comments part.
    /// </summary>
    private static object ApplyAddComment(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];

        foreach (var replyKey in (string[])["replyTo", "inReplyTo", "parent", "parentId"])
        {
            if (props.ContainsKey(replyKey))
            {
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"Comment replies (props.{replyKey}) are not supported yet (planned for M3).",
                    "Add a separate top-level comment on the same anchor instead.");
            }
        }

        var author = session.ResolveAuthor(props);
        var text = props["text"] is { } textNode ? NodeToString(textNode) : null;
        if (string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type comment needs props.text.",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[2]\",\"type\":\"comment\",\"props\":{\"text\":\"Please verify.\"}}.");
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not (Paragraph or Run))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Comments anchor to content (p or run), not '{anchor.Type}'.",
                anchor.Type is "tc"
                    ? $"Anchor to the paragraph inside the cell: {anchor.CanonicalPath}/p[1]."
                    : "Pick a paragraph or run path, e.g. /body/p[2] or /body/p[2]/run[1].");
        }

        var main = doc.MainDocumentPart!;
        var commentsPart = main.WordprocessingCommentsPart;
        if (commentsPart is null)
        {
            commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments();
        }

        var comments = commentsPart.Comments ??= new Comments();
        var id = comments.Elements<Comment>()
            .Select(c => int.TryParse(c.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0)
            .DefaultIfEmpty(0)
            .Max() + 1;
        var idText = id.ToString(CultureInfo.InvariantCulture);

        comments.AppendChild(new Comment(new Paragraph(new Run(NewText(text))))
        {
            Id = idText,
            Author = author,
            Date = DateTime.UtcNow,
        });

        var rangeStart = new CommentRangeStart { Id = idText };
        var rangeEnd = new CommentRangeEnd { Id = idText };
        var referenceRun = new Run(new CommentReference { Id = idText });

        if (anchor.Element is Paragraph paragraph)
        {
            // Range brackets the paragraph's content; pPr must stay first.
            if (paragraph.ParagraphProperties is { } pPr)
            {
                pPr.InsertAfterSelf(rangeStart);
            }
            else
            {
                paragraph.PrependChild(rangeStart);
            }

            paragraph.AppendChild(rangeEnd);
            paragraph.AppendChild(referenceRun);
        }
        else
        {
            var run = anchor.Element;
            run.InsertBeforeSelf(rangeStart);
            run.InsertAfterSelf(rangeEnd);
            rangeEnd.InsertAfterSelf(referenceRun);
        }

        return new { op = "add", type = "comment", path = CommentPath(id), anchor = anchor.CanonicalPath, author };
    }

    // ---------------------------------------------------------------- remove

    /// <summary>remove /comment[@id=N]: drops the part entry and every anchor marker.</summary>
    private static object ApplyRemoveComment(WordprocessingDocument doc, EditOp op)
    {
        var (comment, id) = ResolveComment(doc, DocPath.Parse(op.Path));

        var idText = id.ToString(CultureInfo.InvariantCulture);
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body is not null)
        {
            foreach (var start in body.Descendants<CommentRangeStart>().Where(s => s.Id?.Value == idText).ToList())
            {
                start.Remove();
            }

            foreach (var end in body.Descendants<CommentRangeEnd>().Where(e => e.Id?.Value == idText).ToList())
            {
                end.Remove();
            }

            foreach (var reference in body.Descendants<CommentReference>().Where(r => r.Id?.Value == idText).ToList())
            {
                // The conventional anchor is a run holding only the reference; drop it whole.
                if (reference.Parent is Run run && run.ChildElements.All(c => c is CommentReference or RunProperties))
                {
                    run.Remove();
                }
                else
                {
                    reference.Remove();
                }
            }
        }

        comment.Remove();
        return new { op = "remove", path = CommentPath(id), type = "comment" };
    }

    // ------------------------------------------------------------------ read

    private static object CommentsView(WordprocessingDocument doc)
    {
        var comments = EnumerateComments(doc)
            .Select(c => CommentShape(doc, c.Comment, c.Id))
            .ToList();
        return new { view = "comments", count = comments.Count, comments };
    }

    private static object CommentShape(WordprocessingDocument doc, Comment comment, int id)
    {
        var (anchorPath, anchorText) = CommentAnchor(doc, id);
        return new
        {
            path = CommentPath(id),
            id,
            author = comment.Author?.Value,
            date = comment.Date is { HasValue: true } d
                ? d.Value.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture)
                : null,
            text = comment.InnerText,
            anchorPath,
            anchorText,
        };
    }

    /// <summary>The paragraph holding the range start, and the text between the markers.</summary>
    private static (string? AnchorPath, string? AnchorText) CommentAnchor(WordprocessingDocument doc, int id)
    {
        var idText = id.ToString(CultureInfo.InvariantCulture);
        var body = doc.MainDocumentPart?.Document?.Body;
        var start = body?.Descendants<CommentRangeStart>().FirstOrDefault(s => s.Id?.Value == idText);
        if (start is null)
        {
            return (null, null);
        }

        var paragraph = start.Ancestors<Paragraph>().FirstOrDefault();
        string? anchorPath = null;
        if (paragraph is not null)
        {
            anchorPath = WordAddress.EnumerateAll(doc)
                .FirstOrDefault(n => ReferenceEquals(n.Element, paragraph))?.CanonicalPath;
        }

        var text = new System.Text.StringBuilder();
        for (var node = start.NextSibling(); node is not null; node = node.NextSibling())
        {
            if (node is CommentRangeEnd end && end.Id?.Value == idText)
            {
                break;
            }

            text.Append(node.InnerText);
        }

        return (anchorPath, text.ToString());
    }

    private static IEnumerable<(Comment Comment, int Id)> EnumerateComments(WordprocessingDocument doc) =>
        (doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments?.Elements<Comment>() ?? [])
        .Select(c => (c, int.TryParse(c.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0));

    /// <summary>Resolves /comment[@id=N] (or positional /comment[i]) or throws invalid_path with candidates.</summary>
    private static (Comment Comment, int Id) ResolveComment(WordprocessingDocument doc, DocPath path)
    {
        var comments = EnumerateComments(doc).ToList();
        var segment = path.Segments[0];

        if (path.Segments.Count == 1 && segment.Id is { } idValue &&
            int.TryParse(idValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id))
        {
            var byId = comments.FirstOrDefault(c => c.Id == id);
            if (byId.Comment is not null)
            {
                return byId;
            }

            throw CommentNotFound($"No comment has id {id}.", comments);
        }

        if (path.Segments.Count == 1 && segment.Id is null)
        {
            var index = segment.Index ?? 1;
            if (index <= comments.Count)
            {
                return comments[index - 1];
            }

            throw CommentNotFound(
                comments.Count == 0
                    ? "This document has no comments."
                    : $"/comment[{index}] does not exist; there are {comments.Count} comment(s).",
                comments);
        }

        throw CommentNotFound($"'{path.ToCanonicalString()}' is not a comment path.", comments);
    }

    private static AiofficeException CommentNotFound(string message, List<(Comment Comment, int Id)> comments) => new(
        ErrorCodes.InvalidPath,
        message,
        "Run 'aioffice read <file> --view comments' to list comments with their ids.",
        candidates: [.. comments.Take(5).Select(c => CommentPath(c.Id))]);

    private static string CommentPath(int id) =>
        string.Create(CultureInfo.InvariantCulture, $"/comment[@id={id}]");
}
