using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    private const string W14Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string McNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

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
            if (props.TryGetPropertyValue(replyKey, out var parentNode))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Replies are their own op type, not a comment prop (props.{replyKey}).",
                    "Use {\"op\":\"add\",\"path\":\"/comment[@id=" + NodeToString(parentNode) +
                    "]\",\"type\":\"reply\",\"props\":{\"text\":\"…\"}}.");
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

        var comments = EnsureCommentsRoot(doc);
        var id = NextCommentId(comments);
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

    /// <summary>
    /// <c>{"op":"add","path":"/comment[@id=N]","type":"reply","props":{"text":…,"author"?}}</c>:
    /// adds a threaded reply — a regular comment whose range/reference sit on the
    /// parent's anchor, linked to the parent through w15 commentsExtended
    /// (w15:paraIdParent on the reply's w14:paraId).
    /// </summary>
    private static object ApplyAddCommentReply(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        var author = session.ResolveAuthor(props);
        var text = props["text"] is { } textNode ? NodeToString(textNode) : null;
        if (string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type reply needs props.text.",
                "Example: {\"op\":\"add\",\"path\":\"/comment[@id=1]\",\"type\":\"reply\",\"props\":{\"text\":\"Agreed.\"}}.");
        }

        var (parent, parentId) = ResolveComment(doc, DocPath.Parse(op.Path));
        var parentIdText = parentId.ToString(CultureInfo.InvariantCulture);

        var body = doc.MainDocumentPart?.Document?.Body;
        var parentStart = body?.Descendants<CommentRangeStart>().FirstOrDefault(s => s.Id?.Value == parentIdText);
        var parentEnd = body?.Descendants<CommentRangeEnd>().FirstOrDefault(e => e.Id?.Value == parentIdText);
        var parentReference = body?.Descendants<CommentReference>().FirstOrDefault(r => r.Id?.Value == parentIdText);
        if (parentStart is null || parentEnd is null || parentReference?.Parent is not Run parentReferenceRun)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Comment {parentId} has no anchor in the document body, so a reply cannot attach to it.",
                "Add a fresh comment on the content instead: {\"op\":\"add\",\"path\":\"/body/p[n]\",\"type\":\"comment\",\"props\":{\"text\":\"…\"}}.");
        }

        var comments = EnsureCommentsRoot(doc);
        var id = NextCommentId(comments);
        var idText = id.ToString(CultureInfo.InvariantCulture);

        comments.AppendChild(new Comment(new Paragraph(new Run(NewText(text))))
        {
            Id = idText,
            Author = author,
            Date = DateTime.UtcNow,
        });

        // The reply brackets the same content as its parent.
        parentStart.InsertAfterSelf(new CommentRangeStart { Id = idText });
        parentEnd.InsertAfterSelf(new CommentRangeEnd { Id = idText });
        parentReferenceRun.InsertAfterSelf(new Run(new CommentReference { Id = idText }));

        // w15 thread wiring: reply paraId -> parent paraId.
        var parentParaId = EnsureCommentParaId(doc, parent);
        var replyParaId = EnsureCommentParaId(doc, comments.Elements<Comment>().Last());
        var commentsEx = EnsureCommentsExRoot(doc);
        commentsEx.AppendChild(new W15.CommentEx
        {
            ParaId = replyParaId,
            ParaIdParent = parentParaId,
            Done = false,
        });

        return new { op = "add", type = "reply", path = CommentPath(id), parent = CommentPath(parentId), author };
    }

    // ----------------------------------------------------------- w15 plumbing

    private static Comments EnsureCommentsRoot(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart!;
        var part = main.WordprocessingCommentsPart;
        if (part is null)
        {
            part = main.AddNewPart<WordprocessingCommentsPart>();
            part.Comments = new Comments();
        }

        return part.Comments ??= new Comments();
    }

    private static int NextCommentId(Comments comments) =>
        comments.Elements<Comment>()
            .Select(c => int.TryParse(c.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

    private static W15.CommentsEx EnsureCommentsExRoot(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart!;
        var part = main.WordprocessingCommentsExPart;
        if (part is null)
        {
            part = main.AddNewPart<WordprocessingCommentsExPart>();
            part.CommentsEx = new W15.CommentsEx();
        }

        return part.CommentsEx ??= new W15.CommentsEx();
    }

    /// <summary>
    /// The w14:paraId of a comment's last paragraph (the thread key Word uses),
    /// assigned on demand. The comments root declares mc:Ignorable="w14" so
    /// down-level validation skips the attribute, exactly like Word's own files.
    /// </summary>
    private static string EnsureCommentParaId(WordprocessingDocument doc, Comment comment)
    {
        var paragraph = comment.Elements<Paragraph>().LastOrDefault();
        if (paragraph is null)
        {
            paragraph = new Paragraph();
            comment.AppendChild(paragraph);
        }

        if (paragraph.ParagraphId?.Value is { Length: > 0 } existing)
        {
            return existing;
        }

        var comments = (Comments)comment.Parent!;
        DeclareW14Ignorable(comments);

        var used = new HashSet<string>(
            comments.Descendants<Paragraph>()
                .Select(p => p.ParagraphId?.Value)
                .OfType<string>(),
            StringComparer.OrdinalIgnoreCase);

        // Any unique 8-hex value below 0x80000000 and non-zero is a valid paraId.
        var next = 0x10000000u + (uint)used.Count;
        string paraId;
        do
        {
            next++;
            paraId = next.ToString("X8", CultureInfo.InvariantCulture);
        }
        while (used.Contains(paraId));

        paragraph.ParagraphId = paraId;
        return paraId;
    }

    private static void DeclareW14Ignorable(Comments comments)
    {
        if (comments.LookupNamespace("w14") is null)
        {
            comments.AddNamespaceDeclaration("w14", W14Namespace);
        }

        if (comments.LookupNamespace("mc") is null)
        {
            comments.AddNamespaceDeclaration("mc", McNamespace);
        }

        var mc = comments.MCAttributes ??= new MarkupCompatibilityAttributes();
        var ignorable = mc.Ignorable?.Value ?? string.Empty;
        if (!ignorable.Split(' ', StringSplitOptions.RemoveEmptyEntries).Contains("w14", StringComparer.Ordinal))
        {
            mc.Ignorable = ignorable.Length == 0 ? "w14" : ignorable + " w14";
        }
    }

    /// <summary>comment id -> parent comment id, resolved through the w15 paraId links.</summary>
    private static Dictionary<int, int> CommentParentMap(WordprocessingDocument doc)
    {
        var map = new Dictionary<int, int>();
        var commentsEx = doc.MainDocumentPart?.WordprocessingCommentsExPart?.CommentsEx;
        if (commentsEx is null)
        {
            return map;
        }

        var byParaId = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var (comment, id) in EnumerateComments(doc))
        {
            if (comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value is { Length: > 0 } paraId)
            {
                byParaId[paraId] = id;
            }
        }

        foreach (var entry in commentsEx.Elements<W15.CommentEx>())
        {
            if (entry.ParaId?.Value is { } childParaId &&
                entry.ParaIdParent?.Value is { } parentParaId &&
                byParaId.TryGetValue(childParaId, out var childId) &&
                byParaId.TryGetValue(parentParaId, out var parentId))
            {
                map[childId] = parentId;
            }
        }

        return map;
    }

    // ---------------------------------------------------------------- remove

    /// <summary>
    /// remove /comment[@id=N]: drops the part entry and every anchor marker;
    /// removing a parent removes its replies (a thread cannot dangle).
    /// </summary>
    private static object ApplyRemoveComment(WordprocessingDocument doc, EditOp op)
    {
        var (comment, id) = ResolveComment(doc, DocPath.Parse(op.Path));

        var parents = CommentParentMap(doc);
        var doomed = new List<(Comment Comment, int Id)> { (comment, id) };
        for (var i = 0; i < doomed.Count; i++)
        {
            var currentId = doomed[i].Id;
            doomed.AddRange(EnumerateComments(doc).Where(c =>
                parents.TryGetValue(c.Id, out var parentId) && parentId == currentId &&
                doomed.All(d => d.Id != c.Id)));
        }

        foreach (var (target, targetId) in doomed)
        {
            RemoveOneComment(doc, target, targetId);
        }

        return new
        {
            op = "remove",
            path = CommentPath(id),
            type = "comment",
            removedReplies = doomed.Count > 1 ? (int?)(doomed.Count - 1) : null,
        };
    }

    private static void RemoveOneComment(WordprocessingDocument doc, Comment comment, int id)
    {
        var idText = id.ToString(CultureInfo.InvariantCulture);
        var paraId = comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;
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

        if (paraId is { Length: > 0 } &&
            doc.MainDocumentPart?.WordprocessingCommentsExPart?.CommentsEx is { } commentsEx)
        {
            foreach (var entry in commentsEx.Elements<W15.CommentEx>()
                .Where(e => string.Equals(e.ParaId?.Value, paraId, StringComparison.OrdinalIgnoreCase))
                .ToList())
            {
                entry.Remove();
            }
        }

        comment.Remove();
    }

    // ------------------------------------------------------------------ read

    /// <summary>Threaded view: top-level comments in id order, replies nested under their parent.</summary>
    private static object CommentsView(WordprocessingDocument doc)
    {
        var parents = CommentParentMap(doc);
        var all = EnumerateComments(doc).ToList();
        var comments = all
            .Where(c => !parents.ContainsKey(c.Id))
            .Select(c => CommentShape(doc, c.Comment, c.Id, parents, all))
            .ToList();

        return new { view = "comments", count = all.Count, comments };
    }

    private static object CommentShape(
        WordprocessingDocument doc,
        Comment comment,
        int id,
        Dictionary<int, int>? parents = null,
        List<(Comment Comment, int Id)>? all = null)
    {
        parents ??= CommentParentMap(doc);
        all ??= EnumerateComments(doc).ToList();

        var (anchorPath, anchorText) = CommentAnchor(doc, id);
        var replies = all
            .Where(c => parents.TryGetValue(c.Id, out var parentId) && parentId == id)
            .Select(c => CommentShape(doc, c.Comment, c.Id, parents, all))
            .ToList();

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
            parentId = parents.TryGetValue(id, out var parent) ? (int?)parent : null,
            replies = replies.Count > 0 ? replies : null,
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

            if (node is not (CommentRangeStart or CommentRangeEnd))
            {
                text.Append(node.InnerText);
            }
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
