using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;

namespace AIOffice.Pptx;

/// <summary>
/// Slide comments using the classic (pre-modern) comment parts for validator
/// compatibility: one deck-wide commentAuthors part (p:cmAuthorLst, deduplicated
/// by name) plus one comments part per slide (p:cmLst of p:cm). Comment ids
/// (p:cm/@idx) are kept globally unique so /slide[i]/comment[@id=N] is stable.
/// Replies are threaded the way PowerPoint 2013+ does it on classic parts: a
/// reply is a p:cm whose extLst carries p15:threadingInfo/p15:parentCm pointing
/// at the parent's authorId + idx.
/// </summary>
internal static class PptxComments
{
    /// <summary>The author stamped on comments when props.author is absent.</summary>
    public const string DefaultAuthor = "AIOffice";

    private static readonly IReadOnlyList<string> AddProps = ["text", "author", "x", "y"];

    private static readonly IReadOnlyList<string> ReplyProps = ["text", "author"];

    /// <summary>PowerPoint's extension uri for p15 comment threading info.</summary>
    private const string ThreadingExtensionUri = "{C676402C-5697-4E1C-873F-D02D1690AC5C}";

    // ----- add ---------------------------------------------------------------------

    /// <summary>add /slide[i] type:comment — appends one comment and returns /slide[i]/comment[@id=N].</summary>
    public static string Add(PresentationPart presentation, PptxAddress address, JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (string.Equals(key, "replyTo", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(key, "parent", StringComparison.OrdinalIgnoreCase))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Comments take no '{key}' prop; replies target the parent comment's path.",
                    "Use {\"op\":\"add\",\"path\":\"/slide[2]/comment[@id=1]\",\"type\":\"reply\"," +
                    "\"props\":{\"text\":\"...\"}}.");
            }

            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown comment prop '{key}'.",
                    "Comment props: text (required), author, x, y.",
                    candidates: AddProps);
            }
        }

        if (!props.TryGetPropertyValue("text", out var textNode) || textNode is null ||
            J.ScalarText(textNode).Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add comment requires props.text.",
                "Pass the comment text, e.g. {\"op\":\"add\",\"path\":\"/slide[2]\",\"type\":\"comment\"," +
                "\"props\":{\"text\":\"Tighten this slide\",\"author\":\"Dana\"}}.");
        }

        var text = J.ScalarText(textNode);
        var authorName = props.TryGetPropertyValue("author", out var authorNode) && authorNode is not null &&
            J.ScalarText(authorNode).Trim().Length > 0
                ? J.ScalarText(authorNode).Trim()
                : DefaultAuthor;
        var x = props.TryGetPropertyValue("x", out var xNode) ? Units.ParseLengthEmu("x", xNode) : 0L;
        var y = props.TryGetPropertyValue("y", out var yNode) ? Units.ParseLengthEmu("y", yNode) : 0L;

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var author = EnsureAuthor(presentation, authorName);
        var index = NextCommentIndex(presentation);

        var commentsPart = slidePart.SlideCommentsPart ?? slidePart.AddNewPart<SlideCommentsPart>();
        commentsPart.CommentList ??= new P.CommentList();
        commentsPart.CommentList.Append(new P.Comment(
            new P.Position { X = x, Y = y },
            new P.Text(text))
        {
            AuthorId = author.Id!.Value,
            DateTime = DateTime.UtcNow,
            Index = index,
        });

        if ((author.LastIndex?.Value ?? 0) < index)
        {
            author.LastIndex = index;
        }

        return Units.Inv($"/slide[{address.SlideIndex}]/comment[@id={index}]");
    }

    /// <summary>
    /// add /slide[i]/comment[@id=N] type:reply — appends a threaded reply (a
    /// p:cm carrying p15:parentCm) and returns its own comment path.
    /// </summary>
    public static string AddReply(PresentationPart presentation, PptxAddress address, JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!ReplyProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown reply prop '{key}'.",
                    "Reply props: text (required), author. Replies sit at their parent's anchor.",
                    candidates: ReplyProps);
            }
        }

        if (!props.TryGetPropertyValue("text", out var textNode) || textNode is null ||
            J.ScalarText(textNode).Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add reply requires props.text.",
                "Pass the reply text, e.g. {\"op\":\"add\",\"path\":\"" + address.CanonicalCommentPath +
                "\",\"type\":\"reply\",\"props\":{\"text\":\"Agreed\",\"author\":\"Riley\"}}.");
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var parent = Resolve(slidePart, address);
        if (ParentIdOf(parent) is { } grandparent)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{address.CanonicalCommentPath}' is itself a reply; threads are one level deep.",
                Units.Inv($"Reply to the thread's root instead: {address.CanonicalSlidePath}/comment[@id={grandparent}]."));
        }

        var authorName = props.TryGetPropertyValue("author", out var authorNode) && authorNode is not null &&
            J.ScalarText(authorNode).Trim().Length > 0
                ? J.ScalarText(authorNode).Trim()
                : DefaultAuthor;
        var author = EnsureAuthor(presentation, authorName);
        var index = NextCommentIndex(presentation);
        var parentPosition = parent.GetFirstChild<P.Position>();

        var reply = new P.Comment(
            new P.Position { X = parentPosition?.X?.Value ?? 0L, Y = parentPosition?.Y?.Value ?? 0L },
            new P.Text(J.ScalarText(textNode)),
            new P.CommentExtensionList(new P.CommentExtension(
                new P15.ThreadingInfo(new P15.ParentCommentIdentifier
                {
                    AuthorId = parent.AuthorId?.Value ?? 0,
                    Index = parent.Index?.Value ?? 0,
                }))
            {
                Uri = ThreadingExtensionUri,
            }))
        {
            AuthorId = author.Id!.Value,
            DateTime = DateTime.UtcNow,
            Index = index,
        };

        slidePart.SlideCommentsPart!.CommentList!.Append(reply);
        if ((author.LastIndex?.Value ?? 0) < index)
        {
            author.LastIndex = index;
        }

        return Units.Inv($"/slide[{address.SlideIndex}]/comment[@id={index}]");
    }

    /// <summary>The parent comment idx a reply points at (p15:parentCm); null for root comments.</summary>
    public static uint? ParentIdOf(P.Comment comment) =>
        comment.Descendants<P15.ParentCommentIdentifier>().FirstOrDefault()?.Index?.Value;

    /// <summary>The deck's author entry for a name (case-insensitive dedup), created on first use.</summary>
    private static P.CommentAuthor EnsureAuthor(PresentationPart presentation, string name)
    {
        var authorsPart = presentation.CommentAuthorsPart ?? presentation.AddNewPart<CommentAuthorsPart>();
        authorsPart.CommentAuthorList ??= new P.CommentAuthorList();

        var authors = authorsPart.CommentAuthorList.Elements<P.CommentAuthor>().ToList();
        if (authors.FirstOrDefault(a => string.Equals(a.Name?.Value, name, StringComparison.OrdinalIgnoreCase)) is { } existing)
        {
            return existing;
        }

        var author = new P.CommentAuthor
        {
            Id = authors.Count == 0 ? 0u : authors.Max(a => a.Id?.Value ?? 0u) + 1,
            Name = name,
            Initials = Initials(name),
            LastIndex = 0u,
            ColorIndex = (uint)authors.Count,
        };
        authorsPart.CommentAuthorList.Append(author);
        return author;
    }

    /// <summary>Up to three initial letters of the name's words ("Dana Q. Reviewer" → "DQR").</summary>
    private static string Initials(string name)
    {
        var letters = name
            .Split([' ', '\t', '.'], StringSplitOptions.RemoveEmptyEntries)
            .Take(3)
            .Select(word => char.ToUpperInvariant(word[0]));
        var initials = string.Concat(letters);
        return initials.Length == 0 ? "A" : initials;
    }

    /// <summary>The next globally unique comment idx (per-author uniqueness follows for free).</summary>
    private static uint NextCommentIndex(PresentationPart presentation)
    {
        var max = 0u;
        foreach (var author in presentation.CommentAuthorsPart?.CommentAuthorList?.Elements<P.CommentAuthor>() ?? [])
        {
            if (author.LastIndex?.Value is { } last && last > max)
            {
                max = last;
            }
        }

        foreach (var (_, slidePart) in PptxDoc.Slides(presentation))
        {
            foreach (var comment in Comments(slidePart))
            {
                if (comment.Index?.Value is { } index && index > max)
                {
                    max = index;
                }
            }
        }

        return max + 1;
    }

    // ----- read side -----------------------------------------------------------------

    /// <summary>The slide's comments in document order; empty when the slide has no comments part.</summary>
    public static List<P.Comment> Comments(SlidePart slidePart) =>
        [.. slidePart.SlideCommentsPart?.CommentList?.Elements<P.Comment>() ?? []];

    /// <summary>Resolves /slide[i]/comment[@id=N] or throws invalid_path with candidates.</summary>
    public static P.Comment Resolve(SlidePart slidePart, PptxAddress address)
    {
        var comments = Comments(slidePart);
        var id = address.CommentId!.Value;
        if (comments.FirstOrDefault(c => c.Index?.Value == id) is { } match)
        {
            return match;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            Units.Inv($"No comment @id={id} on slide {address.SlideIndex}; it has {comments.Count} comment(s)."),
            comments.Count > 0
                ? "Run 'aioffice read <file> --view comments' to list comment ids."
                : "Add one first: {\"op\":\"add\",\"path\":\"" + address.CanonicalSlidePath +
                  "\",\"type\":\"comment\",\"props\":{\"text\":\"...\"}}.",
            candidates: [.. comments.Take(10).Select(c =>
                Units.Inv($"{address.CanonicalSlidePath}/comment[@id={c.Index?.Value}]"))]);
    }

    /// <summary>The replies threaded under a comment, in document (idx) order.</summary>
    public static List<P.Comment> RepliesOf(SlidePart slidePart, uint parentId) =>
        [.. Comments(slidePart).Where(c => ParentIdOf(c) == parentId)];

    /// <summary>The `get` projection for /slide[i]/comment[@id=N].</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var comment = Resolve(slidePart, address);
        return Project(presentation, slidePart, comment, address.SlideIndex);
    }

    /// <summary>The `read --view comments` projection: root comments per slide with nested replies.</summary>
    public static object CommentsView(PresentationPart presentation, IEnumerable<(int Index, SlidePart Part)> slides)
    {
        var rows = new List<object>();
        foreach (var (index, slidePart) in slides)
        {
            rows.AddRange(Comments(slidePart)
                .Where(comment => ParentIdOf(comment) is null)
                .Select(comment => Project(presentation, slidePart, comment, index)));
        }

        return new { View = "comments", Count = rows.Count, Comments = rows };
    }

    private static object Project(PresentationPart presentation, SlidePart slidePart, P.Comment comment, int slideIndex)
    {
        var author = presentation.CommentAuthorsPart?.CommentAuthorList?
            .Elements<P.CommentAuthor>()
            .FirstOrDefault(a => a.Id?.Value == comment.AuthorId?.Value);
        var position = comment.GetFirstChild<P.Position>();
        var parentId = ParentIdOf(comment);
        List<P.Comment> replies = parentId is null && comment.Index?.Value is { } id
            ? RepliesOf(slidePart, id)
            : [];
        return new
        {
            Path = Units.Inv($"/slide[{slideIndex}]/comment[@id={comment.Index?.Value}]"),
            Slide = slideIndex,
            Id = comment.Index?.Value,
            Author = author?.Name?.Value,
            AuthorInitials = author?.Initials?.Value,
            Text = comment.GetFirstChild<P.Text>()?.Text ?? string.Empty,
            X = position?.X?.Value is { } x ? Units.EmuToCm(x) : (double?)null,
            Y = position?.Y?.Value is { } y ? Units.EmuToCm(y) : (double?)null,
            Date = comment.DateTime?.Value.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture),
            ParentId = parentId,
            ParentPath = parentId is { } pid ? Units.Inv($"/slide[{slideIndex}]/comment[@id={pid}]") : null,
            Replies = replies.Count == 0
                ? null
                : replies.Select(reply => Project(presentation, slidePart, reply, slideIndex)).ToList(),
        };
    }

    // ----- remove ---------------------------------------------------------------------

    /// <summary>
    /// remove /slide[i]/comment[@id=N]: drops the comment — replies individually,
    /// roots together with their replies (and the part when it empties).
    /// </summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var comment = Resolve(slidePart, address);
        if (ParentIdOf(comment) is null && comment.Index?.Value is { } id)
        {
            foreach (var reply in RepliesOf(slidePart, id))
            {
                reply.Remove(); // a root takes its thread with it; replies never orphan
            }
        }

        comment.Remove();

        if (slidePart.SlideCommentsPart is { } part &&
            (part.CommentList is null || !part.CommentList.Elements<P.Comment>().Any()))
        {
            slidePart.DeletePart(part);
        }

        return address.CanonicalCommentPath;
    }
}
