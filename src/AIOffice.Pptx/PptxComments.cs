using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Slide comments using the classic (pre-modern) comment parts for validator
/// compatibility: one deck-wide commentAuthors part (p:cmAuthorLst, deduplicated
/// by name) plus one comments part per slide (p:cmLst of p:cm). Comment ids
/// (p:cm/@idx) are kept globally unique so /slide[i]/comment[@id=N] is stable.
/// Replies are reserved for M5.
/// </summary>
internal static class PptxComments
{
    /// <summary>The author stamped on comments when props.author is absent.</summary>
    public const string DefaultAuthor = "AIOffice";

    private static readonly IReadOnlyList<string> AddProps = ["text", "author", "x", "y"];

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
                throw RepliesUnsupported();
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

    /// <summary>Replies are an M5 capability; adds targeting a comment path land here.</summary>
    public static AiofficeException RepliesUnsupported() => new(
        ErrorCodes.UnsupportedFeature,
        "Comment replies are not supported yet (planned M5).",
        "Add a top-level comment on the slide instead: {\"op\":\"add\",\"path\":\"/slide[2]\"," +
        "\"type\":\"comment\",\"props\":{\"text\":\"...\"}}.");

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

    /// <summary>The `get` projection for /slide[i]/comment[@id=N].</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var comment = Resolve(slidePart, address);
        return Project(presentation, comment, address.SlideIndex);
    }

    /// <summary>The `read --view comments` projection: every comment across the deck's slides.</summary>
    public static object CommentsView(PresentationPart presentation, IEnumerable<(int Index, SlidePart Part)> slides)
    {
        var rows = new List<object>();
        foreach (var (index, slidePart) in slides)
        {
            rows.AddRange(Comments(slidePart).Select(comment => Project(presentation, comment, index)));
        }

        return new { View = "comments", Count = rows.Count, Comments = rows };
    }

    private static object Project(PresentationPart presentation, P.Comment comment, int slideIndex)
    {
        var author = presentation.CommentAuthorsPart?.CommentAuthorList?
            .Elements<P.CommentAuthor>()
            .FirstOrDefault(a => a.Id?.Value == comment.AuthorId?.Value);
        var position = comment.GetFirstChild<P.Position>();
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
        };
    }

    // ----- remove ---------------------------------------------------------------------

    /// <summary>remove /slide[i]/comment[@id=N]: drops the comment (and the part when it empties).</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var comment = Resolve(slidePart, address);
        comment.Remove();

        if (slidePart.SlideCommentsPart is { } part &&
            (part.CommentList is null || !part.CommentList.Elements<P.Comment>().Any()))
        {
            slidePart.DeletePart(part);
        }

        return address.CanonicalCommentPath;
    }
}
