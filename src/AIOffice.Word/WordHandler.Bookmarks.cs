using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>Word's bookmark-name rules: letter/underscore start, no spaces, 40 chars max.</summary>
    [GeneratedRegex("^[A-Za-z_][A-Za-z0-9_]{0,39}$")]
    private static partial Regex BookmarkNamePattern();

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[3]","type":"bookmark","props":{"name":"Results"}}</c>:
    /// wraps the target paragraph in bookmarkStart/bookmarkEnd with a
    /// document-unique id, so internal links ({"anchor":"Results"}) can target it.
    /// </summary>
    private static object ApplyAddBookmark(WordprocessingDocument doc, EditOp op)
    {
        var name = op.Props?["name"] is { } nameNode ? NodeToString(nameNode) : null;
        if (string.IsNullOrEmpty(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type bookmark needs props.name.",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[3]\",\"type\":\"bookmark\",\"props\":{\"name\":\"Results\"}}.");
        }

        if (!BookmarkNamePattern().IsMatch(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{name}' is not a valid bookmark name.",
                "Bookmark names start with a letter or underscore, use letters/digits/underscores only, max 40 chars (no spaces).");
        }

        if (EnumerateBookmarks(doc).Any(b => string.Equals(b.Name?.Value, name, StringComparison.OrdinalIgnoreCase)))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Bookmark '{name}' already exists.",
                "Bookmark names are unique. Remove the old one first ({\"op\":\"remove\",\"path\":\"" +
                BookmarkPath(name) + "\"}) or pick a different name.");
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Bookmarks wrap paragraphs, not '{anchor.Type}'.",
                anchor.Type is "tc"
                    ? $"Target the paragraph inside the cell: {anchor.CanonicalPath}/p[1]."
                    : "Pick a paragraph path, e.g. /body/p[3].");
        }

        var id = NextBookmarkId(doc).ToString(CultureInfo.InvariantCulture);
        var start = new BookmarkStart { Id = id, Name = name };
        var end = new BookmarkEnd { Id = id };

        if (paragraph.ParagraphProperties is { } pPr)
        {
            pPr.InsertAfterSelf(start);
        }
        else
        {
            paragraph.PrependChild(start);
        }

        paragraph.AppendChild(end);
        return new { op = "add", type = "bookmark", path = BookmarkPath(name), name, anchor = anchor.CanonicalPath };
    }

    // ---------------------------------------------------------------- remove

    /// <summary>remove /bookmark[@name=X]: drops the start/end markers, content stays.</summary>
    private static object ApplyRemoveBookmark(WordprocessingDocument doc, EditOp op)
    {
        var start = ResolveBookmark(doc, DocPath.Parse(op.Path));
        var name = start.Name?.Value ?? string.Empty;
        var id = start.Id?.Value;

        foreach (var root in BookmarkRoots(doc))
        {
            foreach (var end in root.Descendants<BookmarkEnd>().Where(e => e.Id?.Value == id).ToList())
            {
                end.Remove();
            }
        }

        start.Remove();
        return new { op = "remove", path = BookmarkPath(name), type = "bookmark" };
    }

    // ------------------------------------------------------------------ read

    private static IEnumerable<OpenXmlElement> BookmarkRoots(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document?.Body is { } body)
        {
            yield return body;
        }

        foreach (var root in WordAddress.HeaderFooterRoots(doc))
        {
            yield return root.Element;
        }
    }

    /// <summary>All bookmark starts in document order (body first, then headers/footers).</summary>
    private static List<BookmarkStart> EnumerateBookmarks(WordprocessingDocument doc) =>
        [.. BookmarkRoots(doc).SelectMany(root => root.Descendants<BookmarkStart>())];

    private static int NextBookmarkId(WordprocessingDocument doc) =>
        EnumerateBookmarks(doc)
            .Select(b => int.TryParse(b.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0)
            .DefaultIfEmpty(0)
            .Max() + 1;

    /// <summary>get /bookmark[@name=X] data: where it anchors and what it wraps.</summary>
    private static Dictionary<string, object?> GetBookmarkProperties(WordprocessingDocument doc, DocPath path)
    {
        var start = ResolveBookmark(doc, path);
        return BookmarkShape(doc, start);
    }

    private static Dictionary<string, object?> BookmarkShape(WordprocessingDocument doc, BookmarkStart start)
    {
        var paragraph = start.Ancestors<Paragraph>().FirstOrDefault();
        var anchorPath = paragraph is null
            ? null
            : WordAddress.EnumerateAll(doc).FirstOrDefault(n => ReferenceEquals(n.Element, paragraph))?.CanonicalPath;

        return new Dictionary<string, object?>
        {
            ["name"] = start.Name?.Value,
            ["id"] = int.TryParse(start.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0,
            ["anchorPath"] = anchorPath,
            ["snippet"] = paragraph is null ? null : Snippet(paragraph.InnerText),
        };
    }

    /// <summary>Resolves /bookmark[@name=X] (or positional /bookmark[i]) or throws invalid_path with candidates.</summary>
    private static BookmarkStart ResolveBookmark(WordprocessingDocument doc, DocPath path)
    {
        var bookmarks = EnumerateBookmarks(doc);
        var segment = path.Segments[0];

        if (path.Segments.Count == 1 && segment.Id is { } wanted)
        {
            var byName = bookmarks.FirstOrDefault(b =>
                string.Equals(b.Name?.Value, wanted, StringComparison.OrdinalIgnoreCase));
            if (byName is not null)
            {
                return byName;
            }

            throw BookmarkNotFound($"No bookmark is named '{wanted}'.", bookmarks);
        }

        if (path.Segments.Count == 1 && segment.Id is null)
        {
            var index = segment.Index ?? 1;
            if (index <= bookmarks.Count)
            {
                return bookmarks[index - 1];
            }

            throw BookmarkNotFound(
                bookmarks.Count == 0
                    ? "This document has no bookmarks."
                    : $"/bookmark[{index}] does not exist; there are {bookmarks.Count} bookmark(s).",
                bookmarks);
        }

        throw BookmarkNotFound($"'{path.ToCanonicalString()}' is not a bookmark path.", bookmarks);
    }

    private static AiofficeException BookmarkNotFound(string message, List<BookmarkStart> bookmarks) => new(
        ErrorCodes.InvalidPath,
        message,
        "Run 'aioffice read <file> --view structure' to list bookmarks with their names.",
        candidates: [.. bookmarks.Take(5).Select(b => BookmarkPath(b.Name?.Value ?? string.Empty))]);

    private static string BookmarkPath(string name) => $"/bookmark[@name={name}]";
}
