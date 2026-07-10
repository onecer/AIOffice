using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    private static readonly string[] AllowedLinkSchemes = ["http", "https", "mailto"];

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[2]","type":"link","props":{"text":…,"url":…}}</c>:
    /// appends a relationship-based external hyperlink (Hyperlink-styled run)
    /// to the target paragraph. <c>props.anchor</c> instead of url makes an
    /// internal link to a bookmark.
    /// </summary>
    private static object ApplyAddLink(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author"); // batch attribution metadata, never link content

        var text = props["text"] is { } textNode ? NodeToString(textNode) : null;
        if (string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type link needs props.text (the visible link text).",
                "Example: {\"op\":\"add\",\"path\":\"/body/p[2]\",\"type\":\"link\",\"props\":{\"text\":\"AIOffice\",\"url\":\"https://example.com\"}}.");
        }

        var url = props["url"] is { } urlNode ? NodeToString(urlNode) : null;
        var bookmark = props["anchor"] is { } anchorNode ? NodeToString(anchorNode) : null;
        string? tooltip = null;
        if (props["tooltip"] is { } tooltipNode)
        {
            if (tooltipNode is not JsonValue tipValue ||
                !tipValue.TryGetValue<string>(out var tip) ||
                string.IsNullOrWhiteSpace(tip))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "props.tooltip must be a non-empty string (the hyperlink ScreenTip shown on hover).",
                    "Example: {\"op\":\"add\",\"path\":\"/body/p[2]\",\"type\":\"link\",\"props\":{\"text\":\"Docs\",\"url\":\"https://example.com\",\"tooltip\":\"Open the docs\"}}.");
            }

            tooltip = tip;
        }
        if ((url is null) == (bookmark is null))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                url is null
                    ? "add --type link needs props.url (external) or props.anchor (bookmark name)."
                    : "props.url and props.anchor are mutually exclusive; a link is external or internal.",
                "Pass url \"https://…\" / \"mailto:…\" for external links, or anchor \"BookmarkName\" for internal ones.");
        }

        var target = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (target.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Links are appended to paragraphs, not '{target.Type}'.",
                target.Type is "tc"
                    ? $"Target the paragraph inside the cell: {target.CanonicalPath}/p[1]."
                    : "Pick a paragraph path, e.g. /body/p[2].");
        }

        var hyperlink = new Hyperlink();
        if (url is not null)
        {
            if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) ||
                !AllowedLinkSchemes.Contains(uri.Scheme, StringComparer.OrdinalIgnoreCase))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"'{url}' is not a supported link target.",
                    "External links use absolute http(s) or mailto urls. For a spot inside the document, " +
                    "add a bookmark and link it with props.anchor.",
                    candidates: AllowedLinkSchemes);
            }

            var part = OwningPart(doc, paragraph) ?? doc.MainDocumentPart!;
            hyperlink.Id = part.AddHyperlinkRelationship(uri, isExternal: true).Id;
        }
        else
        {
            var names = EnumerateBookmarks(doc).Select(b => b.Name?.Value).OfType<string>().ToList();
            if (!names.Contains(bookmark!, StringComparer.OrdinalIgnoreCase))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"No bookmark is named '{bookmark}', so the internal link would go nowhere.",
                    "Add the bookmark first: {\"op\":\"add\",\"path\":\"/body/p[n]\",\"type\":\"bookmark\",\"props\":{\"name\":\"" + bookmark + "\"}}.",
                    candidates: [.. names.Take(5).Select(BookmarkPath)]);
            }

            hyperlink.Anchor = bookmark;
        }

        if (tooltip is not null)
        {
            hyperlink.Tooltip = tooltip; // w:hyperlink/@w:tooltip — the Word ScreenTip
        }

        EnsureHyperlinkStyle(doc);
        hyperlink.AppendChild(new Run(
            new RunProperties(new RunStyle { Val = "Hyperlink" }),
            NewText(text)));
        paragraph.AppendChild(hyperlink);

        var canonical = WordAddress.EnumerateAll(doc)
            .FirstOrDefault(n => ReferenceEquals(n.Element, hyperlink))?.CanonicalPath
            ?? target.CanonicalPath;

        return new { op = "add", type = "link", path = canonical, url, anchor = bookmark, text, tooltip };
    }

    /// <summary>The classic blue/underlined Hyperlink character style, defined once on demand.</summary>
    private static void EnsureHyperlinkStyle(WordprocessingDocument doc)
    {
        var styles = EnsureStylesRoot(doc);
        if (FindStyle(styles, "Hyperlink") is not null)
        {
            return;
        }

        styles.AppendChild(new Style(
            new StyleName { Val = "Hyperlink" },
            new StyleRunProperties(
                new Color { Val = "0563C1" },
                new Underline { Val = UnderlineValues.Single }))
        {
            Type = StyleValues.Character,
            StyleId = "Hyperlink",
        });
    }

    // ------------------------------------------------------------------ read

    /// <summary>get on a /body/p[n]/link[m] node.</summary>
    private static Dictionary<string, object?> LinkProperties(WordprocessingDocument doc, ResolvedNode node)
    {
        var hyperlink = (Hyperlink)node.Element;
        var props = new Dictionary<string, object?>
        {
            ["kind"] = "link",
            ["url"] = ResolveLinkUrl(doc, hyperlink),
            ["anchor"] = hyperlink.Anchor?.Value,
            ["text"] = hyperlink.InnerText,
        };

        // Conditional: a null value in a Dictionary is NOT dropped by DefaultIgnoreCondition,
        // so add the key only when a real ScreenTip exists — legacy links stay tooltip-free.
        if (hyperlink.Tooltip?.Value is { Length: > 0 } tip)
        {
            props["tooltip"] = tip;
        }

        return props;
    }

    /// <summary>The external url behind a hyperlink's r:id, resolved on its owning part.</summary>
    private static string? ResolveLinkUrl(WordprocessingDocument doc, Hyperlink hyperlink)
    {
        if (hyperlink.Id?.Value is not { Length: > 0 } relId)
        {
            return null;
        }

        var part = OwningPart(doc, hyperlink);
        var relationship = part?.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId);
        return relationship?.Uri.OriginalString;
    }

    // ---------------------------------------------------------------- remove

    /// <summary>Removes a hyperlink element and, when unreferenced, its relationship.</summary>
    private static void RemoveLink(WordprocessingDocument doc, Hyperlink hyperlink)
    {
        var relId = hyperlink.Id?.Value;
        var part = OwningPart(doc, hyperlink);
        hyperlink.Remove();

        if (relId is { Length: > 0 } && part is not null &&
            part.RootElement?.Descendants<Hyperlink>().All(h => h.Id?.Value != relId) == true &&
            part.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId) is { } relationship)
        {
            part.DeleteReferenceRelationship(relationship);
        }
    }
}
