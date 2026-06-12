using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>An element resolved from a document path.</summary>
internal sealed record ResolvedNode(OpenXmlElement Element, string CanonicalPath, string Type);

/// <summary>
/// Resolves the docx addressing grammar (/body/p[3], /body/table[1]/tr[2]/tc[1],
/// /body/p[3]/run[2], /header[1]/p[1], /footer[1]/p[1]) against a live document,
/// and enumerates all addressable nodes with their canonical paths (the query
/// backbone). Every resolution failure is <c>invalid_path</c> WITH nearest-match
/// candidates.
/// </summary>
internal static class WordAddress
{
    /// <summary>Element type names addressable in a docx, child-of-parent rules included.</summary>
    private static readonly Dictionary<string, string[]> ChildNames = new(StringComparer.Ordinal)
    {
        ["body"] = ["p", "table"],
        ["header"] = ["p", "table"],
        ["footer"] = ["p", "table"],
        ["p"] = ["run", "link"],
        ["table"] = ["tr"],
        ["tr"] = ["tc"],
        ["tc"] = ["p", "table"],
        ["run"] = [],
        ["link"] = [],
    };

    public static ResolvedNode Resolve(WordprocessingDocument doc, DocPath path)
    {
        var segments = path.Segments;
        var (current, currentPath, currentType) = ResolveRoot(doc, segments[0]);

        for (var i = 1; i < segments.Count; i++)
        {
            (current, currentPath, currentType) = ResolveChild(current, currentPath, currentType, segments[i]);
        }

        return new ResolvedNode(current, currentPath, currentType);
    }

    private static (OpenXmlElement El, string Path, string Type) ResolveRoot(WordprocessingDocument doc, PathSegment root)
    {
        if (root.Kind != PathSegmentKind.Element)
        {
            throw Invalid($"A docx path starts with /body, /header[n] or /footer[n], not '{root.ToCanonicalString()}'.", RootCandidates(doc));
        }

        switch (root.Name)
        {
            case "body" when root.Index is null:
                var body = doc.MainDocumentPart?.Document?.Body
                    ?? throw new AiofficeException(
                        ErrorCodes.FormatCorrupt,
                        "The document has no body.",
                        "The main document part is missing or empty. Re-export the file from Word.");
                return (body, "/body", "body");

            case "header":
            {
                if (root.Variant is { } headerVariant)
                {
                    return ResolveVariantRoot(doc, "header", headerVariant);
                }

                var headers = doc.MainDocumentPart?.HeaderParts.ToList() ?? [];
                var index = root.Index ?? 1;
                if (headers.Count == 0)
                {
                    throw Invalid(
                        "This document has no headers. Create one with " +
                        "{\"op\":\"add\",\"path\":\"/header[1]\",\"type\":\"header\",\"props\":{\"text\":\"…\"}}.",
                        ["/body"]);
                }

                if (index > headers.Count)
                {
                    throw Invalid(
                        $"/header[{index}] does not exist; the document has {headers.Count} header(s).",
                        [.. Enumerable.Range(1, headers.Count).Select(n => Canon(string.Empty, "header", n))]);
                }

                var header = headers[index - 1].Header ?? throw new AiofficeException(
                    ErrorCodes.FormatCorrupt,
                    $"/header[{index}] exists but its part is empty.",
                    "Re-export the file from Word, or address /body instead.");
                return (header, $"/header[{index}]", "header");
            }

            case "footer":
            {
                if (root.Variant is { } footerVariant)
                {
                    return ResolveVariantRoot(doc, "footer", footerVariant);
                }

                var footers = doc.MainDocumentPart?.FooterParts.ToList() ?? [];
                var index = root.Index ?? 1;
                if (footers.Count == 0)
                {
                    throw Invalid(
                        "This document has no footers. Create one with " +
                        "{\"op\":\"add\",\"path\":\"/footer[1]\",\"type\":\"footer\",\"props\":{\"text\":\"…\"}}.",
                        ["/body"]);
                }

                if (index > footers.Count)
                {
                    throw Invalid(
                        $"/footer[{index}] does not exist; the document has {footers.Count} footer(s).",
                        [.. Enumerable.Range(1, footers.Count).Select(n => Canon(string.Empty, "footer", n))]);
                }

                var footer = footers[index - 1].Footer ?? throw new AiofficeException(
                    ErrorCodes.FormatCorrupt,
                    $"/footer[{index}] exists but its part is empty.",
                    "Re-export the file from Word, or address /body instead.");
                return (footer, $"/footer[{index}]", "footer");
            }

            default:
                throw Invalid($"Unknown docx root '{root.ToCanonicalString()}'; paths start at /body, /header[n] or /footer[n].", RootCandidates(doc));
        }
    }

    // ------------------------------------------------------- named variants

    /// <summary>The named header/footer variants of the addressing grammar (M5).</summary>
    internal static readonly string[] HeaderFooterVariants = ["default", "firstPage", "even"];

    /// <summary>Maps a variant name to its OOXML header/footer reference type.</summary>
    internal static HeaderFooterValues VariantReferenceType(string variant) => variant switch
    {
        "firstPage" => HeaderFooterValues.First,
        "even" => HeaderFooterValues.Even,
        _ => HeaderFooterValues.Default,
    };

    /// <summary>
    /// The header/footer part a named variant resolves to: the first
    /// w:headerReference / w:footerReference of that type in any sectPr,
    /// in document order. Null when no section references one.
    /// </summary>
    internal static OpenXmlPart? FindReferencedPart(WordprocessingDocument doc, string kind, string variant)
    {
        var main = doc.MainDocumentPart;
        if (main?.Document?.Body is not { } body)
        {
            return null;
        }

        var type = VariantReferenceType(variant);
        foreach (var sectPr in body.Descendants<SectionProperties>())
        {
            var relId = kind == "header"
                ? sectPr.Elements<HeaderReference>()
                    .FirstOrDefault(r => (r.Type?.Value ?? HeaderFooterValues.Default) == type)?.Id?.Value
                : sectPr.Elements<FooterReference>()
                    .FirstOrDefault(r => (r.Type?.Value ?? HeaderFooterValues.Default) == type)?.Id?.Value;

            if (relId is { Length: > 0 })
            {
                try
                {
                    return main.GetPartById(relId);
                }
                catch (ArgumentOutOfRangeException)
                {
                    // Dangling reference: keep scanning later sections.
                }
            }
        }

        return null;
    }

    /// <summary>The variant name ("default" | "firstPage" | "even") a part is referenced as, or null when unreferenced.</summary>
    internal static string? HeaderFooterVariantOf(WordprocessingDocument doc, OpenXmlPart part) =>
        HeaderFooterVariants.FirstOrDefault(variant =>
            ReferenceEquals(FindReferencedPart(doc, part is HeaderPart ? "header" : "footer", variant), part));

    /// <summary>Named paths of the variants that exist for one kind (e.g. ["/header[default]", "/header[even]"]).</summary>
    internal static IReadOnlyList<string> ExistingVariantPaths(WordprocessingDocument doc, string kind) =>
        [.. HeaderFooterVariants
            .Where(variant => FindReferencedPart(doc, kind, variant) is not null)
            .Select(variant => $"/{kind}[{variant}]")];

    /// <summary>Resolves /header[firstPage]-style roots; canonical paths stay numeric (part order).</summary>
    private static (OpenXmlElement El, string Path, string Type) ResolveVariantRoot(
        WordprocessingDocument doc, string kind, string variant)
    {
        var part = FindReferencedPart(doc, kind, variant);
        if (part is null)
        {
            var existing = ExistingVariantPaths(doc, kind);
            throw Invalid(
                $"/{kind}[{variant}] does not exist. Create it with " +
                $"{{\"op\":\"add\",\"path\":\"/{kind}[{variant}]\",\"type\":\"{kind}\",\"props\":{{\"text\":\"…\"}}}}.",
                existing.Count > 0 ? existing : ["/body"]);
        }

        OpenXmlElement? element;
        int index;
        if (part is HeaderPart headerPart)
        {
            element = headerPart.Header;
            index = (doc.MainDocumentPart?.HeaderParts.ToList() ?? []).IndexOf(headerPart) + 1;
        }
        else
        {
            var footerPart = (FooterPart)part;
            element = footerPart.Footer;
            index = (doc.MainDocumentPart?.FooterParts.ToList() ?? []).IndexOf(footerPart) + 1;
        }

        return element is null
            ? throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                $"/{kind}[{variant}] exists but its part is empty.",
                "Re-export the file from Word, or address /body instead.")
            : ((OpenXmlElement El, string Path, string Type))(element, Canon(string.Empty, kind, index), kind);
    }

    private static (OpenXmlElement El, string Path, string Type) ResolveChild(
        OpenXmlElement parent, string parentPath, string parentType, PathSegment segment)
    {
        if (segment.Kind != PathSegmentKind.Element || segment.Name is null)
        {
            throw Invalid(
                $"'{segment.ToCanonicalString()}' is not a docx segment under {parentPath}.",
                ChildCandidates(parent, parentPath, parentType));
        }

        var name = segment.Name;
        var allowed = ChildNames[parentType];
        if (!allowed.Contains(name, StringComparer.Ordinal))
        {
            throw Invalid(
                allowed.Length == 0
                    ? $"{parentPath} is a leaf ({parentType}); it has no addressable children."
                    : $"'{name}' cannot appear under {parentPath} ({parentType} contains: {string.Join(", ", allowed)}).",
                ChildCandidates(parent, parentPath, parentType));
        }

        if (segment.Index is not { } index)
        {
            throw Invalid(
                $"'{name}' needs a 1-based index under {parentPath}, e.g. {name}[1].",
                ChildCandidates(parent, parentPath, parentType));
        }

        var children = ChildrenOf(parent, name);
        if (index > children.Count)
        {
            throw Invalid(
                $"{parentPath}/{name}[{index}] does not exist; there are {children.Count} '{name}' element(s).",
                children.Count == 0
                    ? ChildCandidates(parent, parentPath, parentType)
                    : NearestIndexCandidates(parentPath, name, index, children.Count));
        }

        return (children[index - 1], Canon(parentPath, name, index), name);
    }

    /// <summary>Direct children of one addressable type, in document order.</summary>
    private static List<OpenXmlElement> ChildrenOf(OpenXmlElement parent, string name) => name switch
    {
        "p" => [.. parent.ChildElements.OfType<Paragraph>()],
        "table" => [.. parent.ChildElements.OfType<Table>()],
        "run" => [.. parent.ChildElements.OfType<Run>()],
        "link" => [.. parent.ChildElements.OfType<Hyperlink>()],
        "tr" => [.. parent.ChildElements.OfType<TableRow>()],
        "tc" => [.. parent.ChildElements.OfType<TableCell>()],
        _ => [],
    };

    private static string Canon(string parentPath, string name, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"{parentPath}/{name}[{index}]");

    private static IReadOnlyList<string> RootCandidates(WordprocessingDocument doc)
    {
        var candidates = new List<string> { "/body" };
        if (doc.MainDocumentPart?.HeaderParts.Any() == true)
        {
            candidates.Add("/header[1]");
        }

        if (doc.MainDocumentPart?.FooterParts.Any() == true)
        {
            candidates.Add("/footer[1]");
        }

        return candidates;
    }

    /// <summary>First existing path of every child type the parent can hold.</summary>
    private static IReadOnlyList<string> ChildCandidates(OpenXmlElement parent, string parentPath, string parentType)
    {
        var candidates = new List<string>();
        foreach (var name in ChildNames[parentType])
        {
            var count = ChildrenOf(parent, name).Count;
            if (count > 0)
            {
                candidates.Add(Canon(parentPath, name, 1));
            }
        }

        return candidates.Count > 0 ? candidates : [parentPath];
    }

    private static IReadOnlyList<string> NearestIndexCandidates(string parentPath, string name, int requested, int count)
    {
        return [.. Enumerable.Range(1, count)
            .OrderBy(n => Math.Abs(n - requested))
            .Take(5)
            .OrderBy(n => n)
            .Select(n => Canon(parentPath, name, n))];
    }

    private static AiofficeException Invalid(string message, IReadOnlyList<string> candidates) =>
        new(
            ErrorCodes.InvalidPath,
            message,
            "Use a candidate path, or run 'aioffice query <file> \"*\"' to list addressable nodes.",
            candidates);

    // -------------------------------------------------------------- enumerate

    /// <summary>
    /// Yields every addressable node under the body with its canonical path,
    /// in document order: paragraphs, runs, tables, rows, cells, cell content.
    /// </summary>
    public static IEnumerable<ResolvedNode> EnumerateBody(Body body)
    {
        foreach (var node in EnumerateContainer(body, "/body"))
        {
            yield return node;
        }
    }

    /// <summary>
    /// The header and footer roots of a document, with canonical paths
    /// (/header[i], /footer[i]) in part order — the same order Resolve uses.
    /// </summary>
    public static IEnumerable<ResolvedNode> HeaderFooterRoots(WordprocessingDocument doc)
    {
        var headerIndex = 0;
        foreach (var part in doc.MainDocumentPart?.HeaderParts ?? [])
        {
            headerIndex++;
            if (part.Header is { } header)
            {
                yield return new ResolvedNode(header, Canon(string.Empty, "header", headerIndex), "header");
            }
        }

        var footerIndex = 0;
        foreach (var part in doc.MainDocumentPart?.FooterParts ?? [])
        {
            footerIndex++;
            if (part.Footer is { } footer)
            {
                yield return new ResolvedNode(footer, Canon(string.Empty, "footer", footerIndex), "footer");
            }
        }
    }

    /// <summary>
    /// Every addressable node in the document: body content first, then the
    /// content of each header and footer under its /header[i] | /footer[i] path.
    /// </summary>
    public static IEnumerable<ResolvedNode> EnumerateAll(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document?.Body is { } body)
        {
            foreach (var node in EnumerateBody(body))
            {
                yield return node;
            }
        }

        foreach (var root in HeaderFooterRoots(doc))
        {
            foreach (var node in EnumerateContainer(root.Element, root.CanonicalPath))
            {
                yield return node;
            }
        }
    }

    private static IEnumerable<ResolvedNode> EnumerateContainer(OpenXmlElement container, string path)
    {
        var pIndex = 0;
        var tableIndex = 0;
        foreach (var child in container.ChildElements)
        {
            switch (child)
            {
                case Paragraph paragraph:
                {
                    pIndex++;
                    var pPath = Canon(path, "p", pIndex);
                    yield return new ResolvedNode(paragraph, pPath, "p");
                    var runIndex = 0;
                    foreach (var run in paragraph.ChildElements.OfType<Run>())
                    {
                        runIndex++;
                        yield return new ResolvedNode(run, Canon(pPath, "run", runIndex), "run");
                    }

                    var linkIndex = 0;
                    foreach (var link in paragraph.ChildElements.OfType<Hyperlink>())
                    {
                        linkIndex++;
                        yield return new ResolvedNode(link, Canon(pPath, "link", linkIndex), "link");
                    }

                    break;
                }

                case Table table:
                {
                    tableIndex++;
                    var tablePath = Canon(path, "table", tableIndex);
                    yield return new ResolvedNode(table, tablePath, "table");
                    var rowIndex = 0;
                    foreach (var row in table.ChildElements.OfType<TableRow>())
                    {
                        rowIndex++;
                        var rowPath = Canon(tablePath, "tr", rowIndex);
                        yield return new ResolvedNode(row, rowPath, "tr");
                        var cellIndex = 0;
                        foreach (var cell in row.ChildElements.OfType<TableCell>())
                        {
                            cellIndex++;
                            var cellPath = Canon(rowPath, "tc", cellIndex);
                            yield return new ResolvedNode(cell, cellPath, "tc");
                            foreach (var nested in EnumerateContainer(cell, cellPath))
                            {
                                yield return nested;
                            }
                        }
                    }

                    break;
                }

                default:
                    break;
            }
        }
    }
}
