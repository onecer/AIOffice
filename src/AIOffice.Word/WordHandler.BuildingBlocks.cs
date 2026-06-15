using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// Building blocks / Quick Parts (v1.5.0). Reusable content stored in the
/// document's glossary part (<c>w:docPart</c> entries): a named block in a gallery
/// (quickParts | autoText) holding one or more paragraphs of content. Add stores a
/// block; <c>buildingBlockRef</c> inserts a stored block's content into the body by
/// name; <c>get /buildingBlock[@name=…]</c> reports it; <c>remove</c> drops it;
/// <c>read --view structure</c> lists every stored block. Validator-clean.
/// </summary>
public sealed partial class WordHandler
{
    /// <summary>The galleries a building block can live in, mapped to their OOXML gallery values.</summary>
    private static readonly Dictionary<string, DocPartGalleryValues> BuildingBlockGalleries =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["quickParts"] = DocPartGalleryValues.CustomQuickParts,
            ["autoText"] = DocPartGalleryValues.AutoText,
        };

    private const string BuildingBlocksPath = "/buildingBlocks";

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/buildingBlocks","type":"buildingBlock",
    /// "props":{"name":"Disclaimer","gallery":"quickParts|autoText","category"?,"content":"…"}}</c>:
    /// stores a reusable block in the glossary part. <c>content</c> is either plain
    /// text (one paragraph, "\n" splitting further paragraphs) or a JSON array of
    /// paragraph strings. Re-adding a block of the same name replaces it in place.
    /// </summary>
    private static object ApplyAddBuildingBlock(WordprocessingDocument doc, EditOp op)
    {
        if (!IsBuildingBlocksPath(op.Path))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add --type buildingBlock targets /buildingBlocks, not '{op.Path}'.",
                "Use {\"op\":\"add\",\"path\":\"/buildingBlocks\",\"type\":\"buildingBlock\"," +
                "\"props\":{\"name\":\"Disclaimer\",\"gallery\":\"quickParts\",\"content\":\"…\"}}.",
                candidates: [BuildingBlocksPath]);
        }

        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author");

        var name = props["name"] is { } nameNode ? NodeToString(nameNode).Trim() : null;
        if (string.IsNullOrEmpty(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type buildingBlock needs props.name (the block's name).",
                "Pass a unique name, e.g. {\"props\":{\"name\":\"Disclaimer\",\"gallery\":\"quickParts\",\"content\":\"…\"}}.");
        }

        var galleryArg = props["gallery"] is { } galleryNode ? NodeToString(galleryNode) : "quickParts";
        if (!BuildingBlockGalleries.TryGetValue(galleryArg, out var gallery))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown building-block gallery '{galleryArg}'.",
                $"Use gallery {string.Join(" or ", BuildingBlockGalleries.Keys)}.",
                candidates: [.. BuildingBlockGalleries.Keys]);
        }

        var category = props["category"] is { } categoryNode && NodeToString(categoryNode) is { Length: > 0 } cat
            ? cat
            : "General";

        var paragraphs = BuildingBlockContentParagraphs(props["content"]);

        var glossary = EnsureGlossaryDocPart(doc);
        var docParts = glossary.DocParts ??= new DocParts();
        var existing = FindBuildingBlock(doc, name);
        var refreshed = existing is not null;

        var docPart = BuildDocPart(name, gallery, category, paragraphs);
        if (existing is not null)
        {
            existing.InsertAfterSelf(docPart);
            existing.Remove();
        }
        else
        {
            docParts.AppendChild(docPart);
        }

        return new
        {
            op = "add",
            type = "buildingBlock",
            path = BuildingBlockPath(name),
            name,
            gallery = galleryArg,
            category,
            paragraphs = paragraphs.Count,
            refreshed = refreshed ? true : (bool?)null,
        };
    }

    /// <summary>
    /// <c>{"op":"add","path":"/body","type":"buildingBlockRef","props":{"name":"Disclaimer","position"?}}</c>:
    /// inserts a stored block's content into the body. Default placement is the end
    /// of the body (before the section properties); <c>position:"inside"</c> is the
    /// same, and a paragraph anchor with before/after places it relative to that
    /// paragraph (handled by the generic add validation downstream is bypassed —
    /// this resolves its own anchor).
    /// </summary>
    private static object ApplyAddBuildingBlockRef(WordprocessingDocument doc, EditOp op)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author");

        var name = props["name"] is { } nameNode ? NodeToString(nameNode).Trim() : null;
        if (string.IsNullOrEmpty(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type buildingBlockRef needs props.name (the stored block to insert).",
                "Example: {\"op\":\"add\",\"path\":\"/body\",\"type\":\"buildingBlockRef\",\"props\":{\"name\":\"Disclaimer\"}}.");
        }

        var docPart = FindBuildingBlock(doc, name) ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"No building block named '{name}' is stored in this document.",
            "Store it first with add --type buildingBlock, or list the stored blocks with read --view structure.",
            candidates: [.. EnumerateBuildingBlocks(doc).Select(b => BuildingBlockPath(b.Name))]);

        // The block's content paragraphs, cloned into the main document so the
        // glossary part stays the canonical source.
        var content = (docPart.DocPartBody?.Elements<Paragraph>() ?? [])
            .Select(p => (Paragraph)p.CloneNode(true))
            .ToList();
        if (content.Count == 0)
        {
            content.Add(new Paragraph());
        }

        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        var position = op.Position ?? "inside";
        var firstPath = InsertBuildingBlockContent(doc, anchor, content, position, op.Path);

        return new
        {
            op = "add",
            type = "buildingBlockRef",
            name,
            anchor = anchor.CanonicalPath,
            path = firstPath,
            paragraphs = content.Count,
        };
    }

    /// <summary>Places the cloned block paragraphs at the anchor and returns the first new paragraph's canonical path.</summary>
    private static string InsertBuildingBlockContent(
        WordprocessingDocument doc, ResolvedNode anchor, List<Paragraph> content, string position, string rawPath)
    {
        if (position is not (null or "before" or "after" or "inside"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"buildingBlockRef position '{position}' is not valid.",
                "Use position inside (append to a /body, /header[n], /footer[n] or tc), or before/after a paragraph.",
                candidates: ["before", "after", "inside"]);
        }

        switch (anchor.Element)
        {
            case Body or Header or Footer or TableCell when position is "inside":
            {
                if (anchor.Element is Body body && body.Elements<SectionProperties>().FirstOrDefault() is { } sectPr)
                {
                    foreach (var p in content)
                    {
                        sectPr.InsertBeforeSelf(p);
                    }
                }
                else
                {
                    foreach (var p in content)
                    {
                        anchor.Element.AppendChild(p);
                    }
                }

                break;
            }

            case Paragraph when position is "before":
            {
                var cursor = anchor.Element;
                foreach (var p in content)
                {
                    cursor.InsertBeforeSelf(p);
                }

                break;
            }

            case Paragraph when position is "after":
            {
                OpenXmlElement cursor = anchor.Element;
                foreach (var p in content)
                {
                    cursor.InsertAfterSelf(p);
                    cursor = p;
                }

                break;
            }

            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Cannot insert a building block {position} {anchor.CanonicalPath} ({anchor.Type}).",
                    "Insert inside /body, /header[n], /footer[n] or a tc, or before/after an existing paragraph.");
        }

        var first = content[0];
        var match = WordAddress.EnumerateAll(doc).FirstOrDefault(n => ReferenceEquals(n.Element, first));
        return match?.CanonicalPath ?? anchor.CanonicalPath;
    }

    // ---------------------------------------------------------------- remove

    /// <summary><c>{"op":"remove","path":"/buildingBlock[@name=…]"}</c>: drops a stored block from the glossary part.</summary>
    private static object ApplyRemoveBuildingBlock(WordprocessingDocument doc, EditOp op)
    {
        var name = BuildingBlockNameOf(op.Path);
        var docPart = FindBuildingBlock(doc, name) ?? throw BuildingBlockNotFound(doc, name);
        docPart.Remove();
        return new { op = "remove", path = BuildingBlockPath(name), type = "buildingBlock", name };
    }

    // ------------------------------------------------------------------- get

    /// <summary>get /buildingBlock[@name=…] data: {name, gallery, category, paragraphs, content, contentLines}.</summary>
    private static (string Path, Dictionary<string, object?> Properties) GetBuildingBlockProperties(
        WordprocessingDocument doc, string pathArg)
    {
        var name = BuildingBlockNameOf(pathArg);
        var docPart = FindBuildingBlock(doc, name) ?? throw BuildingBlockNotFound(doc, name);
        return (BuildingBlockPath(name), BuildingBlockShape(docPart));
    }

    private static Dictionary<string, object?> BuildingBlockShape(DocPart docPart)
    {
        var lines = (docPart.DocPartBody?.Elements<Paragraph>() ?? [])
            .Select(p => p.InnerText)
            .ToList();
        return new Dictionary<string, object?>
        {
            ["name"] = docPart.DocPartProperties?.DocPartName?.Val?.Value,
            ["gallery"] = GalleryName(docPart.DocPartProperties?.Category?.Gallery?.Val),
            ["category"] = docPart.DocPartProperties?.Category?.Name?.Val?.Value,
            ["paragraphs"] = lines.Count,
            ["content"] = string.Join("\n", lines),
            ["contentLines"] = lines,
        };
    }

    // ------------------------------------------------------------- structure

    /// <summary>A stored building block resolved from the glossary part.</summary>
    internal sealed record BuildingBlockInfo(string Name, string Gallery, string Category, int Paragraphs);

    /// <summary>Every stored building block in glossary order; empty when there is no glossary part.</summary>
    private static List<BuildingBlockInfo> EnumerateBuildingBlocks(WordprocessingDocument doc)
    {
        var glossary = doc.MainDocumentPart?.GlossaryDocumentPart?.GlossaryDocument?.DocParts;
        if (glossary is null)
        {
            return [];
        }

        return [.. glossary.Elements<DocPart>().Select(dp => new BuildingBlockInfo(
            dp.DocPartProperties?.DocPartName?.Val?.Value ?? string.Empty,
            GalleryName(dp.DocPartProperties?.Category?.Gallery?.Val),
            dp.DocPartProperties?.Category?.Name?.Val?.Value ?? "General",
            (dp.DocPartBody?.Elements<Paragraph>() ?? []).Count()))];
    }

    /// <summary>Building-block shapes for read --view structure: {path, name, gallery, category, paragraphs}.</summary>
    private static List<object> BuildingBlocksStructure(WordprocessingDocument doc) =>
        [.. EnumerateBuildingBlocks(doc).Select(b => (object)new
        {
            path = BuildingBlockPath(b.Name),
            name = b.Name,
            gallery = b.Gallery,
            category = b.Category,
            paragraphs = b.Paragraphs,
        })];

    // ---------------------------------------------------------------- build

    /// <summary>Ensures the glossary document part exists and returns its <c>w:glossaryDocument</c> root.</summary>
    private static GlossaryDocument EnsureGlossaryDocPart(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The document has no main part to attach a glossary to.",
            "Re-export the file from Word.");

        var part = main.GlossaryDocumentPart ?? main.AddNewPart<GlossaryDocumentPart>();
        return part.GlossaryDocument ??= new GlossaryDocument(new DocParts());
    }

    /// <summary>A w:docPart with name/category/gallery/id and the content paragraphs in its body.</summary>
    private static DocPart BuildDocPart(string name, DocPartGalleryValues gallery, string category, List<Paragraph> paragraphs)
    {
        var body = new DocPartBody();
        foreach (var paragraph in paragraphs)
        {
            body.AppendChild(paragraph);
        }

        if (!body.Elements<Paragraph>().Any())
        {
            body.AppendChild(new Paragraph());
        }

        return new DocPart(
            new DocPartProperties(
                new DocPartName { Val = name },
                new Category(
                    new Name { Val = category },
                    new Gallery { Val = gallery }),
                new DocPartId { Val = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}" }),
            body);
    }

    /// <summary>
    /// Content paragraphs from the <c>content</c> prop. A JSON array gives one
    /// paragraph per string; a plain string splits on newlines (so a single block
    /// can carry several lines). Null/empty content yields one empty paragraph.
    /// </summary>
    private static List<Paragraph> BuildingBlockContentParagraphs(JsonNode? content)
    {
        if (content is JsonArray array)
        {
            return [.. array.Select(item => WordFactory.Paragraph(NodeToString(item)))];
        }

        var text = content is null ? string.Empty : NodeToString(content);
        if (text.Length == 0)
        {
            return [new Paragraph()];
        }

        return [.. text.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Split('\n')
            .Select(line => WordFactory.Paragraph(line))];
    }

    // ----------------------------------------------------------------- lookup

    private static DocPart? FindBuildingBlock(WordprocessingDocument doc, string name) =>
        doc.MainDocumentPart?.GlossaryDocumentPart?.GlossaryDocument?.DocParts?
            .Elements<DocPart>()
            .FirstOrDefault(dp => string.Equals(
                dp.DocPartProperties?.DocPartName?.Val?.Value, name, StringComparison.Ordinal));

    private static AiofficeException BuildingBlockNotFound(WordprocessingDocument doc, string name) => new(
        ErrorCodes.InvalidPath,
        $"No building block named '{name}' is stored in this document.",
        "List the stored blocks with read --view structure, or store one with add --type buildingBlock.",
        candidates: [.. EnumerateBuildingBlocks(doc).Select(b => BuildingBlockPath(b.Name))]);

    /// <summary>The gallery name ("quickParts" | "autoText" | the raw value) for a stored gallery enum.</summary>
    private static string GalleryName(EnumValue<DocPartGalleryValues>? gallery)
    {
        if (gallery?.Value is not { } value)
        {
            return "quickParts";
        }

        if (value == DocPartGalleryValues.AutoText)
        {
            return "autoText";
        }

        return value == DocPartGalleryValues.CustomQuickParts ? "quickParts" : value.ToString() ?? "quickParts";
    }

    // ------------------------------------------------------------------ paths

    internal static bool IsBuildingBlocksPath(string path) =>
        path == BuildingBlocksPath || path == "/buildingBlocks/";

    /// <summary>The block name from a /buildingBlock[@name=…] path, or invalid_path.</summary>
    private static string BuildingBlockNameOf(string path)
    {
        var docPath = DocPath.Parse(path);
        var segment = docPath.Segments.Count == 1 ? docPath.Segments[0] : null;
        if (segment is { Name: "buildingBlock", IdAttribute: "name", Id: { Length: > 0 } id })
        {
            return id;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"'{path}' is not a building-block path.",
            "Address a stored block by name: /buildingBlock[@name=Disclaimer].");
    }

    private static string BuildingBlockPath(string name) =>
        string.Create(CultureInfo.InvariantCulture, $"/buildingBlock[@name={name}]");
}
