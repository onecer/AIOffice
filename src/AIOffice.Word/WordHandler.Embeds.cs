using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using WObject = DocumentFormat.OpenXml.Wordprocessing.EmbeddedObject;

namespace AIOffice.Word;

/// <summary>
/// The M10 embedded-objects surface for docx: a file of any type is embedded as
/// an OLE package object (a <c>w:object</c> wrapping an <c>o:OLEObject</c> that
/// references an <see cref="EmbeddedPackagePart"/>), listed from the parts, and
/// extracted back out byte-for-byte. Embeds are addressed <c>/embed[i]</c> in a
/// deterministic order — main-document embeds first, then those referenced from
/// each header and footer part (so header/footer embeds are never missed).
/// </summary>
public sealed partial class WordHandler : IEmbedHost
{
    // A square-ish OLE icon placeholder (1x1 PNG); the payload bytes are what
    // matter, the icon is only what Word paints for the package object.
    private const string EmbedIconPngBase64 =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

    // ----------------------------------------------------------- IEmbedHost

    /// <summary>Lists every embedded object's metadata in canonical <c>/embed[i]</c> order.</summary>
    public IReadOnlyList<Core.EmbeddedObject> ListEmbeds(CommandContext ctx)
    {
        var file = RequireFile(ctx, mustExist: true);
        var (doc, ms, _) = OpenCopy(file, editable: false);
        using (doc)
        using (ms)
        {
            return EnumerateEmbeds(doc).Select(e => e.ToMetadata()).ToList();
        }
    }

    /// <summary>Writes the embedded payload addressed by <paramref name="embedPath"/> to the sandbox-resolved dest. Never mutates the source.</summary>
    public void ExtractEmbed(CommandContext ctx, string embedPath, string destPath)
    {
        var file = RequireFile(ctx, mustExist: true);

        // SECURITY: the dest is re-resolved through the sandbox defensively, even
        // though the command layer also resolves it.
        var resolvedDest = ctx.Workspace.Resolve(destPath);

        var (doc, ms, _) = OpenCopy(file, editable: false);
        using (doc)
        using (ms)
        {
            var embed = ResolveEmbed(doc, DocPath.Parse(embedPath));
            WriteEmbedPayload(embed, resolvedDest);
        }
    }

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body","type":"embed","props":{"src":…,"name":…,"icon":…}}</c>:
    /// embeds the sandbox-resolved src as an OLE package object in a new paragraph.
    /// The media type is sniffed from the file; returns the canonical embed path.
    /// </summary>
    private static object ApplyAddEmbed(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        _ = session.ResolveAuthor(props); // consume props.author; embeds are not tracked-change carriers

        var src = props["src"] is { } srcNode ? NodeToString(srcNode) : null;
        if (string.IsNullOrWhiteSpace(src))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type embed needs props.src (a workspace-relative file path).",
                "Example: {\"op\":\"add\",\"path\":\"/body\",\"type\":\"embed\",\"props\":{\"src\":\"data.xlsx\",\"name\":\"Q3 model\"}}.");
        }

        // SECURITY: the only road to the payload bytes is through the sandbox.
        var resolved = session.Workspace.Resolve(src, mustExist: true);
        if (!File.Exists(resolved))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"Embed source not found: {src}",
                "Check the path; it is resolved relative to the workspace root.");
        }

        if (session.Track)
        {
            throw TrackedStructureUnsupported("embed");
        }

        var bytes = File.ReadAllBytes(resolved);
        var displayName = props["name"] is { } nameNode && NodeToString(nameNode) is { Length: > 0 } n
            ? n
            : Path.GetFileName(src);

        // Sniff the media type from the SOURCE file name (the authoritative one),
        // not the display name — a display name like "Q3 model" carries no extension
        // and would otherwise fall back to application/zip for an .xlsx. This keeps
        // the embed media type consistent with the xlsx/pptx handlers, which sniff
        // the src file directly.
        var mediaType = SniffMediaType(Path.GetFileName(src), bytes);

        // Anchor: containers (body/tc/header/footer) take the new paragraph inside;
        // block anchors (p/table) take it before/after, like images.
        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        var position = op.Position ?? (anchor.Type is "body" or "tc" or "header" or "footer" ? "inside" : "after");
        var valid = (anchor.Type, position) switch
        {
            ("body" or "tc" or "header" or "footer", "inside") => true,
            ("p" or "table", "before" or "after") => true,
            _ => false,
        };

        if (!valid)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cannot add an embed {position} {anchor.CanonicalPath} ({anchor.Type}).",
                "Add embeds inside /body (or a tc/header/footer), or before/after an existing paragraph.");
        }

        // The package object's parts hang off the part that owns the anchor, so an
        // embed dropped into a header lands in the header part (and is found there).
        var hostPart = OwningPart(doc, anchor.Element) ?? doc.MainDocumentPart!;

        var pkg = hostPart.AddNewPart<EmbeddedPackagePart>(mediaType);
        using (var stream = new MemoryStream(bytes, writable: false))
        {
            pkg.FeedData(stream);
        }

        var pkgRelId = hostPart.GetIdOfPart(pkg);

        var iconBytes = ResolveIconBytes(session, props);
        var iconPart = hostPart.AddNewPart<ImagePart>(IconContentType(iconBytes), null);
        using (var stream = new MemoryStream(iconBytes, writable: false))
        {
            iconPart.FeedData(stream);
        }

        var iconRelId = hostPart.GetIdOfPart(iconPart);

        var shapeId = NextEmbedShapeId(doc);
        var objectId = NextEmbedObjectId(doc);

        var embeddedObject = new WObject(
            new V.Shape(
                new V.ImageData { RelationshipId = iconRelId, Title = displayName })
            {
                Id = shapeId,
                Style = "width:48pt;height:48pt",
            },
            new Ovml.OleObject
            {
                Type = Ovml.OleValues.Embed,
                ProgId = "Package",
                ShapeId = shapeId,
                DrawAspect = Ovml.OleDrawAspectValues.Icon,
                Id = pkgRelId,
                ObjectId = objectId,
            })
        {
            DxaOriginal = "1440",
            DyaOriginal = "1440",
        };

        var paragraph = new Paragraph(new Run(embeddedObject));
        var landed = Insert(doc, anchor, paragraph, position);

        // Find the canonical /embed[i] path the new object occupies.
        var embedPath = EnumerateEmbeds(doc)
            .FirstOrDefault(e => ReferenceEquals(e.Part, pkg))?.CanonicalPath ?? "/embed[1]";

        return new
        {
            op = "add",
            type = "embed",
            path = embedPath,
            anchor = landed,
            name = displayName,
            mediaType,
            size = (long)bytes.Length,
        };
    }

    // --------------------------------------------------------------- extract

    /// <summary>
    /// <c>{"op":"extract","path":"/embed[1]","props":{"to":…}}</c>: writes the
    /// embedded payload to the sandbox-resolved dest. Mutating-but-producing — it
    /// reads the live document state but does NOT change it.
    /// </summary>
    private static object ApplyExtractEmbed(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var to = op.Props?["to"] is { } toNode ? NodeToString(toNode) : null;
        if (string.IsNullOrWhiteSpace(to))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "extract needs props.to (a workspace-relative destination path).",
                "Example: {\"op\":\"extract\",\"path\":\"/embed[1]\",\"props\":{\"to\":\"out/data.xlsx\"}}.");
        }

        // SECURITY: the extract dest must resolve through the sandbox (no escaping
        // dest), exactly like an embed src must resolve to read its bytes.
        var resolvedDest = session.Workspace.Resolve(to);

        var embed = ResolveEmbed(doc, DocPath.Parse(op.Path));
        var size = WriteEmbedPayload(embed, resolvedDest);

        return new
        {
            op = "extract",
            path = embed.CanonicalPath,
            to,
            name = embed.DisplayName,
            mediaType = embed.MediaType,
            size,
        };
    }

    /// <summary>Streams an embed's payload to a sandbox-resolved dest and returns the byte count.</summary>
    private static long WriteEmbedPayload(EmbedEntry embed, string resolvedDest)
    {
        var dir = Path.GetDirectoryName(resolvedDest);
        if (dir is { Length: > 0 })
        {
            Directory.CreateDirectory(dir);
        }

        using var source = embed.Part.GetStream(FileMode.Open, FileAccess.Read);
        using var dest = new FileStream(resolvedDest, FileMode.Create, FileAccess.Write);
        source.CopyTo(dest);
        return dest.Length;
    }

    // ---------------------------------------------------------------- remove

    /// <summary>remove /embed[i]: deletes the w:object run and the backing package/icon parts.</summary>
    private static object ApplyRemoveEmbed(WordprocessingDocument doc, EditOp op)
    {
        var embed = ResolveEmbed(doc, DocPath.Parse(op.Path));

        // Drop the w:object's containing run; if that empties the paragraph and it
        // is removable, leave it (it is harmless), matching image-removal restraint.
        if (embed.ObjectElement?.Ancestors<Run>().FirstOrDefault() is { } run)
        {
            run.Remove();
        }
        else
        {
            embed.ObjectElement?.Remove();
        }

        // Delete the icon part referenced by the OLE shape, then the package part.
        if (embed.IconPart is { } icon)
        {
            embed.HostPart.DeletePart(icon);
        }

        embed.HostPart.DeletePart(embed.Part);

        return new { op = "remove", path = embed.CanonicalPath, type = "embed" };
    }

    // ------------------------------------------------------------------- get

    /// <summary>get on /embed[i]: the canonical path plus metadata (name, mediaType, size) — NOT the payload bytes.</summary>
    private static (string Path, Dictionary<string, object?> Properties) GetEmbedProperties(WordprocessingDocument doc, DocPath path)
    {
        var embed = ResolveEmbed(doc, path);
        var properties = new Dictionary<string, object?>
        {
            ["name"] = embed.DisplayName,
            ["mediaType"] = embed.MediaType,
            ["size"] = embed.Size,
            ["container"] = embed.Container,
            ["note"] = "Use the extract op to write the payload bytes out: {\"op\":\"extract\",\"path\":\"" +
                embed.CanonicalPath + "\",\"props\":{\"to\":\"out" + GuessExtension(embed) + "\"}}.",
        };
        return (embed.CanonicalPath, properties);
    }

    // ------------------------------------------------------------------ read

    /// <summary>read --view embeds: the EmbeddedObject metadata list.</summary>
    private static object EmbedsView(WordprocessingDocument doc) => new
    {
        view = "embeds",
        embeds = EnumerateEmbeds(doc)
            .Select(e => new
            {
                path = e.CanonicalPath,
                name = e.DisplayName,
                mediaType = e.MediaType,
                size = e.Size,
                container = e.Container,
            })
            .ToList(),
    };

    // ------------------------------------------------------------ enumerate

    /// <summary>One embedded object: its package part, the icon part, the w:object element and addressing facts.</summary>
    private sealed record EmbedEntry(
        OpenXmlPart Part,
        OpenXmlPartContainer HostPart,
        ImagePart? IconPart,
        WObject? ObjectElement,
        string DisplayName,
        string MediaType,
        long Size,
        string? Container,
        string CanonicalPath)
    {
        public Core.EmbeddedObject ToMetadata() =>
            new(CanonicalPath, DisplayName, MediaType, Size, Container);
    }

    /// <summary>
    /// Every embedded package/object in the document, in deterministic order:
    /// the main document part first, then each header part, then each footer part
    /// (so header/footer embeds are discovered). Within an owner, embeds are
    /// ordered by the document position of their referencing <c>o:OLEObject</c>
    /// (so the list matches reading order), with any unreferenced parts appended
    /// by part URI. Numbered /embed[i] 1-based across the whole document.
    /// </summary>
    private static IReadOnlyList<EmbedEntry> EnumerateEmbeds(WordprocessingDocument doc)
    {
        var result = new List<EmbedEntry>();
        var index = 0;

        foreach (var host in EmbedHostParts(doc))
        {
            // EmbeddedPackagePart (a generic package object) and EmbeddedObjectPart
            // (a legacy OLE object) both carry an embedded payload; list both.
            var parts = host.Part.GetPartsOfType<EmbeddedPackagePart>()
                .Cast<OpenXmlPart>()
                .Concat(host.Part.GetPartsOfType<EmbeddedObjectPart>())
                .ToList();

            // The document position of each relId referenced by an OLE object, so
            // embeds list in reading order rather than part-URI order.
            var order = OleReferenceOrder(host.Part);

            var ordered = parts
                .Select(p => (Part: p, RelId: TryGetRelId(host.Part, p)))
                .OrderBy(t => t.RelId is { } id && order.TryGetValue(id, out var pos) ? pos : int.MaxValue)
                .ThenBy(t => t.Part.Uri.ToString(), StringComparer.Ordinal)
                .ToList();

            foreach (var (part, relId) in ordered)
            {
                index++;
                var (objectElement, iconPart, displayName) = DescribeObject(host, relId, part);

                result.Add(new EmbedEntry(
                    Part: part,
                    HostPart: host.Part,
                    IconPart: iconPart,
                    ObjectElement: objectElement,
                    DisplayName: displayName,
                    MediaType: part.ContentType,
                    Size: PartSize(part),
                    Container: host.ContainerPath,
                    CanonicalPath: EmbedPath(index)));
            }
        }

        return result;
    }

    /// <summary>Maps each OLE-referenced relId to the 0-based document position of its w:object, for reading-order listing.</summary>
    private static Dictionary<string, int> OleReferenceOrder(OpenXmlPartContainer container)
    {
        var order = new Dictionary<string, int>(StringComparer.Ordinal);
        if (container is not OpenXmlPart part || RootElementOf(part) is not { } root)
        {
            return order;
        }

        var position = 0;
        foreach (var ole in root.Descendants<Ovml.OleObject>())
        {
            if (ole.Id?.Value is { Length: > 0 } id && !order.ContainsKey(id))
            {
                order[id] = position++;
            }
        }

        return order;
    }

    /// <summary>The parts that can own an embedded object, with the container path embeds under them report.</summary>
    private static IEnumerable<(OpenXmlPartContainer Part, string? ContainerPath)> EmbedHostParts(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart is { } main)
        {
            yield return (main, null);

            var headerIndex = 0;
            foreach (var header in main.HeaderParts)
            {
                headerIndex++;
                yield return (header, $"/header[{headerIndex}]");
            }

            var footerIndex = 0;
            foreach (var footer in main.FooterParts)
            {
                footerIndex++;
                yield return (footer, $"/footer[{footerIndex}]");
            }
        }
    }

    /// <summary>Finds the w:object element referencing a package part, plus its icon part and display name.</summary>
    private static (WObject? Object, ImagePart? Icon, string Name) DescribeObject(
        (OpenXmlPartContainer Part, string? ContainerPath) host, string? relId, OpenXmlPart part)
    {
        if (relId is null || host.Part is not OpenXmlPart ownerPart)
        {
            return (null, null, PartFileName(part));
        }

        var root = RootElementOf(ownerPart);
        var obj = root?
            .Descendants<WObject>()
            .FirstOrDefault(o => o.Descendants<Ovml.OleObject>().Any(ole => ole.Id?.Value == relId));

        var imageData = obj?.Descendants<V.ImageData>().FirstOrDefault();
        var name = imageData?.Title?.Value;

        ImagePart? icon = null;
        var iconRelId = imageData?.RelationshipId?.Value;
        if (iconRelId is { Length: > 0 })
        {
            try
            {
                icon = ownerPart.GetPartById(iconRelId) as ImagePart;
            }
            catch (ArgumentOutOfRangeException)
            {
                // Dangling icon reference: leave icon null.
            }
        }

        return (obj, icon, string.IsNullOrEmpty(name) ? PartFileName(part) : name);
    }

    /// <summary>Resolves /embed[i] (1-based positional) to a live embed, or throws invalid_path with candidates.</summary>
    private static EmbedEntry ResolveEmbed(WordprocessingDocument doc, DocPath path)
    {
        var embeds = EnumerateEmbeds(doc);
        var segment = path.Segments.Count == 1 ? path.Segments[0] : null;

        if (segment is null || segment.Name != "embed")
        {
            throw EmbedNotFound($"'{path.ToCanonicalString()}' is not an embed path.", embeds);
        }

        var index = segment.Index ?? 1;
        if (index <= embeds.Count)
        {
            return embeds[index - 1];
        }

        throw EmbedNotFound(
            embeds.Count == 0
                ? "This document has no embedded objects."
                : $"/embed[{index}] does not exist; there are {embeds.Count} embedded object(s).",
            embeds);
    }

    private static AiofficeException EmbedNotFound(string message, IReadOnlyList<EmbedEntry> embeds) => new(
        ErrorCodes.InvalidPath,
        message,
        "Run 'aioffice read <file> --view embeds' to list embedded objects with their /embed[i] paths.",
        candidates: [.. embeds.Take(5).Select(e => e.CanonicalPath)]);

    // -------------------------------------------------------------- helpers

    private static string EmbedPath(int index) =>
        string.Create(CultureInfo.InvariantCulture, $"/embed[{index}]");

    /// <summary>The relationship id of a part inside its container, or null when unreferenced.</summary>
    private static string? TryGetRelId(OpenXmlPartContainer container, OpenXmlPart part)
    {
        foreach (var pair in container.Parts)
        {
            if (ReferenceEquals(pair.OpenXmlPart, part))
            {
                return pair.RelationshipId;
            }
        }

        return null;
    }

    private static OpenXmlElement? RootElementOf(OpenXmlPart part) => part switch
    {
        MainDocumentPart main => main.Document,
        HeaderPart header => header.Header,
        FooterPart footer => footer.Footer,
        _ => null,
    };

    private static long PartSize(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return stream.Length;
    }

    private static string PartFileName(OpenXmlPart part) => Path.GetFileName(part.Uri.ToString());

    /// <summary>A guessed dest extension from the media type, for the get-op hint.</summary>
    private static string GuessExtension(EmbedEntry embed) => embed.MediaType switch
    {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" => ".xlsx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document" => ".docx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation" => ".pptx",
        "application/pdf" => ".pdf",
        "application/zip" => ".zip",
        _ => Path.GetExtension(embed.DisplayName) is { Length: > 0 } ext ? ext : ".bin",
    };

    /// <summary>Sniffs a media type from the magic bytes first, then the name's extension, then a generic fallback.</summary>
    private static string SniffMediaType(string name, byte[] bytes)
    {
        // OOXML/zip-family files all start with the "PK" local-file header;
        // disambiguate the Office members by extension, fall back to zip.
        var isZip = bytes.Length >= 4 && bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04;
        var isPdf = bytes.Length >= 5 && bytes[0] == 0x25 && bytes[1] == 0x50 && bytes[2] == 0x44 && bytes[3] == 0x46 && bytes[4] == 0x2D;
        var ext = Path.GetExtension(name).ToLowerInvariant();

        if (isPdf)
        {
            return "application/pdf";
        }

        return ext switch
        {
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".pdf" => "application/pdf",
            ".zip" => "application/zip",
            ".csv" => "text/csv",
            ".txt" => "text/plain",
            ".json" => "application/json",
            ".xml" => "application/xml",
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            _ when isZip => "application/zip",
            _ => "application/octet-stream",
        };
    }

    /// <summary>The OLE icon bytes: a caller-supplied PNG/JPEG (props.icon, sandbox-resolved) or the built-in placeholder.</summary>
    private static byte[] ResolveIconBytes(EditSession session, System.Text.Json.Nodes.JsonObject props)
    {
        if (props["icon"] is { } iconNode && NodeToString(iconNode) is { Length: > 0 } icon)
        {
            // SECURITY: a caller-supplied icon path is a file-valued prop — resolve it.
            var resolved = session.Workspace.Resolve(icon, mustExist: true);
            var bytes = File.ReadAllBytes(resolved);
            _ = SniffImage(bytes, icon); // PNG/JPEG only; throws unsupported_feature otherwise
            return bytes;
        }

        return Convert.FromBase64String(EmbedIconPngBase64);
    }

    private static string IconContentType(byte[] iconBytes) =>
        iconBytes.Length >= 3 && iconBytes[0] == 0xFF && iconBytes[1] == 0xD8 && iconBytes[2] == 0xFF
            ? "image/jpeg"
            : "image/png";

    /// <summary>The next free OLE shape id (Word uses the "_x0000_iNNNN" convention).</summary>
    private static string NextEmbedShapeId(WordprocessingDocument doc)
    {
        var max = 1025;
        foreach (var host in EmbedHostParts(doc))
        {
            if (RootElementOf((OpenXmlPart)host.Part) is { } root)
            {
                foreach (var shape in root.Descendants<V.Shape>())
                {
                    var id = shape.Id?.Value;
                    if (id is { } s && s.StartsWith("_x0000_i", StringComparison.Ordinal) &&
                        int.TryParse(s.AsSpan("_x0000_i".Length), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) &&
                        n > max)
                    {
                        max = n;
                    }
                }
            }
        }

        return string.Create(CultureInfo.InvariantCulture, $"_x0000_i{max + 1}");
    }

    /// <summary>The next free OLE object id ("_NNNNNNNNNN"), unique across the document.</summary>
    private static string NextEmbedObjectId(WordprocessingDocument doc)
    {
        var max = 1_000_000_000L;
        foreach (var host in EmbedHostParts(doc))
        {
            if (RootElementOf((OpenXmlPart)host.Part) is { } root)
            {
                foreach (var ole in root.Descendants<Ovml.OleObject>())
                {
                    var id = ole.ObjectId?.Value;
                    if (id is { } s && s.StartsWith('_') &&
                        long.TryParse(s.AsSpan(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) &&
                        n > max)
                    {
                        max = n;
                    }
                }
            }
        }

        return string.Create(CultureInfo.InvariantCulture, $"_{max + 1}");
    }
}
