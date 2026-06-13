using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Embedded objects on a slide: any file (a source .xlsx, a .pdf, a .zip, …)
/// stored as an OLE/package object inside the deck. The wire form is a
/// <c>p:graphicFrame</c> whose <c>a:graphicData</c> (the presentationml/ole uri)
/// carries a <c>p:oleObject</c> referencing an <see cref="EmbeddedPackagePart"/>
/// (the verbatim payload) plus a <c>p:picture</c> fallback PowerPoint paints for
/// the object — exactly the shape PowerPoint itself writes for "Insert Object →
/// from file". The host frame's shape id is the embed's stable id, so
/// <c>/slide[i]/embed[@id=N]</c> survives insertions and removals.
/// </summary>
internal static class PptxEmbeds
{
    /// <summary>The graphicData uri that marks a graphicFrame as a presentation OLE object.</summary>
    public const string OleNamespace = "http://schemas.openxmlformats.org/presentationml/2006/ole";

    /// <summary>The progId aioffice writes for a generic embedded package (PowerPoint "Package").</summary>
    private const string PackageProgId = "Package";

    private static readonly IReadOnlyList<string> AddPropKeys = ["src", "name", "icon", "x", "y", "w", "h"];

    /// <summary>One embedded object resolved on a slide: its host frame, the oleObject node and the payload part.</summary>
    private sealed record EmbedRef(int Ordinal, P.GraphicFrame Frame, P.OleObject Ole, EmbeddedPackagePart Part, uint HostId);

    // ---- add ----------------------------------------------------------------

    /// <summary>
    /// Embeds the sandbox-resolved <c>props.src</c> file as an OLE/package object on
    /// the slide and returns the canonical embed path (by the host frame's id).
    /// </summary>
    public static string Add(SlidePart slidePart, int slideIndex, JsonObject? props, Workspace workspace)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddPropKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown embed prop '{key}'.",
                    "Embed props: src (required), name, icon, x, y, w, h. " +
                    "src points at any file inside the workspace; name is the display name.",
                    candidates: AddPropKeys);
            }
        }

        if (!props.TryGetPropertyValue("src", out var srcNode) || srcNode is null || J.ScalarText(srcNode).Trim().Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add embed requires props.src.",
                "Point src at a file inside the workspace, e.g. " +
                "{\"op\":\"add\",\"path\":\"/slide[1]\",\"type\":\"embed\",\"props\":{\"src\":\"data.xlsx\"}}.");
        }

        // Sandbox first: a src outside the workspace is sandbox_denied, never read.
        var src = workspace.Resolve(J.ScalarText(srcNode).Trim(), mustExist: true);
        var bytes = File.ReadAllBytes(src);
        var mediaType = MediaType.Sniff(src);
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null && J.ScalarText(nameNode).Trim().Length > 0
            ? J.ScalarText(nameNode).Trim()
            : Path.GetFileName(src);

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);

        var x = props.TryGetPropertyValue("x", out var xNode) ? Units.ParseLengthEmu("x", xNode) : Units.CmToEmu(2.5);
        var y = props.TryGetPropertyValue("y", out var yNode) ? Units.ParseLengthEmu("y", yNode) : Units.CmToEmu(2.5);
        var w = props.TryGetPropertyValue("w", out var wNode) ? Units.ParseLengthEmu("w", wNode) : Units.CmToEmu(3);
        var h = props.TryGetPropertyValue("h", out var hNode) ? Units.ParseLengthEmu("h", hNode) : Units.CmToEmu(3);

        // The verbatim payload lives in an EmbeddedPackagePart; the bytes round-trip
        // exactly because FeedData copies them through with no re-encoding.
        var packagePart = slidePart.AddNewPart<EmbeddedPackagePart>(mediaType);
        using (var stream = new MemoryStream(bytes, writable: false))
        {
            packagePart.FeedData(stream);
        }

        var relId = slidePart.GetIdOfPart(packagePart);
        var fallback = BuildFallbackPicture(slidePart, id, name, props, workspace, x, y, w, h);

        var ole = new P.OleObject(new P.OleObjectEmbed(), fallback)
        {
            ShapeId = "0",
            Name = name,
            Id = relId,
            ImageWidth = (int)Math.Max(1, Units.EmuToPx(w)),
            ImageHeight = (int)Math.Max(1, Units.EmuToPx(h)),
            ProgId = PackageProgId,
        };

        tree.Append(new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = w, Cy = h }),
            new A.Graphic(new A.GraphicData(ole) { Uri = OleNamespace })));

        return Units.Inv($"/slide[{slideIndex}]/embed[@id={id}]");
    }

    /// <summary>
    /// The picture PowerPoint paints in place of the object: props.icon (a
    /// PNG/JPEG in the workspace) when given is embedded and shown; otherwise a
    /// neutral placeholder rectangle with an empty blip — both are valid
    /// p:oleObject fallbacks that validate and open in PowerPoint.
    /// </summary>
    private static P.Picture BuildFallbackPicture(
        SlidePart slidePart, uint id, string name, JsonObject props, Workspace workspace, long x, long y, long w, long h)
    {
        var transform = new A.Transform2D(
            new A.Offset { X = x, Y = y },
            new A.Extents { Cx = w, Cy = h });

        // props.icon, when given, becomes the visible preview: reuse the proven
        // PNG/JPEG image path so the blip references a real ImagePart.
        var blip = new A.Blip();
        if (props.TryGetPropertyValue("icon", out var iconNode) && iconNode is not null && J.ScalarText(iconNode).Trim().Length > 0)
        {
            var iconSrc = workspace.Resolve(J.ScalarText(iconNode).Trim(), mustExist: true);
            blip.Embed = PptxImages.EmbedImagePart(slidePart, iconSrc);
        }

        var blipFill = new P.BlipFill(blip, new A.Stretch(new A.FillRectangle()));

        var shapeProperties = new P.ShapeProperties(
            transform,
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

        return new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = id + 1, Name = name },
                new P.NonVisualPictureDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            blipFill,
            shapeProperties);
    }

    // ---- list / get ---------------------------------------------------------

    /// <summary>Every embedded object across the deck, in (slide, ordinal) order.</summary>
    public static List<EmbeddedObject> List(PresentationPart presentation)
    {
        var result = new List<EmbeddedObject>();
        var slides = PptxDoc.Slides(presentation);
        for (var i = 0; i < slides.Count; i++)
        {
            var slideIndex = i + 1;
            foreach (var embed in EmbedsOn(slides[i].Part))
            {
                result.Add(Project(slides[i].Part, slideIndex, embed));
            }
        }

        return result;
    }

    /// <summary>The number of extractable embedded objects on one slide (used by convert to report the loss).</summary>
    public static int SlideEmbedCount(SlidePart slidePart) => EmbedsOn(slidePart).Count;

    /// <summary>The embedded objects on one slide, projected for the structure view.</summary>
    public static List<EmbeddedObject> SlideEmbeds(SlidePart slidePart, int slideIndex) =>
        EmbedsOn(slidePart).Select(e => Project(slidePart, slideIndex, e)).ToList();

    /// <summary>The metadata of one addressed embed (name, mediaType, size) — never the bytes.</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var embed = Resolve(slidePart, address);
        var meta = Project(slidePart, address.SlideIndex, embed);
        return new
        {
            Path = meta.Path,
            Name = meta.Name,
            MediaType = meta.MediaType,
            Size = meta.Size,
            Container = meta.Container,
        };
    }

    private static EmbeddedObject Project(SlidePart slidePart, int slideIndex, EmbedRef embed)
    {
        var name = embed.Ole.Name?.Value;
        if (string.IsNullOrWhiteSpace(name))
        {
            name = PptxDoc.NonVisualProps(embed.Frame)?.Name?.Value;
        }

        return new EmbeddedObject(
            Path: Units.Inv($"/slide[{slideIndex}]/embed[@id={embed.HostId}]"),
            Name: string.IsNullOrWhiteSpace(name) ? Units.Inv($"Object {embed.HostId}") : name!,
            MediaType: embed.Part.ContentType,
            Size: PayloadSize(embed.Part),
            Container: Units.Inv($"/slide[{slideIndex}]"));
    }

    // ---- extract ------------------------------------------------------------

    /// <summary>
    /// Writes the addressed embed's payload to <paramref name="destPath"/> (already
    /// sandbox-resolved). The source document is not modified. Returns the canonical
    /// embed path so the op echoes a stable target.
    /// </summary>
    public static string Extract(PresentationPart presentation, PptxAddress address, string destPath)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var embed = Resolve(slidePart, address);

        if (Path.GetDirectoryName(destPath) is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
        }

        using (var source = embed.Part.GetStream())
        using (var destination = File.Create(destPath))
        {
            source.CopyTo(destination);
        }

        return Units.Inv($"/slide[{address.SlideIndex}]/embed[@id={embed.HostId}]");
    }

    // ---- remove -------------------------------------------------------------

    /// <summary>Removes the addressed embed (its host frame and the orphaned payload part).</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var embed = Resolve(slidePart, address);
        var canonical = Units.Inv($"/slide[{address.SlideIndex}]/embed[@id={embed.HostId}]");

        // Drop the payload part so removing the frame does not orphan it.
        slidePart.DeletePart(embed.Part);
        embed.Frame.Remove();
        return canonical;
    }

    // ---- resolution ---------------------------------------------------------

    /// <summary>The embedded objects on one slide, in document (ordinal) order.</summary>
    private static List<EmbedRef> EmbedsOn(SlidePart slidePart)
    {
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var result = new List<EmbedRef>();
        var ordinal = 0;
        foreach (var frame in tree.Elements<P.GraphicFrame>())
        {
            if (frame.Graphic?.GraphicData is not { } data ||
                !string.Equals(data.Uri?.Value, OleNamespace, StringComparison.Ordinal))
            {
                continue;
            }

            var ole = data.GetFirstChild<P.OleObject>();
            if (ole?.Id?.Value is not { } relId)
            {
                continue;
            }

            // Only embeds aioffice can hand back as files: an embedded package part.
            // Linked OLE objects (r:id on an external relationship) carry no bytes.
            if (!slidePart.TryGetPartById(relId, out var part) || part is not EmbeddedPackagePart packagePart)
            {
                continue;
            }

            ordinal++;
            var hostId = PptxDoc.NonVisualProps(frame)?.Id?.Value ?? 0;
            result.Add(new EmbedRef(ordinal, frame, ole, packagePart, hostId));
        }

        return result;
    }

    private static EmbedRef Resolve(SlidePart slidePart, PptxAddress address)
    {
        var embeds = EmbedsOn(slidePart);
        EmbedRef? match = address.EmbedId is { } id
            ? embeds.FirstOrDefault(e => e.HostId == id)
            : embeds.FirstOrDefault(e => e.Ordinal == address.EmbedOrdinal);

        if (match is not null)
        {
            return match;
        }

        var what = address.EmbedId is { } eid ? $"id {eid}" : $"index {address.EmbedOrdinal}";
        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No embedded object with {what} on slide {address.SlideIndex}; it has {embeds.Count} embed(s).",
            "Run 'aioffice read <file> --view embeds' to list embedded objects and their canonical paths.",
            candidates: [.. embeds.Take(10).Select(e => Units.Inv($"/slide[{address.SlideIndex}]/embed[@id={e.HostId}]"))]);
    }

    /// <summary>The payload byte count of an embedded package part (read without keeping it in memory).</summary>
    private static long PayloadSize(EmbeddedPackagePart part)
    {
        using var stream = part.GetStream();
        return stream.Length;
    }
}

/// <summary>Sniffs a content type from a file's header bytes (extension is a fallback only).</summary>
internal static class MediaType
{
    /// <summary>The generic binary content type when no signature matches.</summary>
    public const string OctetStream = "application/octet-stream";

    /// <summary>Content type of an embedded source: header-sniffed, then by extension, then octet-stream.</summary>
    public static string Sniff(string file)
    {
        byte[] header;
        using (var stream = File.OpenRead(file))
        {
            header = new byte[Math.Min(8, stream.Length)];
            _ = stream.Read(header, 0, header.Length);
        }

        // ZIP-family (xlsx/docx/pptx/odt/jar/zip) all start "PK\x03\x04"; the OOXML
        // ones are distinguished by extension since the magic is identical.
        if (header.Length >= 4 && header[0] == 0x50 && header[1] == 0x4B && header[2] == 0x03 && header[3] == 0x04)
        {
            return ByExtension(file) ?? "application/zip";
        }

        if (header.Length >= 5 && header[0] == 0x25 && header[1] == 0x50 && header[2] == 0x44 && header[3] == 0x46 && header[4] == 0x2D)
        {
            return "application/pdf"; // "%PDF-"
        }

        if (header.Length >= 8 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
        {
            return "image/png";
        }

        if (header.Length >= 3 && header[0] == 0xFF && header[1] == 0xD8 && header[2] == 0xFF)
        {
            return "image/jpeg";
        }

        if (header.Length >= 8 &&
            header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0)
        {
            // Legacy OLE compound file (old .doc/.xls/.ppt): sniff the extension.
            return ByExtension(file) ?? OctetStream;
        }

        return ByExtension(file) ?? OctetStream;
    }

    private static string? ByExtension(string file) =>
        Path.GetExtension(file).ToLowerInvariant() switch
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
            ".doc" => "application/msword",
            ".xls" => "application/vnd.ms-excel",
            ".ppt" => "application/vnd.ms-powerpoint",
            _ => null,
        };
}
