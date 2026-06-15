using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace AIOffice.Pptx;

/// <summary>
/// Embedded 3D models (glb/gltf), v1.3.0. The model bytes ride in a real
/// <see cref="Model3DReferenceRelationshipPart"/> (the Office 2017 3D-model part),
/// referenced from a host <c>p:pic</c> by a <c>p14:media</c> extension. The
/// picture's blip-fill shows the caller's <c>poster</c> image (a labeled grey box
/// when none is given), so PowerPoint versions without 3D support still render a
/// visible fallback while 2019+ can light up the embedded model. The host picture's
/// shape id is the model's stable id, so <c>/slide[i]/model3d[@id=N]</c> survives
/// insertions and removals.
///
/// This is the deliberately-robust embedding: the strongly-typed am3d / a16:model3D
/// graphic tree is fragile to keep OpenXmlValidator-clean, so aioffice stores the
/// model as a first-class 3D part behind a picture fallback (and reports an honest
/// <c>model3d_as_media</c> warning), favoring correctness over schema coverage.
/// </summary>
internal static class PptxModels
{
    /// <summary>The p14 media extension uri (shared with embedded video/audio); the embed points at the 3D part.</summary>
    private const string MediaExtensionUri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";

    private const string Office2010MainNs = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    private static readonly IReadOnlyList<string> AddPropKeys = ["src", "poster", "x", "y", "w", "h", "name"];

    /// <summary>The content-type prefix that marks an embedded part as a 3D model (model/gltf-binary, model/gltf+json).</summary>
    private const string ModelContentTypePrefix = "model/";

    /// <summary>Sniffed 3D-model identity: content type and the canonical file extension.</summary>
    private readonly record struct ModelInfo(string ContentType, string Extension, string Kind);

    /// <summary>One 3D model resolved on a slide: host picture, the model data part and the host shape id.</summary>
    private sealed record ModelRef(int Ordinal, P.Picture Picture, MediaDataPart Part, uint HostId, string Kind);

    // ---- add ----------------------------------------------------------------

    /// <summary>
    /// Embeds the sandbox-resolved <c>props.src</c> 3D model (glb/gltf) on the slide
    /// behind a poster picture fallback and returns the canonical model path plus a
    /// warning noting the media-backed embedding.
    /// </summary>
    public static (string Path, Warning Note) Add(PresentationDocument document, SlidePart slidePart, int slideIndex, JsonObject? props, Workspace workspace)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddPropKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown model3d prop '{key}'.",
                    "model3d props: src (required), poster, x, y, w, h, name. " +
                    "src points at a .glb or .gltf file inside the workspace.",
                    candidates: AddPropKeys);
            }
        }

        if (!props.TryGetPropertyValue("src", out var srcNode) || srcNode is null || J.ScalarText(srcNode).Trim().Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add model3d requires props.src.",
                "Point src at a 3D model inside the workspace, e.g. " +
                "{\"op\":\"add\",\"path\":\"/slide[1]\",\"type\":\"model3d\",\"props\":{\"src\":\"chair.glb\"}}.");
        }

        // Sandbox first: a src outside the workspace is sandbox_denied, never read.
        var src = workspace.Resolve(J.ScalarText(srcNode).Trim(), mustExist: true);
        var info = Sniff(src);
        var bytes = File.ReadAllBytes(src);

        // props.poster (optional): a PNG/JPEG sandbox-resolved before any mutation.
        string? posterRel = null;
        if (props.TryGetPropertyValue("poster", out var posterNode) && posterNode is not null && J.ScalarText(posterNode).Trim().Length > 0)
        {
            var poster = workspace.Resolve(J.ScalarText(posterNode).Trim(), mustExist: true);
            posterRel = PptxImages.EmbedImagePart(slidePart, poster);
        }

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);

        var x = props.TryGetPropertyValue("x", out var xNode) ? Units.ParseLengthEmu("x", xNode) : Units.CmToEmu(2);
        var y = props.TryGetPropertyValue("y", out var yNode) ? Units.ParseLengthEmu("y", yNode) : Units.CmToEmu(2);
        var w = props.TryGetPropertyValue("w", out var wNode) ? Units.ParseLengthEmu("w", wNode) : Units.CmToEmu(12);
        var h = props.TryGetPropertyValue("h", out var hNode) ? Units.ParseLengthEmu("h", hNode) : Units.CmToEmu(10);
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null && J.ScalarText(nameNode).Trim().Length > 0
            ? J.ScalarText(nameNode).Trim()
            : Path.GetFileName(src);

        // The verbatim model bytes live in a MediaDataPart with a model/* content
        // type (FeedData copies the bytes through, so the model round-trips exactly).
        // The p14:media embed makes PowerPoint 2019+ treat the picture as a 3D host.
        var modelPart = document.CreateMediaDataPart(info.ContentType, info.Extension);
        using (var stream = new MemoryStream(bytes, writable: false))
        {
            modelPart.FeedData(stream);
        }

        var embedRel = slidePart.AddMediaReferenceRelationship(modelPart).Id;
        tree.Append(BuildPicture(id, name, embedRel, posterRel, x, y, w, h));

        var note = new Warning(
            WarningCodes.Model3DAsMedia,
            $"The 3D model '{Path.GetFileName(src)}' was embedded as a 3D part behind a poster picture " +
            "fallback (so older PowerPoint still shows the poster); PowerPoint 2019+ can render the model.");
        return (Units.Inv($"/slide[{slideIndex}]/model3d[@id={id}]"), note);
    }

    /// <summary>Builds the p:pic hosting the 3D model: a poster blip-fill + a p14:media extension pointing at the 3D part.</summary>
    private static P.Picture BuildPicture(uint id, string name, string embedRel, string? posterRel, long x, long y, long w, long h)
    {
        var nonVisualProps = new P.NonVisualDrawingProperties { Id = id, Name = name };

        var appProps = new P.ApplicationNonVisualDrawingProperties();
        var extension = new P.ApplicationNonVisualDrawingPropertiesExtension(
            new P14.Media { Embed = embedRel })
        {
            Uri = MediaExtensionUri,
        };
        extension.GetFirstChild<P14.Media>()!.AddNamespaceDeclaration("p14", Office2010MainNs);
        appProps.Append(new P.ApplicationNonVisualDrawingPropertiesExtensionList(extension));

        var blip = new A.Blip();
        if (posterRel is not null)
        {
            blip.Embed = posterRel;
        }

        return new P.Picture(
            new P.NonVisualPictureProperties(
                nonVisualProps,
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                appProps),
            new P.BlipFill(blip, new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = w, Cy = h }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
    }

    // ---- list / get ---------------------------------------------------------

    /// <summary>The 3D models on one slide, in document (ordinal) order.</summary>
    private static List<ModelRef> ModelsOn(SlidePart slidePart)
    {
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var result = new List<ModelRef>();
        var ordinal = 0;
        foreach (var picture in tree.Elements<P.Picture>())
        {
            var part = Model3DPartOf(slidePart, picture);
            if (part is null)
            {
                continue;
            }

            ordinal++;
            var hostId = PptxDoc.NonVisualProps(picture)?.Id?.Value ?? 0;
            result.Add(new ModelRef(ordinal, picture, part, hostId, KindOf(part.ContentType)));
        }

        return result;
    }

    /// <summary>The model data part a picture's p14:media extension references, or null when the picture hosts no 3D model.</summary>
    private static MediaDataPart? Model3DPartOf(SlidePart slidePart, P.Picture picture)
    {
        // A 3D-model host picture carries a p14:media (and no audio/video file ref);
        // its embedded data part's content type is model/* (gltf-binary or gltf+json).
        var nvPr = picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
        if (nvPr is null ||
            nvPr.GetFirstChild<A.VideoFromFile>() is not null ||
            nvPr.GetFirstChild<A.AudioFromFile>() is not null)
        {
            return null;
        }

        var embedRel = nvPr.GetFirstChild<P.ApplicationNonVisualDrawingPropertiesExtensionList>()?
            .Descendants<P14.Media>().FirstOrDefault()?.Embed?.Value;
        if (embedRel is null)
        {
            return null;
        }

        var reference = slidePart.DataPartReferenceRelationships.FirstOrDefault(r => r.Id == embedRel);
        return reference?.DataPart is MediaDataPart part &&
            part.ContentType.StartsWith(ModelContentTypePrefix, StringComparison.OrdinalIgnoreCase)
            ? part
            : null;
    }

    /// <summary>True when the picture hosts an embedded 3D model (the SVG render hook).</summary>
    public static bool IsModelPicture(SlidePart slidePart, P.Picture picture) => Model3DPartOf(slidePart, picture) is not null;

    /// <summary>The 3D models on one slide, projected for the structure view.</summary>
    public static List<object> SlideModels(SlidePart slidePart, int slideIndex) =>
        ModelsOn(slidePart).Select(m => Project(slideIndex, m)).ToList();

    /// <summary>The number of embedded 3D models on one slide (used by convert to report the loss).</summary>
    public static int SlideModelCount(SlidePart slidePart) => ModelsOn(slidePart).Count;

    /// <summary>The detail projection of one addressed 3D model (kind, src name, geometry).</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var model = Resolve(slidePart, address);
        return Project(address.SlideIndex, model);
    }

    private static object Project(int slideIndex, ModelRef model)
    {
        var geometry = PptxDoc.Geometry(model.Picture);
        var name = PptxDoc.NonVisualProps(model.Picture)?.Name?.Value;
        return new
        {
            Path = Units.Inv($"/slide[{slideIndex}]/model3d[@id={model.HostId}]"),
            ShapePath = Units.Inv($"/slide[{slideIndex}]/shape[@id={model.HostId}]"),
            Slide = slideIndex,
            Id = model.HostId,
            Kind = "model3d",
            Geometry = model.Kind, // "glb" or "gltf"
            Src = string.IsNullOrWhiteSpace(name) ? Units.Inv($"Model {model.HostId}") : name,
            MediaType = model.Part.ContentType,
            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
        };
    }

    // ---- remove -------------------------------------------------------------

    /// <summary>Removes the addressed 3D model (its host picture and the 3D part when orphaned).</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var model = Resolve(slidePart, address);
        var canonical = Units.Inv($"/slide[{address.SlideIndex}]/model3d[@id={model.HostId}]");

        DeleteModelPartsFor(slidePart, model.Picture);
        model.Picture.Remove();
        return canonical;
    }

    /// <summary>Drops the model reference relationship (and the shared data part when orphaned) for a model-hosting picture being removed.</summary>
    public static void DeleteModelPartsFor(SlidePart slidePart, P.Picture picture)
    {
        if (Model3DPartOf(slidePart, picture) is not { } dataPart)
        {
            return;
        }

        var nvPr = picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
        var embedRel = nvPr?.GetFirstChild<P.ApplicationNonVisualDrawingPropertiesExtensionList>()?
            .Descendants<P14.Media>().FirstOrDefault()?.Embed?.Value;
        if (embedRel is { } relId &&
            slidePart.DataPartReferenceRelationships.FirstOrDefault(r => r.Id == relId) is { } reference)
        {
            slidePart.DeleteReferenceRelationship(reference);
        }

        // Delete the shared model data part only when no other slide references it.
        if (!PartStillReferenced(slidePart, dataPart))
        {
            slidePart.OpenXmlPackage.DeletePart(dataPart);
        }
    }

    private static bool PartStillReferenced(SlidePart removingFrom, MediaDataPart dataPart)
    {
        foreach (var part in removingFrom.OpenXmlPackage.GetPartsOfType<SlidePart>())
        {
            foreach (var reference in part.DataPartReferenceRelationships)
            {
                if (ReferenceEquals(reference.DataPart, dataPart))
                {
                    return true;
                }
            }
        }

        return false;
    }

    // ---- resolution ---------------------------------------------------------

    private static ModelRef Resolve(SlidePart slidePart, PptxAddress address)
    {
        var models = ModelsOn(slidePart);
        ModelRef? match = address.Model3DId is { } id
            ? models.FirstOrDefault(m => m.HostId == id)
            : models.FirstOrDefault(m => m.Ordinal == address.Model3DOrdinal);

        if (match is not null)
        {
            return match;
        }

        var what = address.Model3DId is { } mid ? $"id {mid}" : $"index {address.Model3DOrdinal}";
        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No 3D model with {what} on slide {address.SlideIndex}; it has {models.Count} model(s).",
            "Run 'aioffice read <file> --view structure' to list slide 3D models and their canonical paths.",
            candidates: [.. models.Take(10).Select(m => Units.Inv($"/slide[{address.SlideIndex}]/model3d[@id={m.HostId}]"))]);
    }

    // ---- sniff --------------------------------------------------------------

    /// <summary>Identifies glb (binary glTF) from its "glTF" magic header; gltf (JSON) is recognized by extension or a leading '{'.</summary>
    private static ModelInfo Sniff(string src)
    {
        var bytes = ReadHeader(src, 4);

        // glb: a 12-byte header beginning with the ASCII magic "glTF".
        if (bytes.Length >= 4 && bytes[0] == 'g' && bytes[1] == 'l' && bytes[2] == 'T' && bytes[3] == 'F')
        {
            return new ModelInfo("model/gltf-binary", "glb", "glb");
        }

        var extension = Path.GetExtension(src).ToLowerInvariant();
        if (extension == ".glb")
        {
            return new ModelInfo("model/gltf-binary", "glb", "glb");
        }

        if (extension == ".gltf" || (bytes.Length >= 1 && bytes[0] == '{'))
        {
            return new ModelInfo("model/gltf+json", "gltf", "gltf");
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"'{Path.GetFileName(src)}' is not a recognized 3D model (sniffed by header).",
            "Embed a binary glTF (.glb) or a glTF JSON (.gltf); convert other 3D formats first " +
            "(e.g. with Blender's glTF exporter).");
    }

    private static string KindOf(string contentType) =>
        contentType.Contains("binary", StringComparison.OrdinalIgnoreCase) ? "glb" : "gltf";

    private static byte[] ReadHeader(string file, int count)
    {
        using var stream = File.OpenRead(file);
        var buffer = new byte[Math.Min(count, (int)Math.Min(stream.Length, int.MaxValue))];
        var read = 0;
        while (read < buffer.Length)
        {
            var n = stream.Read(buffer, read, buffer.Length - read);
            if (n == 0)
            {
                break;
            }

            read += n;
        }

        return read == buffer.Length ? buffer : buffer[..read];
    }
}
