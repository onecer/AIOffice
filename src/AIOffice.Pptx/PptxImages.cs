using System.Buffers.Binary;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Picture shapes: embeds a PNG/JPEG file (header-sniffed, sandbox-resolved)
/// as an ImagePart plus a p:pic in the slide's shape tree. Missing extents are
/// derived from the image's pixel aspect ratio.
/// </summary>
internal static class PptxImages
{
    private static readonly IReadOnlyList<string> ImagePropKeys = ["src", "x", "y", "w", "h", "name"];

    /// <summary>Sniffed image identity: content type for the part and pixel dimensions.</summary>
    private readonly record struct ImageInfo(PartTypeInfo PartType, string Format, int PixelWidth, int PixelHeight);

    /// <summary>
    /// Embeds an already-sandbox-resolved PNG/JPEG file as an ImagePart on the slide
    /// (header-sniffed) and returns its relationship id. Shared by the picture add
    /// path and the embed-object preview (props.icon) path.
    /// </summary>
    public static string EmbedImagePart(SlidePart slidePart, string resolvedSrc)
    {
        var bytes = File.ReadAllBytes(resolvedSrc);
        var info = Sniff(bytes, resolvedSrc);
        var imagePart = slidePart.AddImagePart(info.PartType);
        using (var stream = new MemoryStream(bytes, writable: false))
        {
            imagePart.FeedData(stream);
        }

        return slidePart.GetIdOfPart(imagePart);
    }

    /// <summary>Adds the picture and returns its stable shape id.</summary>
    public static uint AddImage(SlidePart slidePart, JsonObject? props, Workspace workspace)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!ImagePropKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown image prop '{key}'.",
                    "Image props: src (required), x, y, w, h, name. Size from one of w/h keeps the aspect ratio.",
                    candidates: ImagePropKeys);
            }
        }

        if (!props.TryGetPropertyValue("src", out var srcNode) || srcNode is null || J.ScalarText(srcNode).Trim().Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add image requires props.src.",
                "Point src at a PNG or JPEG inside the workspace, e.g. " +
                "{\"op\":\"add\",\"path\":\"/slide[1]\",\"type\":\"image\",\"props\":{\"src\":\"logo.png\",\"w\":\"10cm\"}}.");
        }

        // Sandbox first: a path outside the workspace is sandbox_denied, never read.
        var src = workspace.Resolve(J.ScalarText(srcNode).Trim(), mustExist: true);
        var bytes = File.ReadAllBytes(src);
        var info = Sniff(bytes, src);

        var (w, h) = Extents(props, info);
        var x = props.TryGetPropertyValue("x", out var xNode) ? Units.ParseLengthEmu("x", xNode) : Units.CmToEmu(2.5);
        var y = props.TryGetPropertyValue("y", out var yNode) ? Units.ParseLengthEmu("y", yNode) : Units.CmToEmu(2.5);

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
            ? J.ScalarText(nameNode)
            : Path.GetFileName(src);

        var imagePart = slidePart.AddImagePart(info.PartType);
        using (var stream = new MemoryStream(bytes, writable: false))
        {
            imagePart.FeedData(stream);
        }

        tree.Append(new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = slidePart.GetIdOfPart(imagePart) },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = w, Cy = h }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })));
        return id;
    }

    /// <summary>
    /// Final extents in EMU: explicit w/h win; a single given side derives the
    /// other from the pixel aspect; neither means natural size at 96 dpi.
    /// </summary>
    private static (long W, long H) Extents(JsonObject props, ImageInfo info)
    {
        var w = props.TryGetPropertyValue("w", out var wNode) ? Units.ParseLengthEmu("w", wNode) : (long?)null;
        var h = props.TryGetPropertyValue("h", out var hNode) ? Units.ParseLengthEmu("h", hNode) : (long?)null;

        return (w, h) switch
        {
            ({ } width, { } height) => (width, height),
            ({ } width, null) => (width, (long)Math.Round(width * (double)info.PixelHeight / info.PixelWidth)),
            (null, { } height) => ((long)Math.Round(height * (double)info.PixelWidth / info.PixelHeight), height),
            _ => ((long)Math.Round(info.PixelWidth * Units.EmuPerPixel), (long)Math.Round(info.PixelHeight * Units.EmuPerPixel)),
        };
    }

    /// <summary>Identifies the format from file headers (extension is never trusted).</summary>
    private static ImageInfo Sniff(byte[] bytes, string src)
    {
        if (TrySniffPng(bytes) is { } png)
        {
            return png;
        }

        if (TrySniffJpeg(bytes) is { } jpeg)
        {
            return jpeg;
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"Only PNG and JPEG images can be embedded; '{Path.GetFileName(src)}' is neither (sniffed by header).",
            "Convert the image first, e.g. 'sips -s format png input.gif --out input.png', then add the converted file.");
    }

    private static ImageInfo? TrySniffPng(byte[] bytes)
    {
        ReadOnlySpan<byte> signature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        if (bytes.Length < 24 || !bytes.AsSpan(0, 8).SequenceEqual(signature))
        {
            return null;
        }

        // IHDR is mandated to be the first chunk: width/height at offsets 16/20.
        var width = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(16, 4));
        var height = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(20, 4));
        if (width < 1 || height < 1)
        {
            throw CorruptImage("the PNG header reports a non-positive size");
        }

        return new ImageInfo(ImagePartType.Png, "png", width, height);
    }

    private static ImageInfo? TrySniffJpeg(byte[] bytes)
    {
        if (bytes.Length < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8)
        {
            return null;
        }

        var i = 2;
        while (i + 9 < bytes.Length)
        {
            if (bytes[i] != 0xFF)
            {
                i++;
                continue;
            }

            var marker = bytes[i + 1];
            if (marker is 0xFF or 0x01 or (>= 0xD0 and <= 0xD8))
            {
                i += marker == 0xFF ? 1 : 2; // padding / standalone markers carry no length
                continue;
            }

            if (marker == 0xD9) // EOI before any frame header
            {
                break;
            }

            var length = (bytes[i + 2] << 8) | bytes[i + 3];
            var isFrameHeader = marker is (>= 0xC0 and <= 0xCF) and not (0xC4 or 0xC8 or 0xCC);
            if (isFrameHeader)
            {
                var height = (bytes[i + 5] << 8) | bytes[i + 6];
                var width = (bytes[i + 7] << 8) | bytes[i + 8];
                if (width < 1 || height < 1)
                {
                    throw CorruptImage("the JPEG frame header reports a non-positive size");
                }

                return new ImageInfo(ImagePartType.Jpeg, "jpeg", width, height);
            }

            i += 2 + Math.Max(length, 2);
        }

        throw CorruptImage("no JPEG frame header (SOF) was found");
    }

    private static AiofficeException CorruptImage(string detail) => new(
        ErrorCodes.InvalidArgs,
        $"The image file is unreadable: {detail}.",
        "Re-export the image as a standard PNG or JPEG and try again.");
}
