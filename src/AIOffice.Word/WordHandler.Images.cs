using System.Buffers.Binary;
using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    private const long EmuPerCm = 360_000;
    private const double EmuPerPixel = 9_525; // 96 dpi

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body","type":"image","props":{"src":…,"width":"10cm"}}</c>:
    /// embeds a PNG/JPEG (sniffed by header, never by extension) as an inline
    /// drawing in a new paragraph. The src is workspace-relative and MUST
    /// resolve through the sandbox.
    /// </summary>
    private static object ApplyAddImage(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        var author = session.ResolveAuthor(props);

        var src = props["src"] is { } srcNode ? NodeToString(srcNode) : null;
        if (string.IsNullOrWhiteSpace(src))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type image needs props.src (a workspace-relative image path).",
                "Example: {\"op\":\"add\",\"path\":\"/body\",\"type\":\"image\",\"props\":{\"src\":\"logo.png\",\"width\":\"10cm\"}}.");
        }

        // SECURITY: the only road to the bytes is through the sandbox.
        var resolved = session.Workspace.Resolve(src, mustExist: true);
        if (!File.Exists(resolved))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"Image not found: {src}",
                "Check the path; it is resolved relative to the workspace root.");
        }

        var bytes = File.ReadAllBytes(resolved);
        var (format, pixelWidth, pixelHeight) = SniffImage(bytes, src);

        var widthEmu = props["width"] is { } w ? ParseLengthEmu("width", NodeToString(w)) : (long?)null;
        var heightEmu = props["height"] is { } h ? ParseLengthEmu("height", NodeToString(h)) : (long?)null;
        var aspect = pixelWidth > 0 && pixelHeight > 0 ? (double)pixelHeight / pixelWidth : 1.0;
        var cx = widthEmu ?? (heightEmu is { } he ? (long)Math.Round(he / aspect) : (long)Math.Round(pixelWidth * EmuPerPixel));
        var cy = heightEmu ?? (long)Math.Round(cx * aspect);
        if (cx <= 0 || cy <= 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Image dimensions must be positive; computed {cx}x{cy} EMU.",
                "Pass a positive width/height like \"10cm\", \"4in\" or \"300px\".");
        }

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
                $"Cannot add an image {position} {anchor.CanonicalPath} ({anchor.Type}).",
                "Add images inside /body (or a tc/header/footer), or before/after an existing paragraph.");
        }

        var main = doc.MainDocumentPart!;
        var imagePart = main.AddImagePart(format == "png" ? ImagePartType.Png : ImagePartType.Jpeg);
        using (var stream = new MemoryStream(bytes, writable: false))
        {
            imagePart.FeedData(stream);
        }

        var relId = main.GetIdOfPart(imagePart);
        var docPrId = NextDrawingId(doc);
        var name = Path.GetFileName(src);
        var alt = props["alt"] is { } altNode ? NodeToString(altNode)
            : props["descr"] is { } descrNode ? NodeToString(descrNode)
            : null;

        var paragraph = new Paragraph(new Run(BuildInlineDrawing(relId, docPrId, name, cx, cy, alt)));
        if (session.Track)
        {
            RequireBodyScope(anchor.CanonicalPath, "add");
            MarkParagraphInserted(doc, paragraph, author);
        }

        var canonical = Insert(doc, anchor, paragraph, position);
        return new
        {
            op = "add",
            type = "image",
            path = canonical,
            format,
            widthCm = EmuToCm(cx),
            heightCm = EmuToCm(cy),
        };
    }

    /// <summary>The standard wp:inline &gt; pic:pic markup Word writes for embedded images.</summary>
    private static Drawing BuildInlineDrawing(string relId, uint docPrId, string name, long cx, long cy, string? alt = null) => new(
        new Dw.Inline(
            new Dw.Extent { Cx = cx, Cy = cy },
            new Dw.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new Dw.DocProperties { Id = docPrId, Name = name, Description = alt is { Length: > 0 } ? alt : null },
            new Dw.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new Pic.Picture(
                        new Pic.NonVisualPictureProperties(
                            new Pic.NonVisualDrawingProperties { Id = 0U, Name = name },
                            new Pic.NonVisualPictureDrawingProperties()),
                        new Pic.BlipFill(
                            new A.Blip { Embed = relId },
                            new A.Stretch(new A.FillRectangle())),
                        new Pic.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = cx, Cy = cy }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        });

    /// <summary>The next free wp:docPr id across body, headers and footers.</summary>
    private static uint NextDrawingId(WordprocessingDocument doc)
    {
        var max = 0U;
        var roots = new List<OpenXmlElement>();
        if (doc.MainDocumentPart?.Document is { } document)
        {
            roots.Add(document);
        }

        roots.AddRange(WordAddress.HeaderFooterRoots(doc).Select(r => r.Element));
        foreach (var root in roots)
        {
            foreach (var docPr in root.Descendants<Dw.DocProperties>())
            {
                if (docPr.Id is { } id && id.Value > max)
                {
                    max = id.Value;
                }
            }
        }

        return max + 1;
    }

    // -------------------------------------------------------------- sniffing

    /// <summary>Identifies PNG/JPEG by magic bytes and reads the pixel size from the header.</summary>
    private static (string Format, int Width, int Height) SniffImage(byte[] bytes, string src)
    {
        if (bytes.Length >= 24 &&
            bytes.AsSpan(0, 8).SequenceEqual((byte[])[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]))
        {
            // IHDR is always the first chunk: width/height big-endian at 16/20.
            var width = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(16, 4));
            var height = BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(20, 4));
            return ("png", width, height);
        }

        if (bytes.Length >= 4 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
        {
            return ("jpeg", JpegSize(bytes, src).Width, JpegSize(bytes, src).Height);
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"'{src}' is not a PNG or JPEG (sniffed by file header, not extension).",
            "Convert the image to PNG or JPEG first; other formats arrive in a later milestone.");
    }

    /// <summary>Walks JPEG markers to the first SOF frame for the pixel size.</summary>
    private static (int Width, int Height) JpegSize(byte[] bytes, string src)
    {
        var i = 2;
        while (i + 9 < bytes.Length)
        {
            if (bytes[i] != 0xFF)
            {
                i++;
                continue;
            }

            var marker = bytes[i + 1];
            if (marker == 0xFF)
            {
                i++;
                continue;
            }

            // Standalone markers without a length payload.
            if (marker is 0xD8 or 0x01 or (>= 0xD0 and <= 0xD7))
            {
                i += 2;
                continue;
            }

            var length = (bytes[i + 2] << 8) | bytes[i + 3];

            // SOF0..SOF15 (minus DHT/JPG/DAC) carry the frame dimensions.
            if (marker is >= 0xC0 and <= 0xCF and not (0xC4 or 0xC8 or 0xCC))
            {
                var height = (bytes[i + 5] << 8) | bytes[i + 6];
                var width = (bytes[i + 7] << 8) | bytes[i + 8];
                return (width, height);
            }

            i += 2 + length;
        }

        throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            $"'{src}' looks like a JPEG but has no readable frame header.",
            "Re-export the image, or convert it to PNG.");
    }

    // ----------------------------------------------------------------- units

    /// <summary>Parses a length ("10cm", "4in", "36pt", "300px", "360000emu"; bare number = cm) into EMU.</summary>
    private static long ParseLengthEmu(string key, string raw)
    {
        var text = raw.Trim().ToLowerInvariant();
        var (suffix, factorToEmu) = text switch
        {
            _ when text.EndsWith("emu", StringComparison.Ordinal) => ("emu", 1.0),
            _ when text.EndsWith("cm", StringComparison.Ordinal) => ("cm", (double)EmuPerCm),
            _ when text.EndsWith("mm", StringComparison.Ordinal) => ("mm", EmuPerCm / 10.0),
            _ when text.EndsWith("pt", StringComparison.Ordinal) => ("pt", 12_700.0),
            _ when text.EndsWith("in", StringComparison.Ordinal) => ("in", 914_400.0),
            _ when text.EndsWith("px", StringComparison.Ordinal) => ("px", EmuPerPixel),
            _ => (string.Empty, (double)EmuPerCm),
        };

        var numberText = suffix.Length == 0 ? text : text[..^suffix.Length].Trim();
        if (double.TryParse(numberText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) && number > 0)
        {
            return (long)Math.Round(number * factorToEmu);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Property '{key}' is not a valid length: '{raw}'.",
            "Use centimeters as a number (e.g. 10) or a suffixed string like \"10cm\", \"4in\", \"36pt\", \"300px\".");
    }

    private static double EmuToCm(long emu) => Math.Round((double)emu / EmuPerCm, 2);

    // ------------------------------------------------------------------- set

    /// <summary>
    /// <c>set</c> alt/descr on an image-carrying paragraph or run: writes the
    /// wp:docPr description (the alt text accessibility tools read).
    /// </summary>
    private static object ApplySetImageAlt(ResolvedNode node, System.Text.Json.Nodes.JsonObject props)
    {
        var altNode = props["alt"] ?? props["descr"];
        var alt = NodeToString(altNode);

        var docPr = node.Element.Descendants<Dw.DocProperties>().FirstOrDefault()
            ?? throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"{node.CanonicalPath} carries a drawing with no docPr to describe.",
                "Re-add the image with 'aioffice edit --add … --type image'.");

        docPr.Description = alt;
        return new { op = "set", path = node.CanonicalPath, type = "image", alt };
    }

    // ------------------------------------------------------------------- get

    /// <summary>get on a paragraph/run holding an inline image.</summary>
    private static Dictionary<string, object?> ImageProperties(OpenXmlElement element)
    {
        var extent = element.Descendants<Dw.Inline>().FirstOrDefault()?.Extent;
        var docPr = element.Descendants<Dw.DocProperties>().FirstOrDefault();
        return new Dictionary<string, object?>
        {
            ["kind"] = "image",
            ["widthCm"] = extent?.Cx?.Value is { } cx ? EmuToCm(cx) : null,
            ["heightCm"] = extent?.Cy?.Value is { } cy ? EmuToCm(cy) : null,
            ["name"] = docPr?.Name?.Value,
            ["alt"] = docPr?.Description?.Value,
            ["text"] = element.InnerText.Length > 0 ? element.InnerText : null,
            ["note"] = "The image bytes are embedded in the document; the original src path is not stored.",
        };
    }
}
