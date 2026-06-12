using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>A top-level slide shape with its stable id and 1-based ordinal.</summary>
internal sealed record ShapeView(OpenXmlCompositeElement Element, uint Id, int Ordinal, string Kind, string Name)
{
    public string CanonicalPath(int slideIndex) => Units.Inv($"/slide[{slideIndex}]/shape[@id={Id}]");

    public string OrdinalPath(int slideIndex) => Units.Inv($"/slide[{slideIndex}]/shape[{Ordinal}]");
}

/// <summary>Shape geometry in EMU.</summary>
internal readonly record struct GeometryEmu(long X, long Y, long Cx, long Cy);

/// <summary>Package plumbing shared by every pptx verb: open, enumerate, resolve, extract.</summary>
internal static class PptxDoc
{
    /// <summary>Copies the file into an expandable stream so edits are atomic (write back only on success).</summary>
    public static MemoryStream LoadStream(string file)
    {
        if (!File.Exists(file))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"File not found: {file}",
                "Check the path spelling, or run 'aioffice create' to make a new document.");
        }

        var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(file));
        stream.Position = 0;
        return stream;
    }

    public static PresentationDocument Open(MemoryStream stream, bool editable, string file)
    {
        try
        {
            return PresentationDocument.Open(stream, editable);
        }
        catch (Exception ex) when (ex is not AiofficeException)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                $"Not a readable .pptx package: {Path.GetFileName(file)} ({ex.Message})",
                "Re-export the file from PowerPoint/Keynote, or restore a snapshot with 'aioffice snapshot restore'.",
                innerException: ex);
        }
    }

    public static PresentationPart RequirePresentationPart(PresentationDocument doc, string file)
    {
        return doc.PresentationPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            $"The package has no presentation part: {Path.GetFileName(file)}",
            "The file is not a valid presentation; re-export it from PowerPoint/Keynote.");
    }

    /// <summary>Slide parts in show order (the order of p:sldIdLst, not part names).</summary>
    public static List<(P.SlideId Id, SlidePart Part)> Slides(PresentationPart presentation)
    {
        var result = new List<(P.SlideId, SlidePart)>();
        var list = presentation.Presentation?.SlideIdList;
        if (list is null)
        {
            return result;
        }

        foreach (var slideId in list.Elements<P.SlideId>())
        {
            if (slideId.RelationshipId?.Value is { } relId &&
                presentation.TryGetPartById(relId, out var part) &&
                part is SlidePart slidePart)
            {
                result.Add((slideId, slidePart));
            }
        }

        return result;
    }

    public static SlidePart ResolveSlide(PresentationPart presentation, int index1, string raw)
    {
        var slides = Slides(presentation);
        if (index1 >= 1 && index1 <= slides.Count)
        {
            return slides[index1 - 1].Part;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"Slide index {index1} is out of range in '{raw}'; the deck has {slides.Count} slide(s).",
            "Run 'aioffice read <file> --view outline' to list slides, or query 'slide' for canonical paths.",
            candidates: [.. Enumerable.Range(1, Math.Min(slides.Count, 10)).Select(i => Units.Inv($"/slide[{i}]"))]);
    }

    public static P.ShapeTree RequireShapeTree(SlidePart slidePart)
    {
        return slidePart.Slide?.CommonSlideData?.ShapeTree ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The slide has no shape tree (p:spTree).",
            "The slide part is malformed; re-export the file or restore a snapshot.");
    }

    /// <summary>Top-level shapes of a slide in document order, ordinals starting at 1.</summary>
    public static List<ShapeView> Shapes(SlidePart slidePart)
    {
        var views = new List<ShapeView>();
        var ordinal = 0;
        foreach (var element in RequireShapeTree(slidePart).ChildElements)
        {
            var kind = element switch
            {
                P.Shape => "shape",
                P.Picture => "picture",
                P.GraphicFrame => "graphicFrame",
                P.GroupShape => "group",
                P.ConnectionShape => "connector",
                _ => null,
            };

            if (kind is null)
            {
                continue;
            }

            ordinal++;
            var composite = (OpenXmlCompositeElement)element;
            var props = NonVisualProps(composite);
            views.Add(new ShapeView(
                composite,
                props?.Id?.Value ?? 0,
                ordinal,
                kind,
                props?.Name?.Value ?? string.Empty));
        }

        return views;
    }

    public static ShapeView ResolveShape(SlidePart slidePart, PptxAddress address)
    {
        var shapes = Shapes(slidePart);
        ShapeView? match = address.ShapeId is { } id
            ? shapes.FirstOrDefault(s => s.Id == id)
            : shapes.FirstOrDefault(s => s.Ordinal == address.ShapeOrdinal);

        if (match is not null)
        {
            return match;
        }

        var what = address.ShapeId is { } sid ? $"id {sid}" : $"index {address.ShapeOrdinal}";
        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No shape with {what} on slide {address.SlideIndex}; it has {shapes.Count} shape(s).",
            "Run 'aioffice query <file> shape' to list canonical shape paths.",
            candidates: [.. shapes.Take(10).Select(s => s.CanonicalPath(address.SlideIndex))]);
    }

    public static A.Paragraph ResolveParagraph(ShapeView view, PptxAddress address)
    {
        var paragraphs = (view.Element as P.Shape)?.TextBody?.Elements<A.Paragraph>().ToList() ?? [];
        var index = address.ParagraphIndex!.Value;
        if (index >= 1 && index <= paragraphs.Count)
        {
            return paragraphs[index - 1];
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No paragraph {index} in shape @id={view.Id}; it has {paragraphs.Count} paragraph(s).",
            "Run 'aioffice get <file> <shape path>' to inspect the shape's text first.",
            candidates: [.. Enumerable.Range(1, Math.Min(paragraphs.Count, 10))
                .Select(i => Units.Inv($"{view.CanonicalPath(address.SlideIndex)}/p[{i}]"))]);
    }

    public static P.NonVisualDrawingProperties? NonVisualProps(OpenXmlCompositeElement element) => element switch
    {
        P.Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties,
        P.Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties,
        P.GraphicFrame g => g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
        P.GroupShape g => g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
        P.ConnectionShape c => c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties,
        _ => null,
    };

    /// <summary>Paragraph-joined text of a shape; empty for shapes without a text body.</summary>
    public static string ShapeText(OpenXmlCompositeElement element)
    {
        if (element is not P.Shape shape || shape.TextBody is null)
        {
            return string.Empty;
        }

        return string.Join('\n', shape.TextBody.Elements<A.Paragraph>().Select(ParagraphText));
    }

    public static string ParagraphText(A.Paragraph paragraph)
    {
        var parts = new List<string>();
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case A.Run run when run.Text is { } text:
                    parts.Add(text.Text);
                    break;
                case A.Break:
                    parts.Add("\n");
                    break;
            }
        }

        return string.Concat(parts);
    }

    /// <summary>Solid RRGGBB fill of a shape, when one is set with an explicit RGB color.</summary>
    public static string? FillHex(OpenXmlCompositeElement element)
    {
        var properties = (element as P.Shape)?.ShapeProperties
            ?? (element as P.Picture)?.ShapeProperties
            ?? (element as P.ConnectionShape)?.ShapeProperties;
        return properties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();
    }

    public static GeometryEmu? Geometry(OpenXmlCompositeElement element)
    {
        switch (element)
        {
            case P.Shape s when s.ShapeProperties?.Transform2D is { } t:
                return FromTransform(t.Offset, t.Extents);
            case P.Picture p when p.ShapeProperties?.Transform2D is { } t:
                return FromTransform(t.Offset, t.Extents);
            case P.ConnectionShape c when c.ShapeProperties?.Transform2D is { } t:
                return FromTransform(t.Offset, t.Extents);
            case P.GraphicFrame g when g.Transform is { } t:
                return FromTransform(t.Offset, t.Extents);
            case P.GroupShape g when g.GroupShapeProperties?.TransformGroup is { } t:
                return FromTransform(t.Offset, t.Extents);
            default:
                return null;
        }
    }

    private static GeometryEmu? FromTransform(A.Offset? offset, A.Extents? extents)
    {
        if (offset is null || extents is null)
        {
            return null;
        }

        return new GeometryEmu(
            offset.X?.Value ?? 0,
            offset.Y?.Value ?? 0,
            extents.Cx?.Value ?? 0,
            extents.Cy?.Value ?? 0);
    }

    /// <summary>Next free shape id on a slide (id 1 is the root group).</summary>
    public static uint NextShapeId(P.ShapeTree tree)
    {
        uint max = 1;
        foreach (var props in tree.Descendants<P.NonVisualDrawingProperties>())
        {
            if (props.Id?.Value is { } id && id > max)
            {
                max = id;
            }
        }

        return max + 1;
    }

    /// <summary>Next free p:sldId value (the spec floor is 256).</summary>
    public static uint NextSlideId(P.SlideIdList list)
    {
        uint max = 255;
        foreach (var slideId in list.Elements<P.SlideId>())
        {
            if (slideId.Id?.Value is { } id && id > max)
            {
                max = id;
            }
        }

        return max + 1;
    }
}
