using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>A top-level shape (of a slide, master or layout) with its stable id and 1-based ordinal.</summary>
internal sealed record ShapeView(OpenXmlCompositeElement Element, uint Id, int Ordinal, string Kind, string Name)
{
    public string CanonicalPath(int slideIndex) => CanonicalPathIn(Units.Inv($"/slide[{slideIndex}]"));

    public string OrdinalPath(int slideIndex) => OrdinalPathIn(Units.Inv($"/slide[{slideIndex}]"));

    /// <summary>Stable-id path beneath any container (/slide[2], /master[1]/layout[2], …).</summary>
    public string CanonicalPathIn(string containerPath) => Units.Inv($"{containerPath}/shape[@id={Id}]");

    public string OrdinalPathIn(string containerPath) => Units.Inv($"{containerPath}/shape[{Ordinal}]");
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

        FileSizeGuard.Ensure(file); // file_too_large before any expensive open

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
        return slidePart.Slide?.CommonSlideData?.ShapeTree ?? throw MissingShapeTree("slide");
    }

    public static P.ShapeTree RequireShapeTree(SlideMasterPart masterPart)
    {
        return masterPart.SlideMaster?.CommonSlideData?.ShapeTree ?? throw MissingShapeTree("slide master");
    }

    public static P.ShapeTree RequireShapeTree(SlideLayoutPart layoutPart)
    {
        return layoutPart.SlideLayout?.CommonSlideData?.ShapeTree ?? throw MissingShapeTree("slide layout");
    }

    private static AiofficeException MissingShapeTree(string partKind) => new(
        ErrorCodes.FormatCorrupt,
        $"The {partKind} has no shape tree (p:spTree).",
        $"The {partKind} part is malformed; re-export the file or restore a snapshot.");

    /// <summary>Top-level shapes of a slide in document order, ordinals starting at 1.</summary>
    public static List<ShapeView> Shapes(SlidePart slidePart) => Shapes(RequireShapeTree(slidePart));

    /// <summary>Top-level shapes of any shape tree (slide, master or layout), ordinals starting at 1.</summary>
    public static List<ShapeView> Shapes(P.ShapeTree tree)
    {
        var views = new List<ShapeView>();
        var ordinal = 0;
        foreach (var element in tree.ChildElements)
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
        return ResolveShape(
            Shapes(slidePart),
            address,
            address.CanonicalContainerPath,
            Units.Inv($"on slide {address.SlideIndex}"));
    }

    /// <summary>Resolves a shape segment inside any container's shape list (slide, master or layout).</summary>
    public static ShapeView ResolveShape(List<ShapeView> shapes, PptxAddress address, string containerPath, string containerLabel)
    {
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
            $"No shape with {what} {containerLabel}; it has {shapes.Count} shape(s).",
            "Run 'aioffice query <file> shape' to list canonical shape paths.",
            candidates: [.. shapes.Take(10).Select(s => s.CanonicalPathIn(containerPath))]);
    }

    /// <summary>Master parts in p:sldMasterIdLst order, indices starting at 1.</summary>
    public static List<(int Index, SlideMasterPart Part)> Masters(PresentationPart presentation)
    {
        var result = new List<(int, SlideMasterPart)>();
        var list = presentation.Presentation?.SlideMasterIdList;
        if (list is not null)
        {
            foreach (var masterId in list.Elements<P.SlideMasterId>())
            {
                if (masterId.RelationshipId?.Value is { } relId &&
                    presentation.TryGetPartById(relId, out var part) &&
                    part is SlideMasterPart masterPart)
                {
                    result.Add((result.Count + 1, masterPart));
                }
            }

            return result;
        }

        foreach (var masterPart in presentation.SlideMasterParts)
        {
            result.Add((result.Count + 1, masterPart));
        }

        return result;
    }

    /// <summary>Layout parts of a master in p:sldLayoutIdLst order, indices starting at 1.</summary>
    public static List<(int Index, SlideLayoutPart Part)> Layouts(SlideMasterPart masterPart)
    {
        var result = new List<(int, SlideLayoutPart)>();
        var list = masterPart.SlideMaster?.SlideLayoutIdList;
        if (list is not null)
        {
            foreach (var layoutId in list.Elements<P.SlideLayoutId>())
            {
                if (layoutId.RelationshipId?.Value is { } relId &&
                    masterPart.TryGetPartById(relId, out var part) &&
                    part is SlideLayoutPart layoutPart)
                {
                    result.Add((result.Count + 1, layoutPart));
                }
            }

            return result;
        }

        foreach (var layoutPart in masterPart.SlideLayoutParts)
        {
            result.Add((result.Count + 1, layoutPart));
        }

        return result;
    }

    public static SlideMasterPart ResolveMaster(PresentationPart presentation, int index1, string raw)
    {
        var masters = Masters(presentation);
        if (index1 >= 1 && index1 <= masters.Count)
        {
            return masters[index1 - 1].Part;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"Master index {index1} is out of range in '{raw}'; the deck has {masters.Count} master(s).",
            "Run 'aioffice read <file> --view structure' to list masters and layouts.",
            candidates: [.. Enumerable.Range(1, Math.Min(masters.Count, 10)).Select(i => Units.Inv($"/master[{i}]"))]);
    }

    public static SlideLayoutPart ResolveLayout(SlideMasterPart masterPart, int masterIndex, int index1, string raw)
    {
        var layouts = Layouts(masterPart);
        if (index1 >= 1 && index1 <= layouts.Count)
        {
            return layouts[index1 - 1].Part;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"Layout index {index1} is out of range in '{raw}'; master {masterIndex} has {layouts.Count} layout(s).",
            "Run 'aioffice read <file> --view structure' to list masters and layouts.",
            candidates: [.. Enumerable.Range(1, Math.Min(layouts.Count, 10))
                .Select(i => Units.Inv($"/master[{masterIndex}]/layout[{i}]"))]);
    }

    /// <summary>The canonical /master[m]/layout[l] path of the layout a slide uses, when resolvable.</summary>
    public static string? LayoutPathOf(PresentationPart presentation, SlidePart slidePart)
    {
        var layout = slidePart.SlideLayoutPart;
        if (layout is null)
        {
            return null;
        }

        foreach (var (masterIndex, masterPart) in Masters(presentation))
        {
            foreach (var (layoutIndex, layoutPart) in Layouts(masterPart))
            {
                if (layoutPart.Uri == layout.Uri)
                {
                    return Units.Inv($"/master[{masterIndex}]/layout[{layoutIndex}]");
                }
            }
        }

        return null;
    }

    /// <summary>The layout's p:cSld name ("Blank", "Title Slide", …), when set.</summary>
    public static string? LayoutName(SlideLayoutPart layoutPart) =>
        layoutPart.SlideLayout?.CommonSlideData?.Name?.Value;

    /// <summary>The layout's schema type token ("blank", "titleOnly", …), when set.</summary>
    public static string? LayoutType(SlideLayoutPart layoutPart) =>
        layoutPart.SlideLayout?.Type?.InnerText;

    /// <summary>
    /// The placeholder type token of a shape ("title", "body", "ctrTitle", …);
    /// "body" when the placeholder exists with no explicit type (the schema
    /// default), null when the shape is not a placeholder.
    /// </summary>
    public static string? PlaceholderType(OpenXmlCompositeElement element)
    {
        var placeholder = element switch
        {
            P.Shape s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
            P.Picture p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
            P.GraphicFrame g => g.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
            P.GroupShape g => g.NonVisualGroupShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
            P.ConnectionShape c => c.NonVisualConnectionShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
            _ => null,
        };

        return placeholder is null ? null : placeholder.Type?.InnerText ?? "body";
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

    /// <summary>Solid RRGGBB background of a slide, when one is set with an explicit RGB color.</summary>
    public static string? BackgroundHex(SlidePart slidePart) => slidePart.Slide?.CommonSlideData?.Background?
        .BackgroundProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();

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
