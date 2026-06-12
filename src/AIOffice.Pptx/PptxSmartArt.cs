using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>One SmartArt node: its text, 0-based depth and children in connection order.</summary>
internal sealed record SmartArtNode(string Text, int Level, IReadOnlyList<SmartArtNode> Children);

/// <summary>
/// SmartArt diagrams, read-only: a p:graphicFrame whose graphicData carries
/// dgm:relIds pointing at the diagram data/layout/colors/style parts. aioffice
/// reads the layout name and the node tree (data-part points wired by parOf
/// connections); every edit op is a typed unsupported_feature.
/// </summary>
internal static class PptxSmartArt
{
    /// <summary>The typed rejection every edit op on a smartart path raises.</summary>
    public static AiofficeException EditUnsupported(string path) => new(
        ErrorCodes.UnsupportedFeature,
        $"SmartArt is read-only; '{path}' cannot be edited.",
        "Recreate the content as shapes or a table (add shape/table ops), or edit the diagram in PowerPoint.");

    // ----- enumerate / resolve -------------------------------------------------

    /// <summary>The diagram data part a graphic frame references; null when the frame hosts no SmartArt.</summary>
    public static DiagramDataPart? DataPartOf(SlidePart slidePart, OpenXmlCompositeElement element)
    {
        if (element is not P.GraphicFrame frame ||
            frame.Graphic?.GraphicData?.GetFirstChild<Dgm.RelationshipIds>()?.DataPart?.Value is not { } relId)
        {
            return null;
        }

        return slidePart.TryGetPartById(relId, out var part) && part is DiagramDataPart dataPart ? dataPart : null;
    }

    /// <summary>SmartArt frames on a slide in paint order; indices are 1-based.</summary>
    public static List<(int Index, ShapeView View, DiagramDataPart Part)> List(SlidePart slidePart)
    {
        var result = new List<(int, ShapeView, DiagramDataPart)>();
        foreach (var view in PptxDoc.Shapes(slidePart))
        {
            if (DataPartOf(slidePart, view.Element) is { } part)
            {
                result.Add((result.Count + 1, view, part));
            }
        }

        return result;
    }

    /// <summary>The 1-based SmartArt index of a frame on its slide; null for non-diagrams.</summary>
    public static int? IndexOf(SlidePart slidePart, OpenXmlCompositeElement element) =>
        List(slidePart).Where(s => ReferenceEquals(s.View.Element, element))
            .Select(s => (int?)s.Index)
            .FirstOrDefault();

    /// <summary>Resolves /slide[i]/smartart[k] or throws invalid_path with candidates.</summary>
    public static (int Index, ShapeView View, DiagramDataPart Part) Resolve(SlidePart slidePart, PptxAddress address)
    {
        var diagrams = List(slidePart);
        var index = address.SmartArtIndex!.Value;
        if (index >= 1 && index <= diagrams.Count)
        {
            return diagrams[index - 1];
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            Units.Inv($"No smartart {index} on slide {address.SlideIndex}; it has {diagrams.Count} diagram(s)."),
            "SmartArt indices are 1-based per slide; run 'aioffice read <file> --view structure' to list them.",
            candidates: [.. diagrams.Take(10).Select(s => Units.Inv($"{address.CanonicalSlidePath}/smartart[{s.Index}]"))]);
    }

    // ----- read side ------------------------------------------------------------

    /// <summary>The diagram's layout name: the layout part's dgm:title, falling back to its uniqueId.</summary>
    public static string? LayoutName(SlidePart slidePart, OpenXmlCompositeElement element)
    {
        if (element is not P.GraphicFrame frame ||
            frame.Graphic?.GraphicData?.GetFirstChild<Dgm.RelationshipIds>()?.LayoutPart?.Value is not { } relId ||
            !slidePart.TryGetPartById(relId, out var part) ||
            part is not DiagramLayoutDefinitionPart layoutPart)
        {
            return null;
        }

        var definition = layoutPart.LayoutDefinition;
        var title = definition?.GetFirstChild<Dgm.Title>()?.Val?.Value;
        return string.IsNullOrEmpty(title) ? definition?.UniqueId?.Value : title;
    }

    /// <summary>The node tree of a diagram data part: doc-rooted points wired by parOf connections.</summary>
    public static List<SmartArtNode> Nodes(DiagramDataPart part)
    {
        var model = part.DataModelRoot;
        var points = model?.GetFirstChild<Dgm.PointList>()?.Elements<Dgm.Point>().ToList() ?? [];
        if (points.Count == 0)
        {
            return [];
        }

        var byId = new Dictionary<string, Dgm.Point>(StringComparer.Ordinal);
        foreach (var point in points)
        {
            if (point.ModelId?.Value is { } id)
            {
                byId[id] = point;
            }
        }

        var children = new Dictionary<string, List<(uint Order, string Id)>>(StringComparer.Ordinal);
        foreach (var connection in model?.GetFirstChild<Dgm.ConnectionList>()?.Elements<Dgm.Connection>() ?? [])
        {
            var type = connection.Type?.Value ?? Dgm.ConnectionValues.ParentOf;
            if (type != Dgm.ConnectionValues.ParentOf ||
                connection.SourceId?.Value is not { } source ||
                connection.DestinationId?.Value is not { } destination)
            {
                continue;
            }

            if (!children.TryGetValue(source, out var list))
            {
                children[source] = list = [];
            }

            list.Add((connection.SourcePosition?.Value ?? (uint)list.Count, destination));
        }

        List<SmartArtNode> Build(string parentId, int level)
        {
            if (!children.TryGetValue(parentId, out var list))
            {
                return [];
            }

            var nodes = new List<SmartArtNode>();
            foreach (var (_, id) in list.OrderBy(c => c.Order))
            {
                if (!byId.TryGetValue(id, out var point) || !IsTextNode(point))
                {
                    continue;
                }

                nodes.Add(new SmartArtNode(PointText(point), level, Build(id, level + 1)));
            }

            return nodes;
        }

        var root = points.FirstOrDefault(p => p.Type?.Value == Dgm.PointValues.Document);
        if (root?.ModelId?.Value is { } rootId)
        {
            return Build(rootId, 0);
        }

        // No doc point (foreign minimal data): flatten every text-bearing node point.
        return [.. points.Where(IsTextNode).Select(p => new SmartArtNode(PointText(p), 0, []))];
    }

    /// <summary>Node-like points carry user text; transitions/presentation points do not.</summary>
    private static bool IsTextNode(Dgm.Point point)
    {
        var type = point.Type?.Value ?? Dgm.PointValues.Node;
        return type == Dgm.PointValues.Node || type == Dgm.PointValues.Assistant;
    }

    private static string PointText(Dgm.Point point) => point.TextBody is null
        ? string.Empty
        : string.Join('\n', point.TextBody.Elements<A.Paragraph>().Select(PptxDoc.ParagraphText));

    private static int Count(IReadOnlyList<SmartArtNode> nodes) =>
        nodes.Sum(n => 1 + Count(n.Children));

    /// <summary>Every node text, depth-first (for query matching).</summary>
    public static string FlatText(DiagramDataPart part) => string.Join('\n', FlatLines(Nodes(part), indent: false));

    /// <summary>Node texts flattened depth-first, two spaces of indentation per level.</summary>
    public static List<string> IndentedLines(DiagramDataPart part) => FlatLines(Nodes(part), indent: true);

    private static List<string> FlatLines(IReadOnlyList<SmartArtNode> nodes, bool indent)
    {
        var lines = new List<string>();
        void Walk(IReadOnlyList<SmartArtNode> level)
        {
            foreach (var node in level)
            {
                if (node.Text.Length > 0)
                {
                    lines.Add(indent ? new string(' ', node.Level * 2) + node.Text : node.Text);
                }

                Walk(node.Children);
            }
        }

        Walk(nodes);
        return lines;
    }

    /// <summary>The `get` projection for /slide[i]/smartart[k].</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var (index, view, part) = Resolve(slidePart, address);
        var nodes = Nodes(part);
        var geometry = PptxDoc.Geometry(view.Element);
        return new
        {
            Path = Units.Inv($"/slide[{address.SlideIndex}]/smartart[{index}]"),
            ShapePath = view.CanonicalPath(address.SlideIndex),
            Slide = address.SlideIndex,
            Id = view.Id,
            Kind = "smartart",
            ReadOnly = true,
            Layout = LayoutName(slidePart, view.Element),
            NodeCount = Count(nodes),
            Texts = nodes.Select(Project).ToList(),
            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
        };
    }

    private static object Project(SmartArtNode node) => new
    {
        Text = node.Text,
        Level = node.Level,
        Children = node.Children.Count == 0 ? null : node.Children.Select(Project).ToList(),
    };

    /// <summary>The structure-view row for one diagram (path, layout, node count, indented texts).</summary>
    public static object StructureRow(SlidePart slidePart, int slideIndex, int index, ShapeView view, DiagramDataPart part) => new
    {
        Path = Units.Inv($"/slide[{slideIndex}]/smartart[{index}]"),
        ShapePath = view.CanonicalPath(slideIndex),
        Layout = LayoutName(slidePart, view.Element),
        NodeCount = Count(Nodes(part)),
        Texts = IndentedLines(part),
    };
}
