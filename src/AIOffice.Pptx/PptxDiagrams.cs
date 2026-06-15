using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// SmartArt CREATION (the read side lives in <see cref="PptxSmartArt"/>). Builds the
/// four diagram parts a real SmartArt graphic needs — a dgm:dataModel (the node tree,
/// wired by parOf connections), a layout reference to a BUILT-IN layout by its standard
/// uniqueId, plus a quick-style and a colors part — and the host p:graphicFrame carrying
/// the dgm:relIds. The result is OpenXmlValidator-clean and opens in PowerPoint, which
/// regenerates the diagram visual from the layout + data on open.
///
/// The node list is flat-with-level: each node carries a 0-based <c>level</c>, and a
/// node attaches under the most recent node one level shallower (level 0 hangs off the
/// document point). Reading <c>/slide[i]/smartart[k]</c> back returns the same layout
/// name and node texts via the existing <see cref="PptxSmartArt"/> reader.
/// </summary>
internal static class PptxDiagrams
{
    /// <summary>The drawingml/2006/diagram graphicData URI (the same one the reader matches on).</summary>
    private const string DiagramNs = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

    /// <summary>
    /// The supported layout names and the standard built-in layout uniqueId each maps to.
    /// These are the layout def ids PowerPoint ships, so the diagram regenerates with the
    /// intended arrangement on open. <c>orgChart</c> is the organisation-chart variant of
    /// the hierarchy family.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, string> LayoutIds =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["list"] = "urn:microsoft.com/office/officeart/2005/8/layout/vList2",
            ["process"] = "urn:microsoft.com/office/officeart/2005/8/layout/process1",
            ["hierarchy"] = "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1",
            ["orgChart"] = "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1",
            ["cycle"] = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2",
        };

    /// <summary>The standard colors part uniqueId (accent-1 transparency range).</summary>
    private const string ColorsId = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2";

    /// <summary>The standard quick-style part uniqueId (the "Simple Fill" quick style).</summary>
    private const string QuickStyleId = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1";

    private static readonly IReadOnlyList<string> LayoutNames = [.. LayoutIds.Keys];

    private static readonly IReadOnlyList<string> AddProps =
        ["layout", "nodes", "x", "y", "w", "h", "colorStyle", "name"];

    /// <summary>One node parsed from props.nodes: its text and 0-based depth.</summary>
    private readonly record struct NodeSpec(string Text, int Level);

    /// <summary>
    /// add smartart on a slide: builds the four diagram parts and the host graphicFrame,
    /// returning the canonical <c>/slide[i]/smartart[k]</c> path.
    /// </summary>
    public static string Add(SlidePart slidePart, int slideIndex, JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown smartart prop '{key}'.",
                    "smartart props: layout, nodes, x, y, w, h, colorStyle, name.",
                    candidates: AddProps);
            }
        }

        var layout = ParseLayout(props);
        var nodes = ParseNodes(props);

        var x = Length(props, "x", Units.CmToEmu(2));
        var y = Length(props, "y", Units.CmToEmu(3));
        var w = Length(props, "w", Units.CmToEmu(24));
        var h = Length(props, "h", Units.CmToEmu(12));

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
            ? J.ScalarText(nameNode)
            : Units.Inv($"Diagram {id}");

        var dataPart = slidePart.AddNewPart<DiagramDataPart>();
        var layoutPart = slidePart.AddNewPart<DiagramLayoutDefinitionPart>();
        var colorsPart = slidePart.AddNewPart<DiagramColorsPart>();
        var stylePart = slidePart.AddNewPart<DiagramStylePart>();

        dataPart.DataModelRoot = BuildDataModel(nodes);
        layoutPart.LayoutDefinition = BuildLayoutDefinition(LayoutIds[layout]);
        colorsPart.ColorsDefinition = BuildColorsDefinition();
        stylePart.StyleDefinition = BuildStyleDefinition();

        var relIds = new Dgm.RelationshipIds
        {
            DataPart = slidePart.GetIdOfPart(dataPart),
            LayoutPart = slidePart.GetIdOfPart(layoutPart),
            ColorPart = slidePart.GetIdOfPart(colorsPart),
            StylePart = slidePart.GetIdOfPart(stylePart),
        };

        tree.Append(new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = w, Cy = h }),
            new A.Graphic(new A.GraphicData(relIds) { Uri = DiagramNs })));

        // The new diagram is the slide's last graphicFrame; its 1-based smartart index
        // is therefore the new diagram count (List enumerates in paint order).
        var index = PptxSmartArt.List(slidePart).Count;
        return Units.Inv($"/slide[{slideIndex}]/smartart[{index}]");
    }

    // ----- prop parsing ---------------------------------------------------------

    private static string ParseLayout(JsonObject props)
    {
        if (!props.TryGetPropertyValue("layout", out var node) || node is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add smartart requires props.layout.",
                "Choose a layout: " + string.Join(", ", LayoutNames) + ", e.g. " +
                "{\"op\":\"add\",\"path\":\"/slide[1]\",\"type\":\"smartart\"," +
                "\"props\":{\"layout\":\"process\",\"nodes\":[{\"text\":\"A\",\"level\":0}]}}.",
                candidates: LayoutNames);
        }

        var raw = J.ScalarText(node).Trim();
        foreach (var name in LayoutNames)
        {
            if (string.Equals(name, raw, StringComparison.OrdinalIgnoreCase))
            {
                return name;
            }
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"SmartArt layout '{raw}' is not supported.",
            "Supported layouts: " + string.Join(", ", LayoutNames) +
            ". Pick the closest one and refine it in PowerPoint.",
            candidates: LayoutNames);
    }

    private static List<NodeSpec> ParseNodes(JsonObject props)
    {
        if (!props.TryGetPropertyValue("nodes", out var node) || node is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add smartart requires props.nodes (a non-empty array).",
                "Pass nodes as a flat list with 0-based levels, e.g. " +
                "\"nodes\":[{\"text\":\"Top\",\"level\":0},{\"text\":\"Sub\",\"level\":1}].");
        }

        if (array.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "props.nodes is empty.",
                "A diagram needs at least one node: \"nodes\":[{\"text\":\"Idea\",\"level\":0}].");
        }

        var nodes = new List<NodeSpec>();
        for (var i = 0; i < array.Count; i++)
        {
            if (array[i] is not JsonObject obj)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"props.nodes[{i}] is not an object.",
                    "Each node is {\"text\":\"...\",\"level\":N} with a 0-based level.");
            }

            var text = obj.TryGetPropertyValue("text", out var textNode) && textNode is not null
                ? J.ScalarText(textNode)
                : string.Empty;

            var level = 0;
            if (obj.TryGetPropertyValue("level", out var levelNode) && levelNode is not null)
            {
                if (levelNode is not JsonValue value ||
                    !(Units.TryNumber(value, out var number) ||
                      (value.TryGetValue<string>(out var rawLevel) &&
                       double.TryParse(rawLevel, NumberStyles.Integer, CultureInfo.InvariantCulture, out number))) ||
                    number != Math.Floor(number) || number < 0)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"props.nodes[{i}].level is not a non-negative integer: {levelNode.ToJsonString()}",
                        "Levels are 0-based depths; the top tier is level 0, its children level 1, and so on.");
                }

                level = (int)number;
            }

            // The first node must be a root (level 0); a deeper first node has no parent.
            if (i == 0 && level != 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"props.nodes[0].level is {level} but the first node must be level 0 (a root).",
                    "Start the list with a level-0 node; deeper levels attach to the nearest shallower node before them.");
            }

            nodes.Add(new NodeSpec(text, level));
        }

        return nodes;
    }

    private static long Length(JsonObject props, string key, long fallback) =>
        props.TryGetPropertyValue(key, out var node) ? Units.ParseLengthEmu(key, node) : fallback;

    // ----- data model -----------------------------------------------------------

    /// <summary>
    /// Builds the dgm:dataModel: a document point, one node point per spec (with its text
    /// body), and the parentOf connections wiring each node under the most recent node one
    /// level shallower (level-0 nodes hang off the document point). This is exactly the
    /// shape <see cref="PptxSmartArt.Nodes"/> walks when reading the diagram back.
    /// </summary>
    private static Dgm.DataModelRoot BuildDataModel(IReadOnlyList<NodeSpec> nodes)
    {
        const string docId = "{00000000-0000-0000-0000-000000000000}";
        var pointList = new Dgm.PointList();
        var connectionList = new Dgm.ConnectionList();

        var docPoint = new Dgm.Point { ModelId = docId, Type = Dgm.PointValues.Document };
        docPoint.Append(new Dgm.PropertySet());
        docPoint.Append(new Dgm.ShapeProperties());
        pointList.Append(docPoint);

        // The most recent node id at each depth, so a node finds its parent by level.
        var parentByLevel = new Dictionary<int, string>(); // level -> node modelId
        uint connectionOrder = 0;

        for (var i = 0; i < nodes.Count; i++)
        {
            var spec = nodes[i];
            var nodeId = Units.Inv($"{{10000000-0000-0000-0000-{(i + 1):D12}}}");

            var point = new Dgm.Point { ModelId = nodeId, Type = Dgm.PointValues.Node };
            point.Append(new Dgm.PropertySet());
            point.Append(new Dgm.ShapeProperties());
            point.Append(BuildTextBody(spec.Text));
            pointList.Append(point);

            // Parent: the most recent node at level-1; level-0 attaches to the document.
            var parentId = spec.Level == 0
                ? docId
                : parentByLevel.TryGetValue(spec.Level - 1, out var pid) ? pid : docId;

            connectionList.Append(new Dgm.Connection
            {
                ModelId = Units.Inv($"{{20000000-0000-0000-0000-{(i + 1):D12}}}"),
                Type = Dgm.ConnectionValues.ParentOf,
                SourceId = parentId,
                DestinationId = nodeId,
                SourcePosition = connectionOrder++,
                DestinationPosition = 0U,
            });

            // This node becomes the standing parent at its own level, and invalidates
            // any deeper standing parents (a new branch starts here).
            parentByLevel[spec.Level] = nodeId;
            foreach (var deeper in parentByLevel.Keys.Where(l => l > spec.Level).ToList())
            {
                parentByLevel.Remove(deeper);
            }
        }

        return new Dgm.DataModelRoot(pointList, connectionList, new Dgm.Background(), new Dgm.Whole());
    }

    private static Dgm.TextBody BuildTextBody(string text) => new(
        new A.BodyProperties(),
        new A.ListStyle(),
        new A.Paragraph(new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(text))));

    // ----- layout / colors / style parts ----------------------------------------

    /// <summary>
    /// A minimal valid layout definition that references a built-in layout by its
    /// uniqueId (PowerPoint regenerates the arrangement from it), with the single
    /// layoutNode the schema requires.
    /// </summary>
    private static Dgm.LayoutDefinition BuildLayoutDefinition(string uniqueId) => new(
        new Dgm.Title { Val = string.Empty },
        new Dgm.Description { Val = string.Empty },
        new Dgm.LayoutNode(new Dgm.Shape()) { Name = "diagram" })
    {
        UniqueId = uniqueId,
    };

    /// <summary>A minimal valid colors definition referencing a built-in color style.</summary>
    private static Dgm.ColorsDefinition BuildColorsDefinition() => new()
    {
        UniqueId = ColorsId,
    };

    /// <summary>
    /// A minimal valid quick-style definition: the single styleLbl the schema requires,
    /// carrying the four style references (line/fill/effect/font) that complete a
    /// dgm:style element.
    /// </summary>
    private static Dgm.StyleDefinition BuildStyleDefinition() => new(
        new Dgm.StyleDefinitionTitle { Val = string.Empty },
        new Dgm.StyleLabel(BuildStyle()) { Name = "node0" })
    {
        UniqueId = QuickStyleId,
    };

    private static Dgm.Style BuildStyle() => new(
        new A.LineReference { Index = 0U },
        new A.FillReference { Index = 0U },
        new A.EffectReference { Index = 0U },
        new A.FontReference(new A.RgbColorModelHex { Val = "000000" }) { Index = A.FontCollectionIndexValues.Minor });
}
