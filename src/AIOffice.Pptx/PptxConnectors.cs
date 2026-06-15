using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Shape connectors: a p:cxnSp whose cxnSpPr carries a:stCxn/a:endCxn referencing the
/// two anchor shapes' ids, with the preset geometry the kind selects
/// (line → straightConnector1, elbow → bentConnector3, curved → curvedConnector3). The
/// connector's bounding box spans the two shapes so it draws between them; arrow heads,
/// stroke color and width are optional. The endpoints survive a round-trip and are read
/// back by <c>get</c>.
/// </summary>
internal static class PptxConnectors
{
    private static readonly IReadOnlyList<string> AddProps =
        ["kind", "from", "to", "startArrow", "endArrow", "color", "width", "name"];

    private static readonly IReadOnlyList<string> Kinds = ["straight", "elbow", "curved"];

    private static readonly IReadOnlyList<string> Arrows = ["none", "arrow", "triangle"];

    /// <summary>Default connector stroke width (1pt).</summary>
    private const int DefaultWidthEmu = 12_700;

    private static readonly IReadOnlyDictionary<string, A.ShapeTypeValues> KindGeometry =
        new Dictionary<string, A.ShapeTypeValues>(StringComparer.Ordinal)
        {
            ["straight"] = A.ShapeTypeValues.StraightConnector1,
            ["elbow"] = A.ShapeTypeValues.BentConnector3,
            ["curved"] = A.ShapeTypeValues.CurvedConnector3,
        };

    /// <summary>
    /// add connector on a slide: wires a p:cxnSp between the from/to shapes and returns its
    /// stable shape id. An unknown from/to shape is <c>invalid_path</c> listing the slide's
    /// shape ids.
    /// </summary>
    public static uint Add(SlidePart slidePart, JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown connector prop '{key}'.",
                    "connector props: kind, from, to, startArrow, endArrow, color, width, name.",
                    candidates: AddProps);
            }
        }

        var kind = ParseKind(props);
        var shapes = PptxDoc.Shapes(slidePart);
        var from = ResolveAnchor(props, "from", shapes);
        var to = ResolveAnchor(props, "to", shapes);

        if (from.Id == to.Id)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A connector's from and to must be different shapes.",
                "Point the connector at two distinct shapes (by @id or name).");
        }

        var startArrow = ParseArrow(props, "startArrow");
        var endArrow = ParseArrow(props, "endArrow");
        var color = props.TryGetPropertyValue("color", out var colorNode)
            ? Units.ParseColorHex("color", colorNode)
            : "000000";
        var width = props.TryGetPropertyValue("width", out var widthNode)
            ? Units.ParseLengthEmu("width", widthNode)
            : DefaultWidthEmu;

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
            ? J.ScalarText(nameNode)
            : Units.Inv($"Connector {id}");

        var (x, y, w, h) = BoundingBox(from, to);

        var outline = new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = color })) { Width = (int)width };
        if (startArrow is { } head)
        {
            outline.Append(new A.HeadEnd { Type = head });
        }

        if (endArrow is { } tail)
        {
            outline.Append(new A.TailEnd { Type = tail });
        }

        var connectionShape = new P.ConnectionShape(
            new P.NonVisualConnectionShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualConnectorShapeDrawingProperties(
                    new A.StartConnection { Id = from.Id, Index = 0U },
                    new A.EndConnection { Id = to.Id, Index = 0U }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = w, Cy = h }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = KindGeometry[kind] },
                outline),
            new P.ShapeStyle(
                new A.LineReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 2U },
                new A.FillReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
                new A.EffectReference(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Index = 0U },
                new A.FontReference(new A.SchemeColor { Val = A.SchemeColorValues.Text1 }) { Index = A.FontCollectionIndexValues.Minor }));

        tree.Append(connectionShape);
        return id;
    }

    private static string ParseKind(JsonObject props)
    {
        if (!props.TryGetPropertyValue("kind", out var node) || node is null)
        {
            return "straight"; // a straight line is the default connector
        }

        var raw = J.ScalarText(node).Trim();
        foreach (var kind in Kinds)
        {
            if (string.Equals(kind, raw, StringComparison.OrdinalIgnoreCase))
            {
                return kind;
            }
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"Connector kind '{raw}' is not supported.",
            "Supported kinds: straight (a direct line), elbow (right-angle), curved.",
            candidates: Kinds);
    }

    private static A.LineEndValues? ParseArrow(JsonObject props, string key)
    {
        if (!props.TryGetPropertyValue(key, out var node) || node is null)
        {
            return null;
        }

        var raw = J.ScalarText(node).Trim().ToLowerInvariant();
        return raw switch
        {
            "none" => null,
            "arrow" => A.LineEndValues.Stealth,
            "triangle" => A.LineEndValues.Triangle,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid {key} value: {node.ToJsonString()}",
                "Use none, arrow or triangle.",
                candidates: Arrows),
        };
    }

    /// <summary>Resolves a from/to anchor by @id (a leading @, e.g. "@5") or by shape name.</summary>
    private static ShapeView ResolveAnchor(JsonObject props, string key, IReadOnlyList<ShapeView> shapes)
    {
        if (!props.TryGetPropertyValue(key, out var node) || node is null || J.ScalarText(node).Trim().Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add connector requires props.{key}.",
                "Name both endpoints by @id or name, e.g. " +
                "{\"op\":\"add\",\"path\":\"/slide[1]\",\"type\":\"connector\"," +
                "\"props\":{\"from\":\"@2\",\"to\":\"@3\"}}.");
        }

        var raw = J.ScalarText(node).Trim();

        // "@N" addresses a shape by its stable id; otherwise the token is a shape name.
        if (raw.StartsWith('@'))
        {
            var rest = raw[1..];
            if (uint.TryParse(rest, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var anchorId) &&
                shapes.FirstOrDefault(s => s.Id == anchorId) is { } byId)
            {
                return byId;
            }
        }
        else if (shapes.FirstOrDefault(s => string.Equals(s.Name, raw, StringComparison.Ordinal)) is { } byName)
        {
            return byName;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"Connector {key} '{raw}' does not match any shape on the slide.",
            "Reference an endpoint by @id (e.g. \"@5\") or by its exact name; " +
            "run 'aioffice get <file> --path /slide[i]' to list the slide's shapes.",
            candidates: [.. shapes.Take(10).Select(s => Units.Inv($"@{s.Id}"))]);
    }

    /// <summary>The connector box: the union of the two anchors' geometries (a zero-size box when neither has one).</summary>
    private static (long X, long Y, long W, long H) BoundingBox(ShapeView from, ShapeView to)
    {
        var a = PptxDoc.Geometry(from.Element);
        var b = PptxDoc.Geometry(to.Element);
        if (a is null && b is null)
        {
            return (0L, 0L, 0L, 0L);
        }

        // Span from the min top-left to the max bottom-right across both anchors.
        var boxes = new[] { a, b }.Where(g => g is not null).Select(g => g!.Value).ToList();
        var minX = boxes.Min(g => g.X);
        var minY = boxes.Min(g => g.Y);
        var maxX = boxes.Max(g => g.X + g.Cx);
        var maxY = boxes.Max(g => g.Y + g.Cy);
        return (minX, minY, Math.Max(maxX - minX, 0L), Math.Max(maxY - minY, 0L));
    }

    /// <summary>The two endpoint shape ids a connector references, when it carries cxn references.</summary>
    public static (uint? StartId, uint? EndId) Endpoints(P.ConnectionShape connector)
    {
        var drawing = connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties;
        return (
            drawing?.GetFirstChild<A.StartConnection>()?.Id?.Value,
            drawing?.GetFirstChild<A.EndConnection>()?.Id?.Value);
    }
}
