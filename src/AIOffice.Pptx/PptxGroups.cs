using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Shape grouping (group / ungroup). <see cref="Group"/> wraps the named shapes in a
/// p:grpSp whose child coordinate space (a:chOff/a:chExt) is the identity of its own
/// box (a:off/a:ext), so the wrapped children keep their existing absolute coordinates.
/// <see cref="Ungroup"/> dissolves a group, promoting its children back onto the slide
/// at the same absolute positions. A grouped shape is addressed
/// <c>/slide[i]/group[@id=N]</c>, its children <c>/slide[i]/group[@id=N]/shape[...]</c>.
/// </summary>
internal static class PptxGroups
{
    /// <summary>
    /// group on a slide: wraps the props.shapes (each "@id" or a shape name) in a new
    /// p:grpSp, returning the canonical /slide[i]/group[@id=N] path. The group is placed
    /// where the topmost wrapped shape was (paint order is preserved among them).
    /// </summary>
    public static string Group(SlidePart slidePart, int slideIndex, JsonObject? props)
    {
        if (props is null)
        {
            throw MissingShapes();
        }

        foreach (var (key, _) in props)
        {
            if (key is not ("shapes" or "name"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown group prop '{key}'.",
                    "group props: shapes (the list to wrap) and name.",
                    candidates: ["shapes", "name"]);
            }
        }

        if (!props.TryGetPropertyValue("shapes", out var shapesNode) || shapesNode is not JsonArray array || array.Count == 0)
        {
            throw MissingShapes();
        }

        if (array.Count < 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A group needs at least two shapes.",
                "List two or more shapes to wrap, e.g. \"shapes\":[\"@2\",\"@3\"].");
        }

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var shapes = PptxDoc.Shapes(tree);
        var members = ResolveMembers(array, shapes);

        // The group's box is the union of the members' geometries; children keep their
        // absolute coordinates because the child space (chOff/chExt) equals the box.
        var box = UnionBox(members);

        var id = PptxDoc.NextShapeId(tree);
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
            ? J.ScalarText(nameNode)
            : Units.Inv($"Group {id}");

        var groupShape = new P.GroupShape(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(
                new A.TransformGroup(
                    new A.Offset { X = box.X, Y = box.Y },
                    new A.Extents { Cx = box.W, Cy = box.H },
                    new A.ChildOffset { X = box.X, Y = box.Y },
                    new A.ChildExtents { Cx = box.W, Cy = box.H })));

        // Insert the group where the first (topmost-painted) member sat, then move each
        // member into the group in their original paint order.
        var anchor = members[0].Element;
        tree.InsertBefore(groupShape, anchor);
        foreach (var member in members)
        {
            member.Element.Remove();
            groupShape.Append(member.Element);
        }

        return Units.Inv($"/slide[{slideIndex}]/group[@id={id}]");
    }

    /// <summary>
    /// ungroup a /slide[i]/group[@id=N]: dissolves the group, splicing its children back
    /// into the slide's shape tree where the group sat, each keeping its absolute
    /// coordinates. Returns the canonical group path the op targeted.
    /// </summary>
    public static string Ungroup(SlidePart slidePart, PptxAddress address)
    {
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var group = ResolveGroup(tree, address);

        // Each child's offset is already absolute (the group's child space is the identity
        // of its box), so promotion is a straight splice — no coordinate rebasing needed.
        // If the group used a non-identity child space (a foreign deck), rebase the child
        // offsets so the absolute on-slide position is preserved.
        var transform = group.GroupShapeProperties?.TransformGroup;
        var rebase = ChildToParentRebase(transform);

        OpenXmlElement cursor = group;
        foreach (var child in group.ChildElements
                     .Where(c => c is P.Shape or P.Picture or P.GraphicFrame or P.GroupShape or P.ConnectionShape)
                     .ToList())
        {
            child.Remove();
            if (rebase is { } r)
            {
                RebaseChild(child, r);
            }

            tree.InsertAfter(child, cursor);
            cursor = child;
        }

        group.Remove();
        return address.CanonicalGroupPath;
    }

    /// <summary>Resolves a group element by the address's group id, or throws invalid_path with candidates.</summary>
    public static P.GroupShape ResolveGroup(P.ShapeTree tree, PptxAddress address)
    {
        var groups = tree.Elements<P.GroupShape>().ToList();
        var match = groups.FirstOrDefault(g =>
            g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value == address.GroupId);
        if (match is not null)
        {
            return match;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No group with id {address.GroupId} on slide {address.SlideIndex}; it has {groups.Count} group(s).",
            "Run 'aioffice get <file> --path /slide[i]' to list the slide's shapes (groups included).",
            candidates: [.. groups
                .Select(g => g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value)
                .Where(gid => gid is not null)
                .Take(10)
                .Select(gid => Units.Inv($"/slide[{address.SlideIndex}]/group[@id={gid}]"))]);
    }

    /// <summary>The top-level shapes inside a group, with stable ids and 1-based ordinals.</summary>
    public static List<ShapeView> Children(P.GroupShape group)
    {
        var views = new List<ShapeView>();
        var ordinal = 0;
        foreach (var element in group.ChildElements)
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
            var nv = PptxDoc.NonVisualProps(composite);
            views.Add(new ShapeView(composite, nv?.Id?.Value ?? 0, ordinal, kind, nv?.Name?.Value ?? string.Empty));
        }

        return views;
    }

    /// <summary>Resolves a child shape inside a group by the address's shape segment.</summary>
    public static ShapeView ResolveChild(P.GroupShape group, PptxAddress address)
    {
        var children = Children(group);
        ShapeView? match = address.ShapeId is { } id
            ? children.FirstOrDefault(s => s.Id == id)
            : children.FirstOrDefault(s => s.Ordinal == address.ShapeOrdinal);
        if (match is not null)
        {
            return match;
        }

        var what = address.ShapeId is { } sid ? $"id {sid}" : $"index {address.ShapeOrdinal}";
        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No shape with {what} in group {address.GroupId} on slide {address.SlideIndex}; it has {children.Count} child(ren).",
            "Run 'aioffice get <file> --path " + address.CanonicalGroupPath + "' to list the group's children.",
            candidates: [.. children.Take(10).Select(s =>
                Units.Inv($"{address.CanonicalGroupPath}/shape[@id={s.Id}]"))]);
    }

    // ----- helpers --------------------------------------------------------------

    private static List<ShapeView> ResolveMembers(JsonArray array, IReadOnlyList<ShapeView> shapes)
    {
        var members = new List<ShapeView>();
        var seen = new HashSet<uint>();
        foreach (var node in array)
        {
            var raw = (node is null ? string.Empty : J.ScalarText(node)).Trim();
            if (raw.Length == 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "A group member reference is empty.",
                    "Each entry is \"@id\" or a shape name, e.g. \"shapes\":[\"@2\",\"Title\"].");
            }

            var view = ResolveOne(raw, shapes);
            if (seen.Add(view.Id))
            {
                members.Add(view);
            }
        }

        if (members.Count < 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A group needs at least two distinct shapes.",
                "List two or more different shapes; duplicates are ignored.");
        }

        // Wrap them in slide paint order so the group's internal z-order is stable.
        return [.. members.OrderBy(m => shapes.ToList().FindIndex(s => ReferenceEquals(s.Element, m.Element)))];
    }

    private static ShapeView ResolveOne(string raw, IReadOnlyList<ShapeView> shapes)
    {
        if (raw.StartsWith('@'))
        {
            var rest = raw[1..];
            if (uint.TryParse(rest, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var id) &&
                shapes.FirstOrDefault(s => s.Id == id) is { } byId)
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
            $"Group member '{raw}' does not match any shape on the slide.",
            "Reference each shape by @id (e.g. \"@5\") or its exact name; " +
            "run 'aioffice get <file> --path /slide[i]' to list the slide's shapes.",
            candidates: [.. shapes.Take(10).Select(s => Units.Inv($"@{s.Id}"))]);
    }

    private static (long X, long Y, long W, long H) UnionBox(IReadOnlyList<ShapeView> members)
    {
        var boxes = members.Select(m => PptxDoc.Geometry(m.Element)).Where(g => g is not null).Select(g => g!.Value).ToList();
        if (boxes.Count == 0)
        {
            return (0L, 0L, Units.CmToEmu(10), Units.CmToEmu(3));
        }

        var minX = boxes.Min(b => b.X);
        var minY = boxes.Min(b => b.Y);
        var maxX = boxes.Max(b => b.X + b.Cx);
        var maxY = boxes.Max(b => b.Y + b.Cy);
        return (minX, minY, Math.Max(maxX - minX, 1L), Math.Max(maxY - minY, 1L));
    }

    /// <summary>
    /// The affine that maps a child's coordinate (in the group's child space) to the
    /// slide. With an identity child space (what <see cref="Group"/> writes) this is a
    /// no-op and returns null; a foreign group with a scaled/translated child space
    /// returns the scale + translation so promotion preserves the on-slide position.
    /// </summary>
    private static (double Sx, double Sy, long Tx, long Ty)? ChildToParentRebase(A.TransformGroup? transform)
    {
        var off = transform?.Offset;
        var ext = transform?.Extents;
        var chOff = transform?.ChildOffset;
        var chExt = transform?.ChildExtents;
        if (off?.X is null || off.Y is null || ext?.Cx is null || ext.Cy is null ||
            chOff?.X is null || chOff.Y is null || chExt?.Cx is null || chExt.Cy is null)
        {
            return null;
        }

        var sx = chExt.Cx!.Value == 0 ? 1.0 : (double)ext.Cx!.Value / chExt.Cx.Value;
        var sy = chExt.Cy!.Value == 0 ? 1.0 : (double)ext.Cy!.Value / chExt.Cy.Value;
        var tx = off.X!.Value;
        var ty = off.Y!.Value;

        // Identity child space (Group's own output): nothing to rebase.
        if (sx == 1.0 && sy == 1.0 && chOff.X!.Value == tx && chOff.Y!.Value == ty)
        {
            return null;
        }

        // slideX = off.x + (childX - chOff.x) * sx  ->  fold chOff into the translation.
        return (sx, sy,
            tx - (long)Math.Round(chOff.X!.Value * sx),
            ty - (long)Math.Round(chOff.Y!.Value * sy));
    }

    private static void RebaseChild(OpenXmlElement child, (double Sx, double Sy, long Tx, long Ty) r)
    {
        var transform = child switch
        {
            P.Shape s => s.ShapeProperties?.Transform2D,
            P.Picture p => p.ShapeProperties?.Transform2D,
            P.ConnectionShape c => c.ShapeProperties?.Transform2D,
            _ => null,
        };

        if (transform?.Offset is { } offset && transform.Extents is { } extents)
        {
            offset.X = r.Tx + (long)Math.Round((offset.X?.Value ?? 0) * r.Sx);
            offset.Y = r.Ty + (long)Math.Round((offset.Y?.Value ?? 0) * r.Sy);
            extents.Cx = (long)Math.Round((extents.Cx?.Value ?? 0) * r.Sx);
            extents.Cy = (long)Math.Round((extents.Cy?.Value ?? 0) * r.Sy);
        }
        else if (child is P.GraphicFrame { Transform: { } frameTransform })
        {
            if (frameTransform.Offset is { } fo)
            {
                fo.X = r.Tx + (long)Math.Round((fo.X?.Value ?? 0) * r.Sx);
                fo.Y = r.Ty + (long)Math.Round((fo.Y?.Value ?? 0) * r.Sy);
            }
        }
    }

    private static AiofficeException MissingShapes() => new(
        ErrorCodes.InvalidArgs,
        "add group requires props.shapes (a list of at least two shapes).",
        "List the shapes to wrap by @id or name, e.g. " +
        "{\"op\":\"add\",\"path\":\"/slide[1]\",\"type\":\"group\",\"props\":{\"shapes\":[\"@2\",\"@3\"]}}.");
}
