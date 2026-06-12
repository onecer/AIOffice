using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    public Envelope Edit(CommandContext ctx, IReadOnlyList<EditOp> ops)
    {
        var file = RequireFile(ctx, mustExist: true);
        var dryRun = BoolArg(ctx.Args, "dryRun") || BoolArg(ctx.Args, "dry-run");
        var expectRev = StringArg(ctx.Args, "expectRev") ?? StringArg(ctx.Args, "expect-rev");

        var originalBytes = File.ReadAllBytes(file);
        var currentRev = Rev.OfBytes(originalBytes);
        if (expectRev is not null && !expectRev.Equals(currentRev, StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.StaleAddress,
                $"The file changed since you last read it: rev is {currentRev}, you expected {expectRev}.",
                "Re-run 'aioffice read' or 'aioffice query' to get fresh paths and the current rev, then retry.");
        }

        // Atomic: every op is applied to an in-memory copy; the file is written only when all succeed.
        var ms = new MemoryStream();
        ms.Write(originalBytes);
        ms.Position = 0;

        var summaries = new List<object>(ops.Count);
        using (var doc = OpenPackage(ms, file, editable: true))
        {
            for (var i = 0; i < ops.Count; i++)
            {
                try
                {
                    summaries.Add(ApplyOp(doc, file, ops[i]));
                }
                catch (AiofficeException ex)
                {
                    throw new AiofficeException(
                        ex.Code,
                        $"ops[{i}] ({ops[i].Op} {ops[i].Path}): {ex.Message}",
                        ex.Suggestion,
                        ex.Candidates,
                        ex);
                }
            }
        }

        var newBytes = ms.ToArray();

        if (dryRun)
        {
            return Envelope.Ok(
                new { applied = summaries.Count, dryRun = true, ops = summaries },
                MetaFor(file, currentRev));
        }

        var snapshot = _snapshots.Save(file); // pre-image, so the edit is undoable
        File.WriteAllBytes(file, newBytes);

        return Envelope.Ok(
            new { applied = summaries.Count, snapshot = snapshot.Number, ops = summaries },
            MetaFor(file, Rev.OfBytes(newBytes)));
    }

    // ------------------------------------------------------------------- ops

    private static object ApplyOp(WordprocessingDocument doc, string file, EditOp op) => op.Op switch
    {
        "set" => ApplySet(doc, op),
        "add" => ApplyAdd(doc, file, op),
        "remove" => ApplyRemove(doc, op),
        _ => ApplyMove(doc, op),
    };

    private static object ApplySet(WordprocessingDocument doc, EditOp op)
    {
        var node = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        var props = RequireProps(op);

        foreach (var (name, value) in OrderedProps(props))
        {
            switch (node.Element)
            {
                case Paragraph p:
                    if (name == "style")
                    {
                        EnsureStyleDefined(doc, value);
                    }

                    WordFormatting.SetParagraphProp(p, name, value);
                    break;

                case Run r:
                    WordFormatting.SetRunProp(r, name, value);
                    break;

                case TableCell cell when name == "text":
                    cell.RemoveAllChildren<Paragraph>();
                    cell.AppendChild(WordFactory.Paragraph(value));
                    break;

                default:
                    throw new AiofficeException(
                        ErrorCodes.UnsupportedFeature,
                        $"set is not supported on '{node.Type}' (property '{name}').",
                        node.Type is "tc"
                            ? "A table cell only supports text, or address a paragraph inside it: " + node.CanonicalPath + "/p[1]."
                            : "Set properties on p or run elements; address one with query first.");
            }
        }

        return new { op = "set", path = node.CanonicalPath, type = node.Type };
    }

    private static object ApplyAdd(WordprocessingDocument doc, string file, EditOp op)
    {
        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        var type = op.Type ?? "p";
        var position = op.Position;
        if (position is not (null or "before" or "after" or "inside"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add position '{position}' is not valid.",
                "Use position before, after or inside (inside only for /body, table or tc targets).",
                candidates: ["before", "after", "inside"]);
        }

        OpenXmlElement created = type switch
        {
            "p" => BuildParagraph(doc, op.Props),
            "tr" => BuildRow(op.Props),
            "table" => WordFactory.Table(
                rows: PropInt(op.Props, "rows") ?? 2,
                columns: PropInt(op.Props, "columns") ?? PropInt(op.Props, "cols") ?? 2),
            _ => throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"add --type {type} is not supported for docx.",
                "Add p (use props.style=Heading1 for headings), tr (props.cells=[…]) or table (props.rows/columns). " +
                "For runs, set text on the paragraph instead.",
                candidates: ["p", "tr", "table"]),
        };

        // Default placement: containers receive children, blocks get siblings after them.
        var isContainerAnchor = anchor.Type is "body" or "tc" || (anchor.Type == "table" && created is TableRow);
        var pos = position ?? (isContainerAnchor ? "inside" : "after");

        var valid = (created, anchor.Type, pos) switch
        {
            (Paragraph or Table, "body" or "tc", "inside") => true,
            (Paragraph or Table, "p" or "table", "before" or "after") => true,
            (TableRow, "table", "inside") => true,
            (TableRow, "tr", "before" or "after") => true,
            _ => false,
        };

        if (!valid)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cannot add a '{type}' {pos} {anchor.CanonicalPath} ({anchor.Type}).",
                "Add p/table inside /body or a tc, or before/after an existing p/table; " +
                "add tr inside a table, or before/after an existing tr.");
        }

        var canonical = Insert(GetBody(doc, file), anchor, created, pos);
        return new { op = "add", type, anchor = anchor.CanonicalPath, path = canonical };
    }

    /// <summary>Inserts at the validated position; appending to body keeps sectPr last.</summary>
    private static string Insert(Body body, ResolvedNode anchor, OpenXmlElement created, string pos)
    {
        switch (pos)
        {
            case "before":
                anchor.Element.InsertBeforeSelf(created);
                break;

            case "after":
                anchor.Element.InsertAfterSelf(created);
                break;

            default: // "inside"
                if (anchor.Element is Body b && b.Elements<SectionProperties>().FirstOrDefault() is { } sectPr)
                {
                    sectPr.InsertBeforeSelf(created);
                }
                else
                {
                    anchor.Element.AppendChild(created);
                }

                break;
        }

        // Report where the new node landed as a canonical path.
        var match = WordAddress.EnumerateBody(body).FirstOrDefault(n => ReferenceEquals(n.Element, created));
        return match?.CanonicalPath ?? anchor.CanonicalPath;
    }

    private static object ApplyRemove(WordprocessingDocument doc, EditOp op)
    {
        var node = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (node.Element is not (Paragraph or Table or TableRow))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"remove is not supported on '{node.Type}'.",
                "Remove p, table or tr elements. To clear a run or cell, set its text to \"\" instead.");
        }

        if (node.Element is Paragraph && node.Element.Parent is TableCell cell &&
            cell.Elements<Paragraph>().Count() == 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"{node.CanonicalPath} is the only paragraph in its table cell; a cell must keep one.",
                "Set its text to \"\" instead: {\"op\":\"set\",\"path\":\"" + node.CanonicalPath + "\",\"props\":{\"text\":\"\"}}.");
        }

        node.Element.Remove();
        return new { op = "remove", path = node.CanonicalPath, type = node.Type };
    }

    private static object ApplyMove(WordprocessingDocument doc, EditOp op)
    {
        var source = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (source.Element is not (Paragraph or Table or TableRow))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"move is not supported on '{source.Type}'.",
                "Move p, table or tr elements.");
        }

        var (before, targetPath) = ParseMovePosition(op.Position);
        var target = WordAddress.Resolve(doc, DocPath.Parse(targetPath));
        if (ReferenceEquals(source.Element, target.Element))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "move source and target are the same element.",
                "Pick a different target path in position.");
        }

        source.Element.Remove();
        if (before)
        {
            target.Element.InsertBeforeSelf(source.Element);
        }
        else
        {
            target.Element.InsertAfterSelf(source.Element);
        }

        return new { op = "move", path = source.CanonicalPath, target = target.CanonicalPath, placement = before ? "before" : "after" };
    }

    private static (bool Before, string TargetPath) ParseMovePosition(string? position)
    {
        var parts = position?.Split([' ', ':'], 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts is [var direction, var target] && direction is "before" or "after" && target.StartsWith('/'))
        {
            return (direction == "before", target);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"move needs position '<before|after> <path>', got '{position}'.",
            "Example: {\"op\":\"move\",\"path\":\"/body/p[3]\",\"position\":\"before /body/p[1]\"}.");
    }

    // --------------------------------------------------------------- helpers

    private static Paragraph BuildParagraph(WordprocessingDocument doc, JsonObject? props)
    {
        var paragraph = new Paragraph();
        foreach (var (name, value) in OrderedProps(props ?? []))
        {
            if (name == "style")
            {
                EnsureStyleDefined(doc, value);
            }

            WordFormatting.SetParagraphProp(paragraph, name, value);
        }

        return paragraph;
    }

    private static TableRow BuildRow(JsonObject? props)
    {
        var cells = props?["cells"] as JsonArray;
        if (cells is null || cells.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type tr needs props.cells with one text per cell.",
                "Example: {\"op\":\"add\",\"path\":\"/body/table[1]\",\"type\":\"tr\",\"props\":{\"cells\":[\"a\",\"b\"]}}.");
        }

        return WordFactory.Row([.. cells.Select(c => NodeToString(c))]);
    }

    private static JsonObject RequireProps(EditOp op) =>
        op.Props is { Count: > 0 }
            ? op.Props
            : throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "set needs at least one property in props.",
                "Example: {\"op\":\"set\",\"path\":\"/body/p[1]\",\"props\":{\"text\":\"Hello\",\"bold\":true}}.");

    /// <summary>Props in application order: text first (creates runs), then style, then formatting.</summary>
    private static IEnumerable<(string Name, string Value)> OrderedProps(JsonObject props) =>
        props
            .OrderBy(kv => kv.Key switch { "text" => 0, "style" => 1, _ => 2 })
            .Select(kv => (kv.Key, NodeToString(kv.Value)));

    private static int? PropInt(JsonObject? props, string name) =>
        props?[name] is { } node && int.TryParse(NodeToString(node), out var n) ? n : null;

    internal static string NodeToString(JsonNode? node)
    {
        if (node is null)
        {
            return string.Empty;
        }

        if (node is JsonValue v)
        {
            if (v.TryGetValue<string>(out var s))
            {
                return s;
            }

            if (v.TryGetValue<bool>(out var b))
            {
                return b ? "true" : "false";
            }

            if (v.TryGetValue<double>(out var d))
            {
                return d.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        return node.ToJsonString();
    }

    /// <summary>Adds a built-in Heading1..6 definition when an edit references one that is missing.</summary>
    private static void EnsureStyleDefined(WordprocessingDocument doc, string styleId)
    {
        if (HeadingLevel(styleId) is not { } level || level > 6 || doc.MainDocumentPart is not { } main)
        {
            return; // custom styles pass through untouched; Word renders them unstyled if undefined
        }

        var stylesPart = main.StyleDefinitionsPart;
        if (stylesPart is null)
        {
            WordFactory.AddDefaultStylesPart(main);
            stylesPart = main.StyleDefinitionsPart!;
        }

        var styles = stylesPart.Styles ??= new Styles();
        var canonicalId = "Heading" + level;
        if (!styles.Elements<Style>().Any(s => s.StyleId?.Value == canonicalId))
        {
            var halfPoints = level switch { 1 => "32", 2 => "28", 3 => "26", _ => "24" };
            styles.AppendChild(WordFactory.HeadingStyle(level, halfPoints));
        }
    }
}
