using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Applies one edit op to an open presentation. Callers batch ops over an
/// in-memory stream and only write the file when every op succeeded (atomic).
/// </summary>
internal static class PptxEditor
{
    private static readonly IReadOnlyList<string> AddTypes = ["slide", "shape", "textbox"];

    private static readonly IReadOnlyList<string> ShapePropKeys =
        ["text", "x", "y", "w", "h", "fill", "fontSize", "bold", "color", "align", "name", "title"];

    /// <summary>Applies the op and returns the canonical path of the affected node.</summary>
    public static string Apply(PresentationPart presentation, EditOp op)
    {
        if (PptxAddress.Parse(op.Path).IsMaster)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Editing masters/layouts is not supported yet (planned M2): {op.Path}",
                "Edit slides instead, or copy a layout-derived slide via 'add slide' with props {\"layout\": N}.");
        }

        return op.Op switch
        {
            "add" => ApplyAdd(presentation, op),
            "set" => ApplySet(presentation, op),
            "remove" => ApplyRemove(presentation, op),
            "move" => ApplyMove(presentation, op),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown op '{op.Op}'.",
                "Use set, add, remove or move.",
                candidates: EditOp.Kinds),
        };
    }

    private static string ApplyAdd(PresentationPart presentation, EditOp op)
    {
        var address = PptxAddress.Parse(op.Path);
        var type = op.Type?.Trim().ToLowerInvariant();
        switch (type)
        {
            case "slide":
                return AddSlide(presentation, address, op.Position, op.Props);

            case "shape" or "textbox":
            {
                if (address.HasShape)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"add shape targets a slide, not '{op.Path}'.",
                        "Use the slide path as the target, e.g. {\"op\":\"add\",\"path\":\"/slide[2]\",\"type\":\"shape\"}.");
                }

                var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, op.Path);
                var id = AddTextBox(slidePart, op.Props);
                return Units.Inv($"/slide[{address.SlideIndex}]/shape[@id={id}]");
            }

            default:
                throw new AiofficeException(
                    op.Type is null ? ErrorCodes.InvalidArgs : ErrorCodes.UnsupportedFeature,
                    op.Type is null ? "add requires a type." : $"Cannot add '{op.Type}' yet.",
                    "Addable types today: slide, shape (textbox). For pictures/tables/charts, build the deck and add them in PowerPoint for now.",
                    candidates: AddTypes);
        }
    }

    private static string AddSlide(PresentationPart presentation, PptxAddress address, string? position, JsonObject? props)
    {
        if (address.HasShape)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add slide targets a slide position, not '{address.Raw}'.",
                "Use e.g. {\"op\":\"add\",\"path\":\"/slide[3]\",\"type\":\"slide\"} — the new slide becomes slide 3.");
        }

        var slideIdList = presentation.Presentation?.SlideIdList
            ?? throw CorruptPresentation("p:sldIdLst is missing");
        var count = slideIdList.Elements<P.SlideId>().Count();

        var target = position?.Trim().ToLowerInvariant() switch
        {
            null or "" or "at" or "before" => address.SlideIndex,
            "after" => address.SlideIndex + 1,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown position '{position}' for add slide.",
                "Use \"at\"/\"before\" (new slide takes the path's index) or \"after\".",
                candidates: ["at", "before", "after"]),
        };

        if (target < 1 || target > count + 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"Cannot insert a slide at position {target}; the deck has {count} slide(s).",
                Units.Inv($"Valid positions are 1..{count + 1} (use /slide[{count + 1}] to append)."),
                candidates: [.. Enumerable.Range(1, Math.Min(count + 1, 10)).Select(i => Units.Inv($"/slide[{i}]"))]);
        }

        var layoutPart = PickLayout(presentation, props);

        var slidePart = presentation.AddNewPart<SlidePart>();
        slidePart.Slide = PptxFactory.BuildBlankSlide();
        slidePart.AddPart(layoutPart);

        var slideId = new P.SlideId
        {
            Id = PptxDoc.NextSlideId(slideIdList),
            RelationshipId = presentation.GetIdOfPart(slidePart),
        };
        slideIdList.InsertAt(slideId, target - 1);

        if (props is not null && props.TryGetPropertyValue("title", out var titleNode) && titleNode is not null)
        {
            AddTitleShape(slidePart, J.ScalarText(titleNode));
        }

        return Units.Inv($"/slide[{target}]");
    }

    /// <summary>
    /// The layout a new slide binds to: props.layout is a 1-based index into the
    /// first master's layouts (read --view structure lists them); the default
    /// stays the master's first layout.
    /// </summary>
    private static SlideLayoutPart PickLayout(PresentationPart presentation, JsonObject? props)
    {
        var masters = PptxDoc.Masters(presentation);
        if (masters.Count == 0)
        {
            throw CorruptPresentation("no slide master part exists");
        }

        var layouts = PptxDoc.Layouts(masters[0].Part);
        if (layouts.Count == 0)
        {
            throw CorruptPresentation("no slide layout part exists");
        }

        if (props is null || !props.TryGetPropertyValue("layout", out var layoutNode))
        {
            return layouts[0].Part;
        }

        // Props arrive string-valued through the CLI sugar and the MCP schema
        // ({"layout":"2"}), and as JSON numbers from hand-written ops — accept both.
        double number = 0;
        var numeric = layoutNode is JsonValue value &&
            (Units.TryNumber(value, out number) ||
             (value.TryGetValue<string>(out var raw) &&
              double.TryParse(raw, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out number)));
        if (!numeric || number != Math.Floor(number) || number < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"props.layout is not a valid layout index: {layoutNode?.ToJsonString() ?? "null"}",
                "Use a 1-based integer index into the master's layouts; run 'aioffice read <file> --view structure' to list them.");
        }

        var index = (int)number;
        if (index > layouts.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"props.layout is {index} but master 1 has only {layouts.Count} layout(s)."),
                "Run 'aioffice read <file> --view structure' to list the master's layouts.",
                candidates: [.. layouts.Take(10).Select(l => Units.Inv($"/master[1]/layout[{l.Index}]"))]);
        }

        return layouts[index - 1].Part;
    }

    private static string ApplySet(PresentationPart presentation, EditOp op)
    {
        var address = PptxAddress.Parse(op.Path);
        if (op.Props is null || op.Props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"set on '{op.Path}' has no props.",
                "Pass props, e.g. {\"op\":\"set\",\"path\":\"/slide[1]/shape[2]\",\"props\":{\"text\":\"Hello\"}}.");
        }

        if (!address.HasShape)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"set targets a shape or a paragraph, not the slide '{op.Path}'.",
                "Address a shape, e.g. /slide[1]/shape[2] or /slide[1]/shape[@id=4]/p[1].");
        }

        if (address.RunIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Run-level set is not supported yet.",
                "Set the paragraph instead: target /slide[i]/shape[j]/p[k] with a text prop.");
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, op.Path);
        var view = PptxDoc.ResolveShape(slidePart, address);

        if (address.ParagraphIndex is not null)
        {
            SetParagraphProps(view, address, op.Props);
            return Units.Inv($"{view.CanonicalPath(address.SlideIndex)}/p[{address.ParagraphIndex}]");
        }

        SetShapeProps(view, op.Props);
        return view.CanonicalPath(address.SlideIndex);
    }

    private static string ApplyRemove(PresentationPart presentation, EditOp op)
    {
        var address = PptxAddress.Parse(op.Path);
        if (address.RunIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Run-level remove is not supported yet.",
                "Remove the paragraph (/slide[i]/shape[j]/p[k]) or set the paragraph text instead.");
        }

        if (!address.HasShape)
        {
            RemoveSlide(presentation, address);
            return address.CanonicalSlidePath;
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, op.Path);
        var view = PptxDoc.ResolveShape(slidePart, address);

        if (address.ParagraphIndex is not null)
        {
            PptxDoc.ResolveParagraph(view, address).Remove();
            return Units.Inv($"{view.CanonicalPath(address.SlideIndex)}/p[{address.ParagraphIndex}]");
        }

        var canonical = view.CanonicalPath(address.SlideIndex);
        view.Element.Remove();
        return canonical;
    }

    private static void RemoveSlide(PresentationPart presentation, PptxAddress address)
    {
        var slides = PptxDoc.Slides(presentation);
        if (address.SlideIndex < 1 || address.SlideIndex > slides.Count)
        {
            _ = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw); // throws invalid_path with candidates
        }

        var (slideId, slidePart) = slides[address.SlideIndex - 1];
        var relId = slideId.RelationshipId?.Value;
        slideId.Remove();
        if (relId is not null)
        {
            presentation.DeletePart(relId);
        }
        else
        {
            presentation.DeletePart(slidePart);
        }
    }

    private static string ApplyMove(PresentationPart presentation, EditOp op)
    {
        var address = PptxAddress.Parse(op.Path);
        if (address.HasShape)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "Only slides can be moved in this milestone.",
                "To restyle shape stacking, remove the shape and add it again; slide moves use {\"op\":\"move\",\"path\":\"/slide[3]\",\"position\":\"1\"}.");
        }

        var slideIdList = presentation.Presentation?.SlideIdList
            ?? throw CorruptPresentation("p:sldIdLst is missing");
        var ids = slideIdList.Elements<P.SlideId>().ToList();
        var count = ids.Count;

        if (address.SlideIndex < 1 || address.SlideIndex > count)
        {
            _ = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        }

        var target = ParseMoveTarget(op.Position, address.SlideIndex, count);
        var moving = ids[address.SlideIndex - 1];
        moving.Remove();
        slideIdList.InsertAt(moving, target - 1);
        return Units.Inv($"/slide[{target}]");
    }

    /// <summary>Move targets: a 1-based final index ("3"), or "before:/slide[k]" / "after:/slide[k]".</summary>
    private static int ParseMoveTarget(string? position, int fromIndex, int count)
    {
        const string usage =
            "Use a 1-based destination index (\"1\"), or \"before:/slide[k]\" / \"after:/slide[k]\".";

        if (string.IsNullOrWhiteSpace(position))
        {
            throw new AiofficeException(ErrorCodes.InvalidArgs, "move requires a position.", usage);
        }

        var text = position.Trim();
        int target;
        if (int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var index))
        {
            target = index;
        }
        else if (TryParseAnchor(text, "before:", out var beforeAnchor))
        {
            RequireAnchorInRange(beforeAnchor, count, usage);
            target = fromIndex < beforeAnchor ? beforeAnchor - 1 : beforeAnchor;
        }
        else if (TryParseAnchor(text, "after:", out var afterAnchor))
        {
            RequireAnchorInRange(afterAnchor, count, usage);
            target = fromIndex < afterAnchor ? afterAnchor : afterAnchor + 1;
        }
        else
        {
            throw new AiofficeException(ErrorCodes.InvalidArgs, $"Unknown move position '{position}'.", usage);
        }

        if (target < 1 || target > count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Move destination {target} is out of range 1..{count}."),
                usage);
        }

        return target;
    }

    private static bool TryParseAnchor(string text, string prefix, out int anchor)
    {
        anchor = 0;
        if (!text.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var path = PptxAddress.Parse(text[prefix.Length..].Trim());
        if (path.HasShape || path.IsMaster)
        {
            return false;
        }

        anchor = path.SlideIndex;
        return true;
    }

    private static void RequireAnchorInRange(int anchor, int count, string usage)
    {
        if (anchor < 1 || anchor > count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Move anchor slide {anchor} is out of range 1..{count}."),
                usage);
        }
    }

    /// <summary>Adds a styled title placeholder with explicit geometry (used by create and add slide).</summary>
    internal static P.Shape AddTitleShape(SlidePart slidePart, string title)
    {
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = PptxDoc.NextShapeId(tree), Name = "Title" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 831_850L, Y = 365_125L },
                    new A.Extents { Cx = 10_515_600L, Cy = 1_325_563L }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                BuildParagraph(title, fontSizeHundredths: 4000, bold: null, colorHex: null, align: null)));
        tree.Append(shape);
        return shape;
    }

    /// <summary>Adds a textbox shape from props (x/y/w/h, text, fontSize, bold, color, fill, align, name).</summary>
    private static uint AddTextBox(SlidePart slidePart, JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            RequireKnownPropKey(key);
        }

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);

        var x = props.TryGetPropertyValue("x", out var xNode) ? Units.ParseLengthEmu("x", xNode) : Units.CmToEmu(2.5);
        var y = props.TryGetPropertyValue("y", out var yNode) ? Units.ParseLengthEmu("y", yNode) : Units.CmToEmu(2.5);
        var w = props.TryGetPropertyValue("w", out var wNode) ? Units.ParseLengthEmu("w", wNode) : Units.CmToEmu(10);
        var h = props.TryGetPropertyValue("h", out var hNode) ? Units.ParseLengthEmu("h", hNode) : Units.CmToEmu(3);

        var fontSize = props.TryGetPropertyValue("fontSize", out var sizeNode)
            ? Units.ParseFontSizeHundredths("fontSize", sizeNode)
            : (int?)null;
        var bold = props.TryGetPropertyValue("bold", out var boldNode) ? AsBool("bold", boldNode) : (bool?)null;
        var color = props.TryGetPropertyValue("color", out var colorNode) ? Units.ParseColorHex("color", colorNode) : null;
        A.TextAlignmentTypeValues? align = null;
        if (props.TryGetPropertyValue("align", out var alignNode))
        {
            align = ParseAlign(alignNode) ?? throw InvalidAlign(alignNode);
        }
        var text = props.TryGetPropertyValue("text", out var textNode) && textNode is not null
            ? J.ScalarText(textNode)
            : string.Empty;
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
            ? J.ScalarText(nameNode)
            : Units.Inv($"TextBox {id}");

        var shapeProperties = new P.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = w, Cy = h }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

        if (props.TryGetPropertyValue("fill", out var fillNode))
        {
            shapeProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = Units.ParseColorHex("fill", fillNode) }));
        }

        var body = new P.TextBody(new A.BodyProperties { Wrap = A.TextWrappingValues.Square }, new A.ListStyle());
        foreach (var line in text.Split('\n'))
        {
            body.Append(BuildParagraph(line, fontSize, bold, color, align));
        }

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties { TextBox = true },
                new P.ApplicationNonVisualDrawingProperties()),
            shapeProperties,
            body);
        tree.Append(shape);
        return id;
    }

    private static void SetShapeProps(ShapeView view, JsonObject props)
    {
        if (view.Element is not P.Shape shape)
        {
            var nameOnly = props.Count == 1 && props.ContainsKey("name");
            if (!nameOnly)
            {
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"set on a '{view.Kind}' supports only the name prop in this milestone.",
                    "Text, fill and geometry sets target text shapes; recreate other content in PowerPoint for now.");
            }
        }

        long? x = null, y = null, w = null, h = null;

        foreach (var (key, value) in props)
        {
            switch (RequireKnownPropKey(key))
            {
                case "text" or "title":
                    ReplaceText((P.Shape)view.Element, value is null ? string.Empty : J.ScalarText(value));
                    break;
                case "x":
                    x = Units.ParseLengthEmu(key, value);
                    break;
                case "y":
                    y = Units.ParseLengthEmu(key, value);
                    break;
                case "w":
                    w = Units.ParseLengthEmu(key, value);
                    break;
                case "h":
                    h = Units.ParseLengthEmu(key, value);
                    break;
                case "fill":
                    SetFill((P.Shape)view.Element, Units.ParseColorHex(key, value));
                    break;
                case "fontSize":
                    ApplyRunProps((P.Shape)view.Element, rPr => rPr.FontSize = Units.ParseFontSizeHundredths(key, value));
                    break;
                case "bold":
                    ApplyRunProps((P.Shape)view.Element, rPr => rPr.Bold = AsBool(key, value));
                    break;
                case "color":
                    var hex = Units.ParseColorHex(key, value);
                    ApplyRunProps((P.Shape)view.Element, rPr => SetRunColor(rPr, hex));
                    break;
                case "align":
                    var alignment = ParseAlign(value) ?? throw InvalidAlign(value);
                    foreach (var paragraph in ((P.Shape)view.Element).TextBody?.Elements<A.Paragraph>() ?? [])
                    {
                        SetAlignment(paragraph, alignment);
                    }

                    break;
                case "name":
                    var nonVisual = PptxDoc.NonVisualProps(view.Element) ?? throw CorruptPresentation("shape has no p:cNvPr");
                    nonVisual.Name = value is null ? string.Empty : J.ScalarText(value);
                    break;
            }
        }

        if (x is not null || y is not null || w is not null || h is not null)
        {
            SetGeometry((P.Shape)view.Element, x, y, w, h);
        }
    }

    private static void SetParagraphProps(ShapeView view, PptxAddress address, JsonObject props)
    {
        if (view.Element is not P.Shape)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{view.Kind}' shapes have no addressable paragraphs.",
                "Paragraph sets target text shapes; check the path with 'aioffice get'.");
        }

        var paragraph = PptxDoc.ResolveParagraph(view, address);
        foreach (var (key, value) in props)
        {
            switch (RequireKnownPropKey(key))
            {
                case "text" or "title":
                    ReplaceParagraphText(paragraph, value is null ? string.Empty : J.ScalarText(value));
                    break;
                case "fontSize":
                    ApplyParagraphRunProps(paragraph, rPr => rPr.FontSize = Units.ParseFontSizeHundredths(key, value));
                    break;
                case "bold":
                    ApplyParagraphRunProps(paragraph, rPr => rPr.Bold = AsBool(key, value));
                    break;
                case "color":
                    var hex = Units.ParseColorHex(key, value);
                    ApplyParagraphRunProps(paragraph, rPr => SetRunColor(rPr, hex));
                    break;
                case "align":
                    SetAlignment(paragraph, ParseAlign(value) ?? throw InvalidAlign(value));
                    break;
                default:
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Prop '{key}' does not apply to a paragraph.",
                        "Paragraph props: text, fontSize, bold, color, align. Geometry and fill belong to the shape.",
                        candidates: ["text", "fontSize", "bold", "color", "align"]);
            }
        }
    }

    internal static A.Paragraph BuildParagraph(string text, int? fontSizeHundredths, bool? bold, string? colorHex, A.TextAlignmentTypeValues? align)
    {
        var paragraph = new A.Paragraph();
        if (align is { } alignment)
        {
            paragraph.Append(new A.ParagraphProperties { Alignment = alignment });
        }

        var runProperties = new A.RunProperties { Language = "en-US" };
        if (fontSizeHundredths is { } size)
        {
            runProperties.FontSize = size;
        }

        if (bold is { } isBold)
        {
            runProperties.Bold = isBold;
        }

        if (colorHex is { } hex)
        {
            SetRunColor(runProperties, hex);
        }

        paragraph.Append(new A.Run(runProperties, new A.Text(text)));
        return paragraph;
    }

    /// <summary>Replaces the whole text body, keeping the first run's formatting as the prototype.</summary>
    internal static void ReplaceText(P.Shape shape, string text)
    {
        var body = shape.TextBody;
        if (body is null)
        {
            body = new P.TextBody(new A.BodyProperties(), new A.ListStyle());
            shape.Append(body);
        }

        var runPrototype = body.Descendants<A.RunProperties>().FirstOrDefault();
        var paragraphPrototype = body.Elements<A.Paragraph>().FirstOrDefault()?.ParagraphProperties;

        foreach (var paragraph in body.Elements<A.Paragraph>().ToList())
        {
            paragraph.Remove();
        }

        foreach (var line in text.Split('\n'))
        {
            var paragraph = new A.Paragraph();
            if (paragraphPrototype is not null)
            {
                paragraph.Append((A.ParagraphProperties)paragraphPrototype.CloneNode(true));
            }

            var run = new A.Run();
            if (runPrototype is not null)
            {
                run.Append((A.RunProperties)runPrototype.CloneNode(true));
            }

            run.Append(new A.Text(line));
            paragraph.Append(run);
            body.Append(paragraph);
        }
    }

    private static void ReplaceParagraphText(A.Paragraph paragraph, string text)
    {
        var runPrototype = paragraph.Elements<A.Run>().FirstOrDefault()?.RunProperties;
        var run = new A.Run();
        if (runPrototype is not null)
        {
            run.Append((A.RunProperties)runPrototype.CloneNode(true));
        }

        run.Append(new A.Text(text));

        foreach (var child in paragraph.ChildElements.Where(c => c is A.Run or A.Break or A.Field).ToList())
        {
            child.Remove();
        }

        paragraph.Append(run);
    }

    private static void SetGeometry(P.Shape shape, long? x, long? y, long? w, long? h)
    {
        var properties = shape.ShapeProperties;
        if (properties is null)
        {
            properties = new P.ShapeProperties();
            shape.InsertAfter(properties, shape.NonVisualShapeProperties);
        }

        var transform = properties.Transform2D;
        if (transform is null)
        {
            transform = new A.Transform2D(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = Units.CmToEmu(10), Cy = Units.CmToEmu(3) });
            properties.InsertAt(transform, 0);
        }

        transform.Offset ??= new A.Offset { X = 0L, Y = 0L };
        transform.Extents ??= new A.Extents { Cx = Units.CmToEmu(10), Cy = Units.CmToEmu(3) };

        if (x is not null)
        {
            transform.Offset.X = x;
        }

        if (y is not null)
        {
            transform.Offset.Y = y;
        }

        if (w is not null)
        {
            transform.Extents.Cx = w;
        }

        if (h is not null)
        {
            transform.Extents.Cy = h;
        }
    }

    private static void SetFill(P.Shape shape, string hex)
    {
        var properties = shape.ShapeProperties;
        if (properties is null)
        {
            properties = new P.ShapeProperties();
            shape.InsertAfter(properties, shape.NonVisualShapeProperties);
        }

        foreach (var fill in properties.ChildElements
            .Where(c => c is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill)
            .ToList())
        {
            fill.Remove();
        }

        var solidFill = new A.SolidFill(new A.RgbColorModelHex { Val = hex });
        OpenXmlElement? anchor = (OpenXmlElement?)properties.GetFirstChild<A.PresetGeometry>()
            ?? (OpenXmlElement?)properties.GetFirstChild<A.CustomGeometry>()
            ?? properties.Transform2D;
        if (anchor is not null)
        {
            properties.InsertAfter(solidFill, anchor);
        }
        else
        {
            properties.InsertAt(solidFill, 0);
        }
    }

    private static void ApplyRunProps(P.Shape shape, Action<A.RunProperties> mutate)
    {
        foreach (var paragraph in shape.TextBody?.Elements<A.Paragraph>() ?? [])
        {
            ApplyParagraphRunProps(paragraph, mutate);
        }
    }

    private static void ApplyParagraphRunProps(A.Paragraph paragraph, Action<A.RunProperties> mutate)
    {
        foreach (var run in paragraph.Elements<A.Run>())
        {
            var runProperties = run.RunProperties;
            if (runProperties is null)
            {
                runProperties = new A.RunProperties { Language = "en-US" };
                run.InsertAt(runProperties, 0);
            }

            mutate(runProperties);
        }
    }

    private static void SetRunColor(A.RunProperties runProperties, string hex)
    {
        foreach (var fill in runProperties.ChildElements.Where(c => c is A.SolidFill or A.NoFill or A.GradientFill).ToList())
        {
            fill.Remove();
        }

        runProperties.InsertAt(new A.SolidFill(new A.RgbColorModelHex { Val = hex }), 0);
    }

    private static void SetAlignment(A.Paragraph paragraph, A.TextAlignmentTypeValues alignment)
    {
        var properties = paragraph.ParagraphProperties;
        if (properties is null)
        {
            properties = new A.ParagraphProperties();
            paragraph.InsertAt(properties, 0);
        }

        properties.Alignment = alignment;
    }

    internal static A.TextAlignmentTypeValues? ParseAlign(JsonNode? node)
    {
        var text = node is null ? null : J.ScalarText(node).Trim().ToLowerInvariant();
        return text switch
        {
            "left" => A.TextAlignmentTypeValues.Left,
            "center" => A.TextAlignmentTypeValues.Center,
            "right" => A.TextAlignmentTypeValues.Right,
            "justify" => A.TextAlignmentTypeValues.Justified,
            _ => null,
        };
    }

    private static AiofficeException InvalidAlign(JsonNode? node) => new(
        ErrorCodes.InvalidArgs,
        $"Not a valid align value: {node?.ToJsonString() ?? "null"}",
        "Use left, center, right or justify.",
        candidates: ["left", "center", "right", "justify"]);

    private static bool AsBool(string key, JsonNode? node)
    {
        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var flag))
            {
                return flag;
            }

            if (value.TryGetValue<string>(out var text) && bool.TryParse(text, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Property '{key}' is not a boolean: {node?.ToJsonString() ?? "null"}",
            "Use true or false.");
    }

    private static string RequireKnownPropKey(string key)
    {
        if (ShapePropKeys.Contains(key, StringComparer.Ordinal))
        {
            return key;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Unknown shape prop '{key}'.",
            "Run 'aioffice help properties' for the per-type prop list.",
            candidates: ShapePropKeys);
    }

    private static AiofficeException CorruptPresentation(string detail) => new(
        ErrorCodes.FormatCorrupt,
        $"The presentation is malformed: {detail}.",
        "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
}
