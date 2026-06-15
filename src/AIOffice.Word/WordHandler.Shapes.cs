using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Dw = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace AIOffice.Word;

/// <summary>
/// v1.3.0 body drawing shapes and text boxes: a floating DrawingML
/// <c>wps:wsp</c> (the modern WordprocessingShape, the one Word's Insert &gt;
/// Shapes / Text Box galleries write) anchored in the body. A shape carries a
/// preset geometry (rect/roundRect/ellipse/line/arrow) with an optional fill,
/// outline and inline text; a text box is the same shape forced to a rectangle
/// with required text and a no/square wrap. Both are addressed positionally
/// (<c>/body/shape[i]</c>, <c>/body/textBox[i]</c>), reported by <c>get</c>,
/// edited by <c>set</c> (geometry/fill/line/text) and removed.
///
/// Distinct from inline images (a <c>pic:pic</c> drawing) and from pptx shapes.
/// </summary>
public sealed partial class WordHandler
{
    private const string WordprocessingShapeUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

    /// <summary>The body shape vocabulary -> its DrawingML preset geometry.</summary>
    private static readonly Dictionary<string, A.ShapeTypeValues> ShapePresets = new(StringComparer.Ordinal)
    {
        ["rect"] = A.ShapeTypeValues.Rectangle,
        ["roundRect"] = A.ShapeTypeValues.RoundRectangle,
        ["ellipse"] = A.ShapeTypeValues.Ellipse,
        ["line"] = A.ShapeTypeValues.Line,
        ["arrow"] = A.ShapeTypeValues.Line,
    };

    private static readonly string[] ShapeKinds = ["rect", "roundRect", "ellipse", "line", "arrow"];

    // ------------------------------------------------------------------- add

    /// <summary>
    /// <c>{"op":"add","path":"/body","type":"shape","props":{"shape":"rect","x":"2cm","y":"2cm","w":"6cm","h":"3cm","fill"?,"line"?,"text"?}}</c>
    /// and <c>{"op":"add","path":"/body","type":"textBox","props":{"x","y","w","h","text","fill"?,"wrap"?}}</c>:
    /// a floating wps:wsp anchored in a new (or the anchor) paragraph.
    /// </summary>
    private static object ApplyAddBodyShape(WordprocessingDocument doc, EditOp op, EditSession session, bool isTextBox)
    {
        var props = op.Props?.DeepClone().AsObject() ?? [];
        var author = session.ResolveAuthor(props);

        var kind = isTextBox
            ? "rect"
            : props["shape"] is { } shapeNode ? NodeToString(shapeNode) : null;
        if (!isTextBox && kind is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type shape needs props.shape (rect, roundRect, ellipse, line or arrow).",
                "Example: {\"op\":\"add\",\"path\":\"/body\",\"type\":\"shape\",\"props\":{\"shape\":\"rect\",\"x\":\"2cm\",\"y\":\"2cm\",\"w\":\"6cm\",\"h\":\"3cm\"}}.",
                candidates: ShapeKinds);
        }

        if (!ShapePresets.ContainsKey(kind!))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Shape '{kind}' is not supported.",
                "Use shape rect, roundRect, ellipse, line or arrow.",
                candidates: ShapeKinds);
        }

        var isLine = kind is "line" or "arrow";
        var xEmu = props["x"] is { } xn ? ParseLengthEmu("x", NodeToString(xn)) : 0L;
        var yEmu = props["y"] is { } yn ? ParseLengthEmu("y", NodeToString(yn)) : 0L;
        var wEmu = props["w"] is { } wn ? ParseLengthEmu("w", NodeToString(wn))
            : props["width"] is { } wn2 ? ParseLengthEmu("w", NodeToString(wn2))
            : (long)(EmuPerCm * 6);

        // A line is 1-D: its height is always 0 (the prop is ignored, so "0cm" is
        // fine even though ParseLengthEmu rejects a non-positive length).
        var hEmu = isLine
            ? 0L
            : props["h"] is { } hn ? ParseLengthEmu("h", NodeToString(hn))
                : props["height"] is { } hn2 ? ParseLengthEmu("h", NodeToString(hn2))
                    : (long)(EmuPerCm * 3);

        // A line's height of 0 is legitimate (a horizontal rule); everything else
        // must enclose a positive area so Word can render it.
        if (wEmu <= 0 || (hEmu <= 0 && !isLine))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Shape dimensions must be positive; computed {wEmu}x{hEmu} EMU.",
                "Pass a positive w/h like \"6cm\", \"3cm\" (a line may have h=0).");
        }

        var fill = props["fill"] is { } fillNode ? WordFormatting.ParseHexColor(NodeToString(fillNode)) : null;
        var lineColor = props["line"] is { } lineNode ? WordFormatting.ParseHexColor(NodeToString(lineNode)) : null;
        var text = props["text"] is { } textNode ? NodeToString(textNode) : null;
        if (isTextBox && string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type textBox needs props.text.",
                "Example: {\"op\":\"add\",\"path\":\"/body\",\"type\":\"textBox\",\"props\":{\"x\":\"2cm\",\"y\":\"2cm\",\"w\":\"6cm\",\"h\":\"3cm\",\"text\":\"Note\"}}.");
        }

        var wrap = props["wrap"] is { } wrapNode ? NodeToString(wrapNode) : null;

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
                $"Cannot add a {(isTextBox ? "text box" : "shape")} {position} {anchor.CanonicalPath} ({anchor.Type}).",
                "Add shapes inside /body (or a tc/header/footer), or before/after an existing paragraph.");
        }

        // A reopen+save hoists the wp/a/wps namespaces to the w:document root;
        // declaring them there up-front keeps the round-trip law byte-identical.
        DeclareShapeNamespaces(doc);

        var docPrId = NextDrawingId(doc);
        var name = (isTextBox ? "Text Box " : "Shape ") + docPrId.ToString(CultureInfo.InvariantCulture);
        var anchorElement = BuildShapeAnchor(
            kind!, docPrId, name, xEmu, yEmu, wEmu, hEmu, fill, lineColor, isLine, isTextBox, text, wrap);

        var paragraph = new Paragraph(new Run(new Drawing(anchorElement)));
        if (session.Track)
        {
            RequireBodyScope(anchor.CanonicalPath, "add");
            MarkParagraphInserted(doc, paragraph, author);
        }

        var canonical = Insert(doc, anchor, paragraph, position);
        var shapePath = BodyShapePathFor(doc, anchorElement, isTextBox);
        return new
        {
            op = "add",
            type = isTextBox ? "textBox" : "shape",
            path = shapePath ?? canonical,
            shape = kind,
            xCm = EmuToCm(xEmu),
            yCm = EmuToCm(yEmu),
            wCm = EmuToCm(wEmu),
            hCm = EmuToCm(hEmu),
        };
    }

    private const string WordprocessingDrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private const string DrawingMainNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>
    /// Declares wp/a/wps on the w:document root before adding a shape: a reopen
    /// +save hoists these declarations there, so declaring them up-front keeps the
    /// round-trip law byte-identical (the same fix the watermark uses for VML).
    /// </summary>
    private static void DeclareShapeNamespaces(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document is not { } document)
        {
            return;
        }

        if (document.LookupNamespace("wp") is null)
        {
            document.AddNamespaceDeclaration("wp", WordprocessingDrawingNamespace);
        }

        if (document.LookupNamespace("a") is null)
        {
            document.AddNamespaceDeclaration("a", DrawingMainNamespace);
        }

        if (document.LookupNamespace("wps") is null)
        {
            document.AddNamespaceDeclaration("wps", WordprocessingShapeUri);
        }
    }

    /// <summary>The floating wp:anchor &gt; a:graphic &gt; wps:wsp markup Word writes for a body shape/text box.</summary>
    private static Dw.Anchor BuildShapeAnchor(
        string kind, uint docPrId, string name, long x, long y, long cx, long cy,
        string? fill, string? lineColor, bool isLine, bool isTextBox, string? text, string? wrap)
    {
        var spPr = new Wps.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = cx, Cy = cy }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = ShapePresets[kind] });

        if (isLine)
        {
            // A line/arrow has no fill; its visible body is the outline. An arrow
            // adds a tail-end arrowhead.
            spPr.AppendChild(new A.NoFill());
            var outline = new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = lineColor ?? "333333" }))
            {
                Width = 12700, // 1pt
            };
            if (kind == "arrow")
            {
                outline.AppendChild(new A.TailEnd { Type = A.LineEndValues.Arrow });
            }

            spPr.AppendChild(outline);
        }
        else
        {
            spPr.AppendChild(fill is { Length: > 0 }
                ? new A.SolidFill(new A.RgbColorModelHex { Val = fill })
                : (OpenXmlElement)new A.SolidFill(new A.RgbColorModelHex { Val = isTextBox ? "FFFFFF" : "DDDDDD" }));
            if (lineColor is { Length: > 0 })
            {
                spPr.AppendChild(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = lineColor })) { Width = 12700 });
            }
        }

        var wsp = new Wps.WordprocessingShape(
            new Wps.NonVisualDrawingProperties { Id = 0U, Name = name },
            new Wps.NonVisualDrawingShapeProperties(),
            spPr);

        if (!isLine && text is { Length: > 0 })
        {
            wsp.AppendChild(new Wps.TextBoxInfo2(new TextBoxContent(new Paragraph(new Run(NewText(text))))));
            wsp.AppendChild(new Wps.TextBodyProperties { Rotation = 0, UseParagraphSpacing = false });
        }
        else
        {
            wsp.AppendChild(new Wps.TextBodyProperties { Rotation = 0, UseParagraphSpacing = false });
        }

        // wrap: textBox defaults to "square" (text flows around it); a shape
        // defaults to "none" (floats free, no flow). Honor an explicit override.
        var wrapMode = wrap ?? (isTextBox ? "square" : "none");
        OpenXmlElement wrapElement = wrapMode switch
        {
            "square" => new Dw.WrapSquare { WrapText = Dw.WrapTextValues.BothSides },
            "tight" => new Dw.WrapTight { WrapText = Dw.WrapTextValues.BothSides },
            "topAndBottom" => new Dw.WrapTopBottom(),
            _ => new Dw.WrapNone(),
        };

        return new Dw.Anchor(
            new Dw.SimplePosition { X = 0L, Y = 0L },
            new Dw.HorizontalPosition(new Dw.PositionOffset(x.ToString(CultureInfo.InvariantCulture)))
            {
                RelativeFrom = Dw.HorizontalRelativePositionValues.Column,
            },
            new Dw.VerticalPosition(new Dw.PositionOffset(y.ToString(CultureInfo.InvariantCulture)))
            {
                RelativeFrom = Dw.VerticalRelativePositionValues.Paragraph,
            },
            new Dw.Extent { Cx = cx, Cy = Math.Max(cy, 0L) },
            new Dw.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            wrapElement,
            new Dw.DocProperties { Id = docPrId, Name = name },
            new Dw.NonVisualGraphicFrameDrawingProperties(),
            new A.Graphic(new A.GraphicData(wsp) { Uri = WordprocessingShapeUri }))
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 114300U,
            DistanceFromRight = 114300U,
            SimplePos = false,
            RelativeHeight = 251658240U + docPrId,
            BehindDoc = false,
            Locked = false,
            LayoutInCell = true,
            AllowOverlap = true,
        };
    }

    // ------------------------------------------------------------------- set

    /// <summary>set /body/shape[i] or /body/textBox[i]: updates geometry, fill, line and/or text.</summary>
    private static object ApplySetBodyShape(WordprocessingDocument doc, EditOp op, bool isTextBox)
    {
        var (anchorEl, index) = ResolveBodyShape(doc, DocPath.Parse(op.Path), isTextBox);
        var props = RequireProps(op);
        var wsp = anchorEl.Descendants<Wps.WordprocessingShape>().First();
        var spPr = wsp.GetFirstChild<Wps.ShapeProperties>()!;
        var isLine = IsLineShape(wsp);

        foreach (var (rawName, value) in props.Select(kv => (kv.Key, NodeToString(kv.Value))))
        {
            switch (rawName)
            {
                case "x":
                    SetOffset(anchorEl, horizontal: true, ParseLengthEmu("x", value));
                    break;
                case "y":
                    SetOffset(anchorEl, horizontal: false, ParseLengthEmu("y", value));
                    break;
                case "w":
                case "width":
                    SetExtent(anchorEl, spPr, ParseLengthEmu("w", value), null);
                    break;
                case "h":
                case "height":
                    SetExtent(anchorEl, spPr, null, ParseLengthEmu("h", value));
                    break;
                case "fill":
                    SetShapeFill(spPr, WordFormatting.ParseHexColor(value), isLine);
                    break;
                case "line":
                    SetShapeOutline(spPr, WordFormatting.ParseHexColor(value));
                    break;
                case "shape":
                    SetShapePreset(spPr, value);
                    break;
                case "text":
                    SetShapeText(wsp, value, isLine);
                    break;
                default:
                    throw new AiofficeException(
                        ErrorCodes.UnsupportedFeature,
                        $"Property '{rawName}' is not supported on a body {(isTextBox ? "text box" : "shape")}.",
                        "Set x, y, w, h, fill, line, shape or text.",
                        candidates: ["x", "y", "w", "h", "fill", "line", "shape", "text"]);
            }
        }

        return new { op = "set", path = BodyShapePath(index, isTextBox), type = isTextBox ? "textBox" : "shape" };
    }

    private static void SetOffset(Dw.Anchor anchor, bool horizontal, long emu)
    {
        var offset = horizontal
            ? anchor.GetFirstChild<Dw.HorizontalPosition>()?.GetFirstChild<Dw.PositionOffset>()
            : anchor.GetFirstChild<Dw.VerticalPosition>()?.GetFirstChild<Dw.PositionOffset>();
        if (offset is not null)
        {
            offset.Text = emu.ToString(CultureInfo.InvariantCulture);
        }
    }

    private static void SetExtent(Dw.Anchor anchor, Wps.ShapeProperties spPr, long? cx, long? cy)
    {
        var extent = anchor.GetFirstChild<Dw.Extent>();
        var extents = spPr.Transform2D?.Extents;
        if (cx is { } w)
        {
            if (extent is not null)
            {
                extent.Cx = w;
            }

            if (extents is not null)
            {
                extents.Cx = w;
            }
        }

        if (cy is { } h)
        {
            if (extent is not null)
            {
                extent.Cy = Math.Max(h, 0L);
            }

            if (extents is not null)
            {
                extents.Cy = Math.Max(h, 0L);
            }
        }
    }

    private static void SetShapeFill(Wps.ShapeProperties spPr, string hex, bool isLine)
    {
        if (isLine)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "A line/arrow has no fill; set its 'line' color instead.",
                "Use {\"line\":\"RRGGBB\"} to recolor a line or arrow.");
        }

        foreach (var existing in spPr.Elements<A.SolidFill>().ToList())
        {
            existing.Remove();
        }

        foreach (var noFill in spPr.Elements<A.NoFill>().ToList())
        {
            noFill.Remove();
        }

        // The fill belongs right after the geometry, before any outline.
        var outline = spPr.GetFirstChild<A.Outline>();
        var fill = new A.SolidFill(new A.RgbColorModelHex { Val = hex });
        if (outline is not null)
        {
            outline.InsertBeforeSelf(fill);
        }
        else
        {
            spPr.AppendChild(fill);
        }
    }

    private static void SetShapeOutline(Wps.ShapeProperties spPr, string hex)
    {
        var outline = spPr.GetFirstChild<A.Outline>();
        if (outline is null)
        {
            outline = new A.Outline { Width = 12700 };
            spPr.AppendChild(outline);
        }

        foreach (var fill in outline.Elements<A.SolidFill>().ToList())
        {
            fill.Remove();
        }

        outline.InsertAt(new A.SolidFill(new A.RgbColorModelHex { Val = hex }), 0);
    }

    private static void SetShapePreset(Wps.ShapeProperties spPr, string kind)
    {
        if (!ShapePresets.TryGetValue(kind, out var preset))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"Shape '{kind}' is not supported.",
                "Use shape rect, roundRect, ellipse, line or arrow.",
                candidates: ShapeKinds);
        }

        var geometry = spPr.GetFirstChild<A.PresetGeometry>();
        if (geometry is not null)
        {
            geometry.Preset = preset;
        }
    }

    private static void SetShapeText(Wps.WordprocessingShape wsp, string text, bool isLine)
    {
        if (isLine)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "A line/arrow carries no text.",
                "Set text on a rect, roundRect, ellipse shape or a text box.");
        }

        var box = wsp.GetFirstChild<Wps.TextBoxInfo2>();
        if (box is null)
        {
            box = new Wps.TextBoxInfo2(new TextBoxContent(new Paragraph(new Run(NewText(text)))));
            // The text box must precede the body properties.
            var bodyPr = wsp.GetFirstChild<Wps.TextBodyProperties>();
            if (bodyPr is not null)
            {
                bodyPr.InsertBeforeSelf(box);
            }
            else
            {
                wsp.AppendChild(box);
                wsp.AppendChild(new Wps.TextBodyProperties());
            }

            return;
        }

        var content = box.GetFirstChild<TextBoxContent>() ?? box.AppendChild(new TextBoxContent());
        var paragraph = content.GetFirstChild<Paragraph>();
        if (paragraph is null)
        {
            paragraph = new Paragraph();
            content.AppendChild(paragraph);
        }

        WordFormatting.ReplaceParagraphText(paragraph, text);
        foreach (var extra in content.Elements<Paragraph>().Skip(1).ToList())
        {
            extra.Remove();
        }
    }

    // ---------------------------------------------------------------- remove

    /// <summary>remove /body/shape[i] or /body/textBox[i]: drops the carrier paragraph (or just the run if it is shared).</summary>
    private static object ApplyRemoveBodyShape(WordprocessingDocument doc, EditOp op, bool isTextBox)
    {
        var (anchorEl, index) = ResolveBodyShape(doc, DocPath.Parse(op.Path), isTextBox);
        var run = anchorEl.Ancestors<Run>().FirstOrDefault();
        var paragraph = anchorEl.Ancestors<Paragraph>().FirstOrDefault();

        // The carrier paragraph holds only this drawing -> remove the paragraph;
        // otherwise (shared paragraph) drop just the run.
        if (paragraph is not null && run is not null &&
            paragraph.Elements<Run>().Count() == 1 &&
            paragraph.InnerText.Length == 0 &&
            paragraph.Parent is not TableCell &&
            !(paragraph.Parent is Body && ReferenceEquals(paragraph, paragraph.Parent.Elements<Paragraph>().LastOrDefault())
              && paragraph.Parent.Elements<Paragraph>().Count() == 1))
        {
            paragraph.Remove();
        }
        else if (run is not null)
        {
            run.Remove();
        }
        else
        {
            anchorEl.Ancestors<Drawing>().FirstOrDefault()?.Remove();
        }

        return new { op = "remove", path = BodyShapePath(index, isTextBox), type = isTextBox ? "textBox" : "shape" };
    }

    // ------------------------------------------------------------------- get

    /// <summary>get /body/shape[i] or /body/textBox[i] -> geometry + style + text.</summary>
    private static (string Path, Dictionary<string, object?> Properties) GetBodyShapeProperties(
        WordprocessingDocument doc, DocPath path, bool isTextBox)
    {
        var (anchorEl, index) = ResolveBodyShape(doc, path, isTextBox);
        return (BodyShapePath(index, isTextBox), BodyShapeShape(anchorEl, isTextBox));
    }

    private static Dictionary<string, object?> BodyShapeShape(Dw.Anchor anchor, bool isTextBox)
    {
        var wsp = anchor.Descendants<Wps.WordprocessingShape>().First();
        var spPr = wsp.GetFirstChild<Wps.ShapeProperties>();
        var preset = spPr?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value;
        var extent = anchor.GetFirstChild<Dw.Extent>();
        var text = wsp.GetFirstChild<Wps.TextBoxInfo2>()?.InnerText;

        return new Dictionary<string, object?>
        {
            ["kind"] = isTextBox ? "textBox" : "shape",
            ["shape"] = PresetName(preset, IsLineShape(wsp)),
            ["xCm"] = ParseOffsetCm(anchor, horizontal: true),
            ["yCm"] = ParseOffsetCm(anchor, horizontal: false),
            ["wCm"] = extent?.Cx?.Value is { } cx ? EmuToCm(cx) : null,
            ["hCm"] = extent?.Cy?.Value is { } cy ? EmuToCm(cy) : null,
            ["fill"] = IsLineShape(wsp) ? null : ShapeFillHex(spPr),
            ["line"] = ShapeOutlineHex(spPr),
            ["text"] = string.IsNullOrEmpty(text) ? null : text,
        };
    }

    private static double? ParseOffsetCm(Dw.Anchor anchor, bool horizontal)
    {
        var offset = horizontal
            ? anchor.GetFirstChild<Dw.HorizontalPosition>()?.GetFirstChild<Dw.PositionOffset>()
            : anchor.GetFirstChild<Dw.VerticalPosition>()?.GetFirstChild<Dw.PositionOffset>();
        return offset?.Text is { } t && long.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out var emu)
            ? EmuToCm(emu)
            : null;
    }

    private static string? ShapeFillHex(Wps.ShapeProperties? spPr) =>
        spPr?.Elements<A.SolidFill>().FirstOrDefault()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;

    private static string? ShapeOutlineHex(Wps.ShapeProperties? spPr) =>
        spPr?.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;

    /// <summary>The shape kind name ("rect"/"line"/"arrow"…) reported for a stored preset.</summary>
    private static string? PresetName(A.ShapeTypeValues? preset, bool isLine)
    {
        if (preset is null)
        {
            return null;
        }

        if (preset == A.ShapeTypeValues.Rectangle)
        {
            return "rect";
        }

        if (preset == A.ShapeTypeValues.RoundRectangle)
        {
            return "roundRect";
        }

        if (preset == A.ShapeTypeValues.Ellipse)
        {
            return "ellipse";
        }

        // A line preset is a "line", or an "arrow" when it carries a tail arrowhead.
        return isLine ? "line/arrow" : preset.Value.ToString();
    }

    private static bool IsLineShape(Wps.WordprocessingShape wsp) =>
        wsp.GetFirstChild<Wps.ShapeProperties>()?.GetFirstChild<A.PresetGeometry>()?.Preset?.Value is { } p
        && p == A.ShapeTypeValues.Line;

    // ------------------------------------------------------------ addressing

    /// <summary>Every floating wps:wsp drawing anchor in the body, in document order, split by carries-text.</summary>
    private static List<Dw.Anchor> EnumerateBodyShapes(WordprocessingDocument doc, bool textBoxesOnly)
    {
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body is null)
        {
            return [];
        }

        return [.. body.Descendants<Dw.Anchor>()
            .Where(a => a.Descendants<Wps.WordprocessingShape>().Any())
            .Where(a => ShapeIsTextBox(a) == textBoxesOnly)];
    }

    private const string TextBoxNamePrefix = "Text Box ";

    /// <summary>
    /// A shape and a text box are disjoint, stable index spaces. The distinction
    /// is structural — fixed at creation by the wp:docPr name prefix ("Text Box "
    /// vs "Shape ") — NOT by whether the shape currently holds text. That keeps
    /// /body/shape[i] addressing stable even after you set text on a shape (and
    /// /body/textBox[i] stable even after you clear a text box's text).
    /// </summary>
    private static bool ShapeIsTextBox(Dw.Anchor anchor) =>
        anchor.GetFirstChild<Dw.DocProperties>()?.Name?.Value?.StartsWith(TextBoxNamePrefix, StringComparison.Ordinal) == true;

    /// <summary>The canonical path of a freshly inserted anchor, by locating it in the proper index space.</summary>
    private static string? BodyShapePathFor(WordprocessingDocument doc, Dw.Anchor anchor, bool isTextBox)
    {
        var list = EnumerateBodyShapes(doc, isTextBox);
        var i = list.FindIndex(a => ReferenceEquals(a, anchor));
        return i >= 0 ? BodyShapePath(i + 1, isTextBox) : null;
    }

    /// <summary>True for a two-segment /body/shape[i] or /body/textBox[i] path (the set/remove dispatch guard).</summary>
    private static bool IsBodyShapePath(DocPath path, string typeName) =>
        path.Segments.Count == 2 &&
        path.Segments[0].Name == "body" && path.Segments[0].Index is null &&
        path.Segments[1].Name == typeName;

    private static (Dw.Anchor Anchor, int Index) ResolveBodyShape(WordprocessingDocument doc, DocPath path, bool isTextBox)
    {
        var typeName = isTextBox ? "textBox" : "shape";
        var segment = path.Segments.Count == 2 ? path.Segments[1] : null;
        var shapes = EnumerateBodyShapes(doc, isTextBox);

        if (path.Segments.Count != 2 ||
            path.Segments[0].Name != "body" || path.Segments[0].Index is not null ||
            segment!.Name != typeName || segment.Id is not null)
        {
            throw BodyShapeNotFound(
                $"'{path.ToCanonicalString()}' is not a /body/{typeName}[i] path.", typeName, shapes);
        }

        var index = segment.Index ?? 1;
        if (shapes.Count == 0 || index > shapes.Count)
        {
            throw BodyShapeNotFound(
                shapes.Count == 0
                    ? $"This document has no body {(isTextBox ? "text boxes" : "shapes")}."
                    : $"/body/{typeName}[{index}] does not exist; there are {shapes.Count}.",
                typeName,
                shapes);
        }

        return (shapes[index - 1], index);
    }

    private static AiofficeException BodyShapeNotFound(string message, string typeName, List<Dw.Anchor> shapes) => new(
        ErrorCodes.InvalidPath,
        message,
        $"Add one with {{\"op\":\"add\",\"path\":\"/body\",\"type\":\"{typeName}\",\"props\":{{…}}}}, " +
        $"or run 'aioffice read <file> --view structure' to list shapes.",
        candidates: [.. Enumerable.Range(1, shapes.Count).Take(5).Select(i => BodyShapePath(i, typeName == "textBox"))]);

    private static string BodyShapePath(int index, bool isTextBox) =>
        string.Create(CultureInfo.InvariantCulture, $"/body/{(isTextBox ? "textBox" : "shape")}[{index}]");

    /// <summary>read --view structure shapes/textBoxes lists (kept compact: path + shape + size + text snippet).</summary>
    private static List<object> BodyShapesStructure(WordprocessingDocument doc, bool textBoxes) =>
        [.. EnumerateBodyShapes(doc, textBoxes).Select((a, i) => (object)new
        {
            path = BodyShapePath(i + 1, textBoxes),
            properties = BodyShapeShape(a, textBoxes),
        })];
}
