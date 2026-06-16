using System.Text;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Approximate-but-inspectable rendering: each slide becomes an SVG of
/// positioned shapes (geometry-true outlines), text runs and mini-charts drawn
/// from the chart's cached data.
/// </summary>
internal static class PptxRenderer
{
    private const double DefaultFontPt = 18;

    /// <summary>The accent palette charts cycle through (matches the built-in theme).</summary>
    private static readonly string[] ChartPalette =
        ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"];

    public static (double WidthPx, double HeightPx) SlideSizePx(PresentationPart presentation)
    {
        var size = presentation.Presentation?.SlideSize;
        var cx = (long)(size?.Cx?.Value ?? PptxFactory.SlideWidthEmu);
        var cy = (long)(size?.Cy?.Value ?? PptxFactory.SlideHeightEmu);
        return (Units.EmuToPx(cx), Units.EmuToPx(cy));
    }

    /// <summary>
    /// The first gradient stop colour of a shape's 1.8 gradient fill, used only as a
    /// flat render approximation (the document fill stays gradient/blip in OOXML).
    /// </summary>
    private static string? ShapeGradientStartHex(DocumentFormat.OpenXml.OpenXmlCompositeElement element)
    {
        var properties = (element as P.Shape)?.ShapeProperties;
        return GradientStartHex(properties?.GetFirstChild<A.GradientFill>());
    }

    /// <summary>The first gradient stop colour of a slide/master/layout background, for the flat render fallback.</summary>
    private static string? BackgroundGradientStartHex(P.CommonSlideData? slideData)
    {
        var gradient = slideData?.Background?.BackgroundProperties?.GetFirstChild<A.GradientFill>();
        return GradientStartHex(gradient);
    }

    /// <summary>The lowest-positioned stop's RGB hex (the gradient's visual "start").</summary>
    private static string? GradientStartHex(A.GradientFill? gradient)
    {
        var stop = gradient?.GradientStopList?.Elements<A.GradientStop>()
            .OrderBy(s => s.Position?.Value ?? 0)
            .FirstOrDefault();
        return stop?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();
    }

    public static string RenderSlideSvg(PresentationPart presentation, SlidePart slidePart, int slideIndex)
    {
        var (width, height) = SlideSizePx(presentation);
        var background = PptxDoc.BackgroundHex(slidePart)
            ?? BackgroundGradientStartHex(slidePart.Slide?.CommonSlideData);
        var svg = new StringBuilder();
        svg.Append(Units.Inv($"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 {width:0.#} {height:0.#}\" "));
        svg.Append(Units.Inv($"width=\"{width:0.#}\" height=\"{height:0.#}\" font-family=\"Helvetica, Arial, sans-serif\" "));
        svg.Append(Units.Inv($"data-slide=\"{slideIndex}\">\n"));
        svg.Append(Units.Inv($"  <rect x=\"0\" y=\"0\" width=\"{width:0.#}\" height=\"{height:0.#}\" "));
        svg.Append(Units.Inv($"fill=\"#{background?.ToLowerInvariant() ?? "ffffff"}\" stroke=\"#cccccc\"/>\n"));

        foreach (var shape in PptxDoc.Shapes(slidePart))
        {
            AppendShape(svg, presentation, slidePart, shape, slideIndex);
        }

        svg.Append("</svg>");
        return svg.ToString();
    }

    /// <summary>
    /// Emits one shape wrapped in a group carrying the data-aio-path render
    /// contract: a browser click on any child maps back to the canonical
    /// stable-id document path.
    /// </summary>
    private static void AppendShape(StringBuilder svg, PresentationPart presentation, SlidePart slidePart, ShapeView shape, int slideIndex)
    {
        var geometry = PptxDoc.Geometry(shape.Element) ?? new GeometryEmu(0, 0, Units.CmToEmu(4), Units.CmToEmu(1.5));
        var x = Units.EmuToPx(geometry.X);
        var y = Units.EmuToPx(geometry.Y);
        var w = Units.EmuToPx(geometry.Cx);
        var h = Units.EmuToPx(geometry.Cy);

        // A solid fill renders true; a 1.8 gradient/image fill renders as a flat
        // approximation (the gradient's start stop, or a tint marker) so the shape
        // shows colour rather than blank white — get/audit still read solid-only.
        var fill = PptxDoc.FillHex(shape.Element) ?? ShapeGradientStartHex(shape.Element);

        // A linked shape carries its resolved target on the group so a browser /
        // assistive tech can surface the click action (the render contract's
        // sibling to data-aio-path).
        var hyperlink = PptxHyperlinks.Resolve(presentation, slidePart, shape.Element);
        var linkAttr = hyperlink is null ? string.Empty : Units.Inv($" data-aio-hyperlink=\"{Escape(hyperlink)}\"");

        svg.Append(Units.Inv($"  <g data-aio-path=\"{Escape(shape.CanonicalPath(slideIndex))}\" data-name=\"{Escape(shape.Name)}\"{linkAttr}>\n"));

        // Accessible name: an SVG <title> child is what assistive tech announces.
        if ((PptxDoc.AltText(shape.Element) ?? PptxDoc.AltTitle(shape.Element)) is { } altText)
        {
            svg.Append(Units.Inv($"    <title>{Escape(altText)}</title>\n"));
        }
        else if (hyperlink is not null)
        {
            // No alt text, but a link: announce the link target as the title.
            svg.Append(Units.Inv($"    <title>link: {Escape(hyperlink)}</title>\n"));
        }

        if (PptxCharts.ChartPartOf(slidePart, shape.Element) is { } chartPart)
        {
            svg.Append(Units.Inv($"    <rect x=\"{x:0.#}\" y=\"{y:0.#}\" width=\"{w:0.#}\" height=\"{h:0.#}\" "));
            svg.Append("fill=\"#ffffff\" stroke=\"#999999\"/>\n");
            var hints = chartPart.ChartSpace is { } cs ? PptxChartPolish.RenderHints(cs) : ("none", false);
            var data = PptxCharts.ReadData(chartPart);

            // A legend strip eats a slice of the frame on its named edge; the plot
            // shrinks into the remainder so the mini-chart and legend never overlap.
            var ph = h;
            var py = y;
            if (hints.Item1 is "bottom")
            {
                ph -= 16;
            }
            else if (hints.Item1 is "top")
            {
                ph -= 16;
                py += 16;
            }

            AppendChart(svg, data, x, py, w, ph, hints.Item2);
            if (hints.Item1 != "none")
            {
                AppendChartLegend(svg, data, hints.Item1, x, y, w, h);
            }

            svg.Append("  </g>\n");
            return;
        }

        if (PptxTables.TableOf(shape.Element) is { } table &&
            PptxTables.IndexOf(slidePart, shape.Element) is { } tableIndex)
        {
            AppendTable(svg, table, slideIndex, tableIndex, x, y, w, h);
            svg.Append("  </g>\n");
            return;
        }

        // A zoom navigation object (slide/section/summary): a thumbnail placeholder, never a slide render.
        if (PptxZoom.ZoomViewOf(slidePart, shape.Element) is { } zoom)
        {
            PptxZoom.AppendPlaceholder(svg, zoom, x, y, w, h);
            svg.Append("  </g>\n");
            return;
        }

        if (PptxSmartArt.DataPartOf(slidePart, shape.Element) is not null)
        {
            // SmartArt is read-only: a labeled placeholder box, never a fake redraw.
            var layout = PptxSmartArt.LayoutName(slidePart, shape.Element);
            svg.Append(Units.Inv($"    <rect x=\"{x:0.#}\" y=\"{y:0.#}\" width=\"{w:0.#}\" height=\"{h:0.#}\" "));
            svg.Append("fill=\"none\" stroke=\"#999999\" stroke-dasharray=\"4 3\"/>\n");
            svg.Append(Units.Inv($"    <text x=\"{x + (w / 2):0.#}\" y=\"{y + (h / 2):0.#}\" font-size=\"12\" "));
            svg.Append(Units.Inv($"text-anchor=\"middle\" fill=\"#666666\">[smartart]{(layout is null ? string.Empty : " " + Escape(layout))}</text>\n"));
            svg.Append("  </g>\n");
            return;
        }

        // An embedded 3D model renders as a labeled placeholder box (a real redraw
        // would need a 3D renderer); the poster, when present, is never rasterized.
        if (shape.Element is P.Picture modelPicture && PptxModels.IsModelPicture(slidePart, modelPicture))
        {
            AppendModel3DPlaceholder(svg, shape.Name, x, y, w, h);
            svg.Append("  </g>\n");
            return;
        }

        // Embedded media (video/audio) renders as a placeholder rect with a play
        // glyph, never a rasterized frame — honest, and clickable via data-aio-path.
        if (shape.Element is P.Picture mediaPicture && PptxMedia.MediaKindOf(mediaPicture) is { } mediaKind)
        {
            AppendMediaPlaceholder(svg, mediaKind, shape.Name, x, y, w, h);
            svg.Append("  </g>\n");
            return;
        }

        // A connector with cxn endpoints draws a line between its two anchor shapes.
        if (shape.Element is P.ConnectionShape connector && PptxConnectors.Endpoints(connector) is { StartId: not null } or { EndId: not null })
        {
            AppendConnector(svg, slidePart, connector, x, y, w, h);
            svg.Append("  </g>\n");
            return;
        }

        // A group renders its children at their absolute coordinates (the group's child
        // space is the identity of its box, the way aioffice writes groups).
        if (shape.Element is P.GroupShape groupShape)
        {
            AppendGroupChildren(svg, slidePart, groupShape, slideIndex);
            svg.Append("  </g>\n");
            return;
        }

        // An action button (navigation): a 3D-style button face with a centered glyph.
        if (shape.Element is P.Shape actionShape &&
            actionShape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset is { } actionPreset &&
            PptxActionButtons.IsActionButtonGeometry(actionPreset))
        {
            AppendActionButton(svg, actionPreset, x, y, w, h, fill);
            svg.Append("  </g>\n");
            return;
        }

        AppendOutline(svg, shape, x, y, w, h, fill);

        if (shape.Kind == "picture")
        {
            // Pictures render as a labeled placeholder; the bytes are never rasterized into the svg.
            svg.Append(Units.Inv($"    <text x=\"{x + (w / 2):0.#}\" y=\"{y + (h / 2):0.#}\" font-size=\"12\" "));
            svg.Append(Units.Inv($"text-anchor=\"middle\" fill=\"#666666\">[image] {Escape(shape.Name)}</text>\n"));
        }

        if (shape.Element is P.Shape { TextBody: { } textBody })
        {
            var cursorY = y + 4;
            foreach (var paragraph in textBody.Elements<A.Paragraph>())
            {
                var text = PptxDoc.ParagraphText(paragraph);
                var runProperties = paragraph.Elements<A.Run>().FirstOrDefault()?.RunProperties;
                var fontPt = runProperties?.FontSize?.Value is { } size ? size / 100.0 : DefaultFontPt;
                var fontPx = fontPt * 4.0 / 3.0;
                var bold = runProperties?.Bold?.Value == true;
                var color = runProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value ?? "111111";

                var alignment = paragraph.ParagraphProperties?.Alignment;
                var (anchor, textX) = alignment is not null && alignment.Value == A.TextAlignmentTypeValues.Center
                    ? ("middle", x + (w / 2))
                    : alignment is not null && alignment.Value == A.TextAlignmentTypeValues.Right
                        ? ("end", x + w - 6)
                        : ("start", x + 6);

                var baseline = cursorY + fontPx;
                if (text.Length > 0)
                {
                    svg.Append(Units.Inv($"    <text x=\"{textX:0.#}\" y=\"{baseline:0.#}\" font-size=\"{fontPx:0.#}\" "));
                    svg.Append(Units.Inv($"text-anchor=\"{anchor}\"{(bold ? " font-weight=\"bold\"" : string.Empty)} "));
                    svg.Append(Units.Inv($"fill=\"#{color}\">{Escape(text.Replace('\n', ' '))}</text>\n"));
                }

                cursorY += fontPx * 1.35;
            }
        }

        svg.Append("  </g>\n");
    }

    /// <summary>
    /// Draws an action button: a rounded "button" face plus the navigation glyph for
    /// its preset geometry (|◄ ◄ ► ►| ⌂ ↩ …). The button reads as clickable and the
    /// surrounding g still carries the data-aio-path + data-aio-hyperlink contract.
    /// </summary>
    private static void AppendActionButton(StringBuilder svg, A.ShapeTypeValues preset, double x, double y, double w, double h, string? fill)
    {
        var face = fill is null ? "e5e7eb" : fill.ToLowerInvariant();
        var radius = Math.Min(w, h) / 8;
        svg.Append(Units.Inv($"    <rect class=\"aio-action-button\" x=\"{x:0.#}\" y=\"{y:0.#}\" width=\"{w:0.#}\" height=\"{h:0.#}\" "));
        svg.Append(Units.Inv($"rx=\"{radius:0.#}\" fill=\"#{face}\" stroke=\"#6b7280\" stroke-width=\"1.5\"/>\n"));

        var glyph = PptxActionButtons.GlyphFor(preset);
        if (glyph is { Length: > 0 })
        {
            var fontPx = Math.Max(Math.Min(w, h) * 0.45, 8);
            svg.Append(Units.Inv($"    <text x=\"{x + (w / 2):0.#}\" y=\"{y + (h / 2) + (fontPx * 0.35):0.#}\" font-size=\"{fontPx:0.#}\" "));
            svg.Append(Units.Inv($"text-anchor=\"middle\" fill=\"#374151\">{Escape(glyph)}</text>\n"));
        }
    }

    /// <summary>
    /// Draws an embedded-media placeholder: a dark rect with a centered play
    /// triangle (video) or speaker glyph (audio) and a label. The bytes are never
    /// decoded — this is an honest stand-in, not a rasterized frame.
    /// </summary>
    private static void AppendMediaPlaceholder(StringBuilder svg, string mediaKind, string name, double x, double y, double w, double h)
    {
        svg.Append(Units.Inv($"    <rect class=\"aio-media\" x=\"{x:0.#}\" y=\"{y:0.#}\" width=\"{w:0.#}\" height=\"{h:0.#}\" "));
        svg.Append("fill=\"#1f2937\" stroke=\"#111111\"/>\n");

        var cx = x + (w / 2);
        var cy = y + (h / 2);
        var glyph = Math.Min(w, h) * 0.18;
        if (glyph < 6)
        {
            glyph = 6;
        }

        // A centered play triangle reads as "media" for both video and audio.
        svg.Append(Units.Inv($"    <polygon class=\"aio-media-play\" points=\"{cx - (glyph * 0.6):0.#},{cy - glyph:0.#} "));
        svg.Append(Units.Inv($"{cx + glyph:0.#},{cy:0.#} {cx - (glyph * 0.6):0.#},{cy + glyph:0.#}\" fill=\"#ffffff\"/>\n"));

        svg.Append(Units.Inv($"    <text x=\"{cx:0.#}\" y=\"{y + h - 6:0.#}\" font-size=\"11\" "));
        svg.Append(Units.Inv($"text-anchor=\"middle\" fill=\"#e5e7eb\">[{Escape(mediaKind)}] {Escape(name)}</text>\n"));
    }

    /// <summary>
    /// Draws an embedded-3D-model placeholder: a dashed rect with a small cube glyph
    /// and a label. The model bytes are never rendered — this is an honest stand-in.
    /// </summary>
    private static void AppendModel3DPlaceholder(StringBuilder svg, string name, double x, double y, double w, double h)
    {
        svg.Append(Units.Inv($"    <rect class=\"aio-model3d\" x=\"{x:0.#}\" y=\"{y:0.#}\" width=\"{w:0.#}\" height=\"{h:0.#}\" "));
        svg.Append("fill=\"#eef2f7\" stroke=\"#64748b\" stroke-dasharray=\"4 3\"/>\n");

        var cx = x + (w / 2);
        var cy = y + (h / 2);
        var s = Math.Max(Math.Min(w, h) * 0.16, 8);

        // A small isometric cube: a front face plus two foreshortened side/top faces.
        var d = s * 0.5;
        svg.Append(Units.Inv($"    <polygon class=\"aio-model3d-cube\" points=\"{cx - s:0.#},{cy - (s / 2):0.#} {cx:0.#},{cy - s:0.#} {cx + s:0.#},{cy - (s / 2):0.#} {cx:0.#},{cy:0.#}\" "));
        svg.Append("fill=\"#cbd5e1\" stroke=\"#64748b\"/>\n");
        svg.Append(Units.Inv($"    <polygon points=\"{cx - s:0.#},{cy - (s / 2):0.#} {cx:0.#},{cy:0.#} {cx:0.#},{cy + d:0.#} {cx - s:0.#},{cy - (s / 2) + d:0.#}\" "));
        svg.Append("fill=\"#94a3b8\" stroke=\"#64748b\"/>\n");
        svg.Append(Units.Inv($"    <polygon points=\"{cx + s:0.#},{cy - (s / 2):0.#} {cx:0.#},{cy:0.#} {cx:0.#},{cy + d:0.#} {cx + s:0.#},{cy - (s / 2) + d:0.#}\" "));
        svg.Append("fill=\"#b6c2d1\" stroke=\"#64748b\"/>\n");

        svg.Append(Units.Inv($"    <text x=\"{cx:0.#}\" y=\"{y + h - 6:0.#}\" font-size=\"11\" "));
        svg.Append(Units.Inv($"text-anchor=\"middle\" fill=\"#475569\">[3d model] {Escape(name)}</text>\n"));
    }

    /// <summary>
    /// Draws a connector as a line between the centers of its two anchor shapes. An elbow
    /// connector (bentConnector3) draws an orthogonal two-segment path; a curved connector
    /// draws a quadratic curve; a straight connector draws a direct line. Falls back to the
    /// connector's own box diagonal when an endpoint shape cannot be located.
    /// </summary>
    private static void AppendConnector(StringBuilder svg, SlidePart slidePart, P.ConnectionShape connector, double x, double y, double w, double h)
    {
        var (startId, endId) = PptxConnectors.Endpoints(connector);
        var shapes = PptxDoc.Shapes(slidePart);

        (double Cx, double Cy)? Center(uint? id)
        {
            if (id is null)
            {
                return null;
            }

            var anchor = shapes.FirstOrDefault(s => s.Id == id.Value);
            if (anchor is null || PptxDoc.Geometry(anchor.Element) is not { } g)
            {
                return null;
            }

            return (Units.EmuToPx(g.X + (g.Cx / 2)), Units.EmuToPx(g.Y + (g.Cy / 2)));
        }

        var from = Center(startId) ?? (x, y);
        var to = Center(endId) ?? (x + w, y + h);

        var stroke = PptxDoc.LineHex(connector)?.ToLowerInvariant() ?? "333333";
        var token = PptxDoc.GeometryToken(connector);

        switch (token)
        {
            case "bentConnector3":
                // An L-shaped elbow: go horizontally to the midpoint x, then vertically.
                var midX = (from.Item1 + to.Item1) / 2;
                svg.Append(Units.Inv($"    <polyline points=\"{from.Item1:0.#},{from.Item2:0.#} {midX:0.#},{from.Item2:0.#} {midX:0.#},{to.Item2:0.#} {to.Item1:0.#},{to.Item2:0.#}\" "));
                svg.Append(Units.Inv($"fill=\"none\" stroke=\"#{stroke}\" stroke-width=\"2\"/>\n"));
                break;
            case "curvedConnector3":
                var ctrlX = (from.Item1 + to.Item1) / 2;
                svg.Append(Units.Inv($"    <path d=\"M {from.Item1:0.#} {from.Item2:0.#} Q {ctrlX:0.#} {from.Item2:0.#} {to.Item1:0.#} {to.Item2:0.#}\" "));
                svg.Append(Units.Inv($"fill=\"none\" stroke=\"#{stroke}\" stroke-width=\"2\"/>\n"));
                break;
            default:
                svg.Append(Units.Inv($"    <line x1=\"{from.Item1:0.#}\" y1=\"{from.Item2:0.#}\" x2=\"{to.Item1:0.#}\" y2=\"{to.Item2:0.#}\" "));
                svg.Append(Units.Inv($"stroke=\"#{stroke}\" stroke-width=\"2\"/>\n"));
                break;
        }
    }

    /// <summary>
    /// Draws a group's children at their absolute coordinates, each wrapped in its own
    /// data-aio-path group (the /slide[i]/group[@id=N]/shape[...] path), so a click maps
    /// back to the addressable child.
    /// </summary>
    private static void AppendGroupChildren(StringBuilder svg, SlidePart slidePart, P.GroupShape groupShape, int slideIndex)
    {
        var groupId = groupShape.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
        foreach (var child in PptxGroups.Children(groupShape))
        {
            var geometry = PptxDoc.Geometry(child.Element) ?? new GeometryEmu(0, 0, Units.CmToEmu(4), Units.CmToEmu(1.5));
            var cx = Units.EmuToPx(geometry.X);
            var cy = Units.EmuToPx(geometry.Y);
            var cw = Units.EmuToPx(geometry.Cx);
            var ch = Units.EmuToPx(geometry.Cy);
            var fill = PptxDoc.FillHex(child.Element);

            var childPath = Units.Inv($"/slide[{slideIndex}]/group[@id={groupId}]/shape[@id={child.Id}]");
            svg.Append(Units.Inv($"    <g data-aio-path=\"{Escape(childPath)}\" data-name=\"{Escape(child.Name)}\">\n"));
            AppendOutline(svg, child, cx, cy, cw, ch, fill);
            if (child.Element is P.Shape { TextBody: { } } && PptxDoc.ShapeText(child.Element) is { Length: > 0 } text)
            {
                svg.Append(Units.Inv($"      <text x=\"{cx + (cw / 2):0.#}\" y=\"{cy + (ch / 2):0.#}\" font-size=\"12\" "));
                svg.Append(Units.Inv($"text-anchor=\"middle\" fill=\"#111111\">{Escape(text.Split('\n')[0])}</text>\n"));
            }

            svg.Append("    </g>\n");
        }
    }

    /// <summary>Draws the shape's outline truthfully per its preset geometry.</summary>
    private static void AppendOutline(StringBuilder svg, ShapeView shape, double x, double y, double w, double h, string? fill)
    {
        var paint = Units.Inv($"fill=\"{(fill is null ? "none" : "#" + fill)}\" stroke=\"#999999\"");
        var token = PptxDoc.GeometryToken(shape.Element);

        switch (token)
        {
            case "ellipse":
                svg.Append(Units.Inv($"    <ellipse cx=\"{x + (w / 2):0.#}\" cy=\"{y + (h / 2):0.#}\" "));
                svg.Append(Units.Inv($"rx=\"{w / 2:0.#}\" ry=\"{h / 2:0.#}\" {paint}/>\n"));
                return;

            case "roundRect":
                var radius = Math.Min(w, h) / 6; // the OOXML default adjust (16667/100000)
                svg.Append(Units.Inv($"    <rect x=\"{x:0.#}\" y=\"{y:0.#}\" width=\"{w:0.#}\" height=\"{h:0.#}\" "));
                svg.Append(Units.Inv($"rx=\"{radius:0.#}\" {paint}/>\n"));
                return;

            case "triangle":
                svg.Append(Units.Inv($"    <polygon points=\"{x + (w / 2):0.#},{y:0.#} {x + w:0.#},{y + h:0.#} {x:0.#},{y + h:0.#}\" {paint}/>\n"));
                return;

            case "diamond":
                svg.Append(Units.Inv($"    <polygon points=\"{x + (w / 2):0.#},{y:0.#} {x + w:0.#},{y + (h / 2):0.#} "));
                svg.Append(Units.Inv($"{x + (w / 2):0.#},{y + h:0.#} {x:0.#},{y + (h / 2):0.#}\" {paint}/>\n"));
                return;

            case "arrow": // rightArrow with default adjusts: head is half the width capped at the height
            {
                var head = Math.Min(w / 2, h);
                var bodyTop = y + (h / 4);
                var bodyBottom = y + (h * 3 / 4);
                var neck = x + w - head;
                svg.Append(Units.Inv($"    <polygon points=\"{x:0.#},{bodyTop:0.#} {neck:0.#},{bodyTop:0.#} {neck:0.#},{y:0.#} "));
                svg.Append(Units.Inv($"{x + w:0.#},{y + (h / 2):0.#} {neck:0.#},{y + h:0.#} {neck:0.#},{bodyBottom:0.#} "));
                svg.Append(Units.Inv($"{x:0.#},{bodyBottom:0.#}\" {paint}/>\n"));
                return;
            }

            case "line":
            {
                // The line spans the box corner-to-corner; flips mirror it inside the box.
                var flip = PptxDoc.FlipToken(shape.Element) ?? string.Empty;
                var flipH = flip.Contains('h', StringComparison.Ordinal);
                var flipV = flip.Contains('v', StringComparison.Ordinal);
                var (x1, x2) = flipH ? (x + w, x) : (x, x + w);
                var (y1, y2) = flipV ? (y + h, y) : (y, y + h);
                var stroke = PptxDoc.LineHex(shape.Element)?.ToLowerInvariant() ?? "333333";
                svg.Append(Units.Inv($"    <line x1=\"{x1:0.#}\" y1=\"{y1:0.#}\" x2=\"{x2:0.#}\" y2=\"{y2:0.#}\" "));
                svg.Append(Units.Inv($"stroke=\"#{stroke}\" stroke-width=\"2\"/>\n"));
                return;
            }

            default:
            {
                var dash = shape.Kind == "shape" ? string.Empty : " stroke-dasharray=\"4 3\"";
                svg.Append(Units.Inv($"    <rect x=\"{x:0.#}\" y=\"{y:0.#}\" width=\"{w:0.#}\" height=\"{h:0.#}\" {paint}{dash}/>\n"));
                return;
            }
        }
    }

    // ----- tables ----------------------------------------------------------------

    /// <summary>
    /// Draws the table grid truthfully: per-cell rects (merged blocks span their
    /// full grid area, covered cells are skipped), cell fills, and text runs —
    /// each visible cell wrapped in a g carrying its data-aio-path.
    /// </summary>
    private static void AppendTable(StringBuilder svg, A.Table table, int slideIndex, int tableIndex, double x, double y, double w, double h)
    {
        var colWidths = (table.TableGrid?.Elements<A.GridColumn>() ?? [])
            .Select(c => Units.EmuToPx(c.Width?.Value ?? 0))
            .ToList();
        var rows = table.Elements<A.TableRow>().ToList();
        var rowHeights = rows.Select(r => Units.EmuToPx(r.Height?.Value ?? 0)).ToList();

        // Scale the grid into the frame box so the drawing stays truthful to it.
        var sumW = colWidths.Sum();
        var sumH = rowHeights.Sum();
        var scaleX = sumW > 0 ? w / sumW : 1;
        var scaleY = sumH > 0 ? h / sumH : 1;

        for (var r = 0; r < rows.Count; r++)
        {
            var cells = rows[r].Elements<A.TableCell>().ToList();
            for (var c = 0; c < cells.Count && c < colWidths.Count; c++)
            {
                var cell = cells[c];
                if (PptxTables.IsCovered(cell))
                {
                    continue;
                }

                var gridSpan = Math.Max(cell.GridSpan?.Value ?? 1, 1);
                var rowSpan = Math.Max(cell.RowSpan?.Value ?? 1, 1);
                var cellX = x + (colWidths.Take(c).Sum() * scaleX);
                var cellY = y + (rowHeights.Take(r).Sum() * scaleY);
                var cellW = colWidths.Skip(c).Take(gridSpan).Sum() * scaleX;
                var cellH = rowHeights.Skip(r).Take(rowSpan).Sum() * scaleY;
                var fill = PptxTables.CellFillHex(cell)?.ToLowerInvariant();

                var path = Units.Inv($"/slide[{slideIndex}]/table[{tableIndex}]/tr[{r + 1}]/tc[{c + 1}]");
                svg.Append(Units.Inv($"    <g data-aio-path=\"{Escape(path)}\">\n"));
                svg.Append(Units.Inv($"      <rect x=\"{cellX:0.#}\" y=\"{cellY:0.#}\" width=\"{cellW:0.#}\" height=\"{cellH:0.#}\" "));
                svg.Append(Units.Inv($"fill=\"{(fill is null ? "none" : "#" + fill)}\" stroke=\"#999999\"/>\n"));
                AppendCellText(svg, cell, cellX, cellY, cellW);
                svg.Append("    </g>\n");
            }
        }
    }

    private static void AppendCellText(StringBuilder svg, A.TableCell cell, double x, double y, double w)
    {
        if (cell.TextBody is not { } body)
        {
            return;
        }

        var cursorY = y + 2;
        foreach (var paragraph in body.Elements<A.Paragraph>())
        {
            var text = PptxDoc.ParagraphText(paragraph);
            var runProperties = paragraph.Elements<A.Run>().FirstOrDefault()?.RunProperties;
            var fontPt = runProperties?.FontSize?.Value is { } size ? size / 100.0 : DefaultFontPt;
            var fontPx = fontPt * 4.0 / 3.0;
            var bold = runProperties?.Bold?.Value == true;
            var color = runProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value ?? "111111";

            var alignment = paragraph.ParagraphProperties?.Alignment;
            var (anchor, textX) = alignment is not null && alignment.Value == A.TextAlignmentTypeValues.Center
                ? ("middle", x + (w / 2))
                : alignment is not null && alignment.Value == A.TextAlignmentTypeValues.Right
                    ? ("end", x + w - 4)
                    : ("start", x + 4);

            var baseline = cursorY + fontPx;
            if (text.Length > 0)
            {
                svg.Append(Units.Inv($"      <text x=\"{textX:0.#}\" y=\"{baseline:0.#}\" font-size=\"{fontPx:0.#}\" "));
                svg.Append(Units.Inv($"text-anchor=\"{anchor}\"{(bold ? " font-weight=\"bold\"" : string.Empty)} "));
                svg.Append(Units.Inv($"fill=\"#{color}\">{Escape(text.Replace('\n', ' '))}</text>\n"));
            }

            cursorY += fontPx * 1.35;
        }
    }

    // ----- mini-charts ---------------------------------------------------------

    /// <summary>
    /// Draws a real (approximate but truthful) mini-chart from the cached data:
    /// axes + bars/lines/wedges + category labels, scaled into the frame box.
    /// When <paramref name="dataLabels"/> is on, bar charts get per-bar value labels.
    /// </summary>
    private static void AppendChart(StringBuilder svg, PptxChartData data, double x, double y, double w, double h, bool dataLabels = false)
    {
        var top = y + 8.0;
        if (data.Title is { } title)
        {
            svg.Append(Units.Inv($"    <text x=\"{x + (w / 2):0.#}\" y=\"{top + 12:0.#}\" font-size=\"13\" "));
            svg.Append(Units.Inv($"text-anchor=\"middle\" font-weight=\"bold\" fill=\"#111111\">{Escape(title)}</text>\n"));
            top += 22;
        }

        var plotX = x + 36;
        var plotY = top + 4;
        var plotW = Math.Max(w - 48, 10);
        var plotH = Math.Max(y + h - 26 - plotY, 10);

        switch (data.Kind)
        {
            case "bar":
                AppendAxes(svg, data, plotX, plotY, plotW, plotH);
                AppendBars(svg, data, plotX, plotY, plotW, plotH);
                if (dataLabels)
                {
                    AppendBarValueLabels(svg, data, plotX, plotY, plotW, plotH);
                }

                AppendCategoryLabels(svg, data, plotX, plotY + plotH, plotW);
                break;
            case "stackedBar":
            case "percentStackedBar":
                AppendAxes(svg, data, plotX, plotY, plotW, plotH, percent: data.Kind == "percentStackedBar");
                AppendStackedBars(svg, data, plotX, plotY, plotW, plotH, percent: data.Kind == "percentStackedBar");
                AppendCategoryLabels(svg, data, plotX, plotY + plotH, plotW);
                break;
            case "line":
                AppendAxes(svg, data, plotX, plotY, plotW, plotH);
                AppendLines(svg, data, plotX, plotY, plotW, plotH);
                AppendCategoryLabels(svg, data, plotX, plotY + plotH, plotW);
                break;
            case "stackedArea":
                AppendAxes(svg, data, plotX, plotY, plotW, plotH);
                AppendStackedAreas(svg, data, plotX, plotY, plotW, plotH);
                AppendCategoryLabels(svg, data, plotX, plotY + plotH, plotW);
                break;
            case "combo":
                // First series as columns, the rest as a line — the same split the chart uses.
                AppendAxes(svg, data, plotX, plotY, plotW, plotH);
                AppendComboColumns(svg, data, plotX, plotY, plotW, plotH);
                AppendComboLine(svg, data, plotX, plotY, plotW, plotH);
                AppendCategoryLabels(svg, data, plotX, plotY + plotH, plotW);
                break;
            case "pie":
                AppendPie(svg, data, x, plotY, w, plotH + 20);
                break;
            case "doughnut":
                AppendDoughnut(svg, data, x, plotY, w, plotH + 20);
                break;
            default:
                // bubble/radar render as an honest labeled placeholder: a real
                // mini-redraw of these would mislead more than it informs.
                svg.Append(Units.Inv($"    <text x=\"{x + (w / 2):0.#}\" y=\"{y + (h / 2):0.#}\" font-size=\"12\" "));
                svg.Append(Units.Inv($"text-anchor=\"middle\" fill=\"#666666\">[chart] {Escape(data.Kind)}</text>\n"));
                break;
        }
    }

    private static double MaxValue(PptxChartData data)
    {
        var max = 0.0;
        foreach (var series in data.Series)
        {
            foreach (var value in series.Values)
            {
                if (value is { } v && v > max)
                {
                    max = v;
                }
            }
        }

        return max > 0 ? max : 1;
    }

    private static void AppendAxes(StringBuilder svg, PptxChartData data, double plotX, double plotY, double plotW, double plotH, bool percent = false)
    {
        svg.Append(Units.Inv($"    <line x1=\"{plotX:0.#}\" y1=\"{plotY:0.#}\" x2=\"{plotX:0.#}\" y2=\"{plotY + plotH:0.#}\" stroke=\"#666666\"/>\n"));
        svg.Append(Units.Inv($"    <line x1=\"{plotX:0.#}\" y1=\"{plotY + plotH:0.#}\" x2=\"{plotX + plotW:0.#}\" y2=\"{plotY + plotH:0.#}\" stroke=\"#666666\"/>\n"));

        var top = percent ? "100%" : Units.Inv($"{MaxValue(data):0.##}");
        svg.Append(Units.Inv($"    <text x=\"{plotX - 4:0.#}\" y=\"{plotY + 4:0.#}\" font-size=\"9\" text-anchor=\"end\" fill=\"#666666\">{top}</text>\n"));
        svg.Append(Units.Inv($"    <text x=\"{plotX - 4:0.#}\" y=\"{plotY + plotH:0.#}\" font-size=\"9\" text-anchor=\"end\" fill=\"#666666\">0</text>\n"));
    }

    /// <summary>The largest per-category column total (used to scale stacked bars/areas).</summary>
    private static double MaxStackTotal(PptxChartData data)
    {
        var max = 0.0;
        for (var c = 0; c < data.Categories.Count; c++)
        {
            var total = 0.0;
            foreach (var series in data.Series)
            {
                if (c < series.Values.Count && series.Values[c] is { } v && v > 0)
                {
                    total += v;
                }
            }

            max = Math.Max(max, total);
        }

        return max > 0 ? max : 1;
    }

    /// <summary>Stacked columns: per category a single stack whose segments are the positive series values.</summary>
    private static void AppendStackedBars(StringBuilder svg, PptxChartData data, double plotX, double plotY, double plotW, double plotH, bool percent)
    {
        var groups = Math.Max(data.Categories.Count, 1);
        var groupW = plotW / groups;
        var barW = groupW * 0.6;
        var globalMax = MaxStackTotal(data);

        for (var c = 0; c < data.Categories.Count; c++)
        {
            var columnTotal = 0.0;
            for (var s = 0; s < data.Series.Count; s++)
            {
                if (c < data.Series[s].Values.Count && data.Series[s].Values[c] is { } v && v > 0)
                {
                    columnTotal += v;
                }
            }

            var scale = percent ? (columnTotal > 0 ? columnTotal : 1) : globalMax;
            var baseline = plotY + plotH;
            for (var s = 0; s < data.Series.Count; s++)
            {
                if (c >= data.Series[s].Values.Count || data.Series[s].Values[c] is not { } value || value <= 0)
                {
                    continue;
                }

                var segH = plotH * value / scale;
                var barX = plotX + (c * groupW) + ((groupW - barW) / 2);
                var color = ChartPalette[s % ChartPalette.Length].ToLowerInvariant();
                svg.Append(Units.Inv($"    <rect class=\"aio-chart-bar\" x=\"{barX:0.#}\" y=\"{baseline - segH:0.#}\" "));
                svg.Append(Units.Inv($"width=\"{barW:0.#}\" height=\"{segH:0.#}\" fill=\"#{color}\"/>\n"));
                baseline -= segH;
            }
        }
    }

    /// <summary>Stacked areas: each series is a band added on top of the running total below it.</summary>
    private static void AppendStackedAreas(StringBuilder svg, PptxChartData data, double plotX, double plotY, double plotW, double plotH)
    {
        var groups = Math.Max(data.Categories.Count, 1);
        var step = groups > 1 ? plotW / (groups - 1) : 0;
        var max = MaxStackTotal(data);
        var lower = new double[groups]; // running cumulative total per category

        for (var s = 0; s < data.Series.Count; s++)
        {
            var color = ChartPalette[s % ChartPalette.Length].ToLowerInvariant();
            var upperPoints = new List<string>();
            var lowerPoints = new List<string>();
            for (var c = 0; c < groups; c++)
            {
                var value = c < data.Series[s].Values.Count && data.Series[s].Values[c] is { } v && v > 0 ? v : 0;
                var px = groups > 1 ? plotX + (c * step) : plotX + (plotW / 2);
                var lowerY = plotY + plotH - (plotH * lower[c] / max);
                lower[c] += value;
                var upperY = plotY + plotH - (plotH * lower[c] / max);
                upperPoints.Add(Units.Inv($"{px:0.#},{upperY:0.#}"));
                lowerPoints.Insert(0, Units.Inv($"{px:0.#},{lowerY:0.#}"));
            }

            svg.Append(Units.Inv($"    <polygon class=\"aio-chart-area\" points=\"{string.Join(' ', upperPoints.Concat(lowerPoints))}\" "));
            svg.Append(Units.Inv($"fill=\"#{color}\" fill-opacity=\"0.7\" stroke=\"#{color}\"/>\n"));
        }
    }

    /// <summary>Combo columns: only the first series, drawn as clustered single columns.</summary>
    private static void AppendComboColumns(StringBuilder svg, PptxChartData data, double plotX, double plotY, double plotW, double plotH)
    {
        if (data.Series.Count == 0)
        {
            return;
        }

        var max = MaxValue(data);
        var groups = Math.Max(data.Categories.Count, 1);
        var groupW = plotW / groups;
        var barW = groupW * 0.5;
        var color = ChartPalette[0].ToLowerInvariant();
        for (var c = 0; c < data.Categories.Count; c++)
        {
            if (c >= data.Series[0].Values.Count || data.Series[0].Values[c] is not { } value || value <= 0)
            {
                continue;
            }

            var barH = plotH * Math.Min(value, max) / max;
            var barX = plotX + (c * groupW) + ((groupW - barW) / 2);
            svg.Append(Units.Inv($"    <rect class=\"aio-chart-bar\" x=\"{barX:0.#}\" y=\"{plotY + plotH - barH:0.#}\" "));
            svg.Append(Units.Inv($"width=\"{barW:0.#}\" height=\"{barH:0.#}\" fill=\"#{color}\"/>\n"));
        }
    }

    /// <summary>Combo line: the second and later series, each a polyline (matching the chart's split).</summary>
    private static void AppendComboLine(StringBuilder svg, PptxChartData data, double plotX, double plotY, double plotW, double plotH)
    {
        var max = MaxValue(data);
        var groups = Math.Max(data.Categories.Count, 1);
        var step = groups > 1 ? plotW / (groups - 1) : 0;

        for (var s = 1; s < data.Series.Count; s++)
        {
            var color = ChartPalette[s % ChartPalette.Length].ToLowerInvariant();
            var points = new List<string>();
            for (var c = 0; c < data.Series[s].Values.Count && c < groups; c++)
            {
                if (data.Series[s].Values[c] is not { } value)
                {
                    continue;
                }

                var px = groups > 1 ? plotX + (c * step) : plotX + (plotW / 2);
                var py = plotY + plotH - (plotH * Math.Clamp(value, 0, max) / max);
                points.Add(Units.Inv($"{px:0.#},{py:0.#}"));
            }

            if (points.Count > 0)
            {
                svg.Append(Units.Inv($"    <polyline class=\"aio-chart-line\" points=\"{string.Join(' ', points)}\" "));
                svg.Append(Units.Inv($"fill=\"none\" stroke=\"#{color}\" stroke-width=\"2\"/>\n"));
            }
        }
    }

    /// <summary>A doughnut: the pie wedges with a punched-out white center hole.</summary>
    private static void AppendDoughnut(StringBuilder svg, PptxChartData data, double x, double top, double w, double h)
    {
        AppendPie(svg, data, x, top, w, h);

        var cx = x + (w / 2);
        var cy = top + (h / 2);
        var radius = Math.Max(Math.Min(w, h) / 2 - 16, 8);
        svg.Append(Units.Inv($"    <circle class=\"aio-chart-hole\" cx=\"{cx:0.#}\" cy=\"{cy:0.#}\" r=\"{radius * 0.5:0.#}\" fill=\"#ffffff\"/>\n"));
    }

    private static void AppendBars(StringBuilder svg, PptxChartData data, double plotX, double plotY, double plotW, double plotH)
    {
        var max = MaxValue(data);
        var groups = data.Categories.Count;
        var seriesCount = Math.Max(data.Series.Count, 1);
        var groupW = plotW / Math.Max(groups, 1);
        var barW = groupW / (seriesCount + 1); // one bar-width of breathing room per group

        for (var s = 0; s < data.Series.Count; s++)
        {
            var color = ChartPalette[s % ChartPalette.Length].ToLowerInvariant();
            for (var c = 0; c < groups; c++)
            {
                if (data.Series[s].Values.Count <= c || data.Series[s].Values[c] is not { } value || value <= 0)
                {
                    continue;
                }

                var barH = plotH * Math.Min(value, max) / max;
                var barX = plotX + (c * groupW) + (barW / 2) + (s * barW);
                svg.Append(Units.Inv($"    <rect class=\"aio-chart-bar\" x=\"{barX:0.#}\" y=\"{plotY + plotH - barH:0.#}\" "));
                svg.Append(Units.Inv($"width=\"{barW * 0.9:0.#}\" height=\"{barH:0.#}\" fill=\"#{color}\"/>\n"));
            }
        }
    }

    /// <summary>Per-bar value labels above each clustered bar (mirrors AppendBars' geometry).</summary>
    private static void AppendBarValueLabels(StringBuilder svg, PptxChartData data, double plotX, double plotY, double plotW, double plotH)
    {
        var max = MaxValue(data);
        var groups = data.Categories.Count;
        var seriesCount = Math.Max(data.Series.Count, 1);
        var groupW = plotW / Math.Max(groups, 1);
        var barW = groupW / (seriesCount + 1);

        for (var s = 0; s < data.Series.Count; s++)
        {
            for (var c = 0; c < groups; c++)
            {
                if (data.Series[s].Values.Count <= c || data.Series[s].Values[c] is not { } value || value <= 0)
                {
                    continue;
                }

                var barH = plotH * Math.Min(value, max) / max;
                var barX = plotX + (c * groupW) + (barW / 2) + (s * barW);
                var labelX = barX + (barW * 0.45);
                var labelY = plotY + plotH - barH - 2;
                svg.Append(Units.Inv($"    <text class=\"aio-chart-label\" x=\"{labelX:0.#}\" y=\"{labelY:0.#}\" font-size=\"8\" "));
                svg.Append(Units.Inv($"text-anchor=\"middle\" fill=\"#333333\">{Escape(FormatNumber(value))}</text>\n"));
            }
        }
    }

    /// <summary>A legend swatch+name strip on the chart's named edge (honest mini-legend for the SVG render).</summary>
    private static void AppendChartLegend(StringBuilder svg, PptxChartData data, string position, double x, double y, double w, double h)
    {
        var names = data.Series.Select(s => s.Name).Where(n => n.Length > 0).ToList();
        if (names.Count == 0)
        {
            return;
        }

        if (position is "right" or "left")
        {
            // A vertical stack of swatch + name down the named side.
            var legendX = position == "right" ? x + w - 56 : x + 6;
            var ly = y + (h / 2) - (names.Count * 7);
            for (var i = 0; i < names.Count; i++)
            {
                var color = ChartPalette[i % ChartPalette.Length].ToLowerInvariant();
                svg.Append(Units.Inv($"    <rect class=\"aio-chart-legend\" x=\"{legendX:0.#}\" y=\"{ly + (i * 14):0.#}\" width=\"8\" height=\"8\" fill=\"#{color}\"/>\n"));
                svg.Append(Units.Inv($"    <text x=\"{legendX + 11:0.#}\" y=\"{ly + (i * 14) + 8:0.#}\" font-size=\"9\" fill=\"#333333\">{Escape(Truncate(names[i], 10))}</text>\n"));
            }

            return;
        }

        // A horizontal swatch row along the top or bottom edge.
        var rowY = position == "top" ? y + 10 : y + h - 6;
        var slot = w / Math.Max(names.Count, 1);
        for (var i = 0; i < names.Count; i++)
        {
            var color = ChartPalette[i % ChartPalette.Length].ToLowerInvariant();
            var sx = x + (i * slot) + 6;
            svg.Append(Units.Inv($"    <rect class=\"aio-chart-legend\" x=\"{sx:0.#}\" y=\"{rowY - 7:0.#}\" width=\"8\" height=\"8\" fill=\"#{color}\"/>\n"));
            svg.Append(Units.Inv($"    <text x=\"{sx + 11:0.#}\" y=\"{rowY:0.#}\" font-size=\"9\" fill=\"#333333\">{Escape(Truncate(names[i], 10))}</text>\n"));
        }
    }

    private static string FormatNumber(double value) =>
        value == Math.Floor(value) && Math.Abs(value) < 1e15
            ? ((long)value).ToString(System.Globalization.CultureInfo.InvariantCulture)
            : value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);

    private static string Truncate(string text, int max) =>
        text.Length <= max ? text : text[..(max - 1)] + "…";

    private static void AppendLines(StringBuilder svg, PptxChartData data, double plotX, double plotY, double plotW, double plotH)
    {
        var max = MaxValue(data);
        var groups = Math.Max(data.Categories.Count, 1);
        var step = groups > 1 ? plotW / (groups - 1) : 0;

        for (var s = 0; s < data.Series.Count; s++)
        {
            var color = ChartPalette[s % ChartPalette.Length].ToLowerInvariant();
            var points = new List<string>();
            for (var c = 0; c < data.Series[s].Values.Count && c < groups; c++)
            {
                if (data.Series[s].Values[c] is not { } value)
                {
                    continue; // gaps interrupt nothing in the approximation; the polyline just skips them
                }

                var px = groups > 1 ? plotX + (c * step) : plotX + (plotW / 2);
                var py = plotY + plotH - (plotH * Math.Clamp(value, 0, max) / max);
                points.Add(Units.Inv($"{px:0.#},{py:0.#}"));
            }

            if (points.Count > 0)
            {
                svg.Append(Units.Inv($"    <polyline class=\"aio-chart-line\" points=\"{string.Join(' ', points)}\" "));
                svg.Append(Units.Inv($"fill=\"none\" stroke=\"#{color}\" stroke-width=\"2\"/>\n"));
            }
        }
    }

    private static void AppendCategoryLabels(StringBuilder svg, PptxChartData data, double plotX, double axisY, double plotW)
    {
        var groups = data.Categories.Count;
        if (groups == 0)
        {
            return;
        }

        var groupW = plotW / groups;
        for (var c = 0; c < groups; c++)
        {
            var cx = plotX + (c * groupW) + (groupW / 2);
            svg.Append(Units.Inv($"    <text x=\"{cx:0.#}\" y=\"{axisY + 14:0.#}\" font-size=\"10\" "));
            svg.Append(Units.Inv($"text-anchor=\"middle\" fill=\"#444444\">{Escape(data.Categories[c])}</text>\n"));
        }
    }

    private static void AppendPie(StringBuilder svg, PptxChartData data, double x, double top, double w, double h)
    {
        var values = data.Series.Count > 0 ? data.Series[0].Values : [];
        var total = values.Sum(v => v is > 0 ? v.Value : 0);
        var cx = x + (w / 2);
        var cy = top + (h / 2);
        var radius = Math.Max(Math.Min(w, h) / 2 - 16, 8);

        if (total <= 0)
        {
            svg.Append(Units.Inv($"    <circle cx=\"{cx:0.#}\" cy=\"{cy:0.#}\" r=\"{radius:0.#}\" fill=\"none\" stroke=\"#999999\"/>\n"));
            return;
        }

        var angle = -Math.PI / 2; // PowerPoint's first slice starts at 12 o'clock
        var slice = 0;
        for (var i = 0; i < values.Count; i++)
        {
            if (values[i] is not { } value || value <= 0)
            {
                continue;
            }

            var color = ChartPalette[slice % ChartPalette.Length].ToLowerInvariant();
            var sweep = 2 * Math.PI * value / total;
            var label = i < data.Categories.Count ? data.Categories[i] : string.Empty;

            if (sweep >= (2 * Math.PI) - 0.0001)
            {
                svg.Append(Units.Inv($"    <circle class=\"aio-chart-slice\" cx=\"{cx:0.#}\" cy=\"{cy:0.#}\" r=\"{radius:0.#}\" fill=\"#{color}\"/>\n"));
            }
            else
            {
                var x1 = cx + (radius * Math.Cos(angle));
                var y1 = cy + (radius * Math.Sin(angle));
                var x2 = cx + (radius * Math.Cos(angle + sweep));
                var y2 = cy + (radius * Math.Sin(angle + sweep));
                var largeArc = sweep > Math.PI ? 1 : 0;
                svg.Append(Units.Inv($"    <path class=\"aio-chart-slice\" d=\"M {cx:0.#} {cy:0.#} L {x1:0.#} {y1:0.#} "));
                svg.Append(Units.Inv($"A {radius:0.#} {radius:0.#} 0 {largeArc} 1 {x2:0.#} {y2:0.#} Z\" fill=\"#{color}\"/>\n"));
            }

            if (label.Length > 0)
            {
                var mid = angle + (sweep / 2);
                var lx = cx + ((radius + 12) * Math.Cos(mid));
                var ly = cy + ((radius + 12) * Math.Sin(mid));
                var anchor = Math.Cos(mid) < -0.1 ? "end" : Math.Cos(mid) > 0.1 ? "start" : "middle";
                svg.Append(Units.Inv($"    <text x=\"{lx:0.#}\" y=\"{ly:0.#}\" font-size=\"10\" "));
                svg.Append(Units.Inv($"text-anchor=\"{anchor}\" fill=\"#444444\">{Escape(label)}</text>\n"));
            }

            angle += sweep;
            slice++;
        }
    }

    public static string WrapHtml(IReadOnlyList<(string Path, string Svg)> slides, double widthPx)
    {
        var html = new StringBuilder();
        html.Append("<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"utf-8\">\n<title>aioffice render</title>\n</head>\n");
        html.Append("<body style=\"background:#f3f4f6;margin:0;padding:24px 0\">\n");
        foreach (var (path, svg) in slides)
        {
            html.Append(Units.Inv($"<div style=\"margin:0 auto 24px;width:{widthPx:0.#}px\" data-path=\"{Escape(path)}\">\n"));
            html.Append(svg);
            html.Append("\n</div>\n");
        }

        html.Append("</body>\n</html>\n");
        return html.ToString();
    }

    internal static string Escape(string text) => text
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal)
        .Replace("\"", "&quot;", StringComparison.Ordinal);
}
