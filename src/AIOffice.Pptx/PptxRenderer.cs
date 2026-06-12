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

    public static string RenderSlideSvg(PresentationPart presentation, SlidePart slidePart, int slideIndex)
    {
        var (width, height) = SlideSizePx(presentation);
        var background = PptxDoc.BackgroundHex(slidePart);
        var svg = new StringBuilder();
        svg.Append(Units.Inv($"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 {width:0.#} {height:0.#}\" "));
        svg.Append(Units.Inv($"width=\"{width:0.#}\" height=\"{height:0.#}\" font-family=\"Helvetica, Arial, sans-serif\" "));
        svg.Append(Units.Inv($"data-slide=\"{slideIndex}\">\n"));
        svg.Append(Units.Inv($"  <rect x=\"0\" y=\"0\" width=\"{width:0.#}\" height=\"{height:0.#}\" "));
        svg.Append(Units.Inv($"fill=\"#{background?.ToLowerInvariant() ?? "ffffff"}\" stroke=\"#cccccc\"/>\n"));

        foreach (var shape in PptxDoc.Shapes(slidePart))
        {
            AppendShape(svg, slidePart, shape, slideIndex);
        }

        svg.Append("</svg>");
        return svg.ToString();
    }

    /// <summary>
    /// Emits one shape wrapped in a group carrying the data-aio-path render
    /// contract: a browser click on any child maps back to the canonical
    /// stable-id document path.
    /// </summary>
    private static void AppendShape(StringBuilder svg, SlidePart slidePart, ShapeView shape, int slideIndex)
    {
        var geometry = PptxDoc.Geometry(shape.Element) ?? new GeometryEmu(0, 0, Units.CmToEmu(4), Units.CmToEmu(1.5));
        var x = Units.EmuToPx(geometry.X);
        var y = Units.EmuToPx(geometry.Y);
        var w = Units.EmuToPx(geometry.Cx);
        var h = Units.EmuToPx(geometry.Cy);
        var fill = PptxDoc.FillHex(shape.Element);

        svg.Append(Units.Inv($"  <g data-aio-path=\"{Escape(shape.CanonicalPath(slideIndex))}\" data-name=\"{Escape(shape.Name)}\">\n"));

        if (PptxCharts.ChartPartOf(slidePart, shape.Element) is { } chartPart)
        {
            svg.Append(Units.Inv($"    <rect x=\"{x:0.#}\" y=\"{y:0.#}\" width=\"{w:0.#}\" height=\"{h:0.#}\" "));
            svg.Append("fill=\"#ffffff\" stroke=\"#999999\"/>\n");
            AppendChart(svg, PptxCharts.ReadData(chartPart), x, y, w, h);
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
    /// </summary>
    private static void AppendChart(StringBuilder svg, PptxChartData data, double x, double y, double w, double h)
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
                AppendCategoryLabels(svg, data, plotX, plotY + plotH, plotW);
                break;
            case "line":
                AppendAxes(svg, data, plotX, plotY, plotW, plotH);
                AppendLines(svg, data, plotX, plotY, plotW, plotH);
                AppendCategoryLabels(svg, data, plotX, plotY + plotH, plotW);
                break;
            case "pie":
                AppendPie(svg, data, x, plotY, w, plotH + 20);
                break;
            default:
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

    private static void AppendAxes(StringBuilder svg, PptxChartData data, double plotX, double plotY, double plotW, double plotH)
    {
        svg.Append(Units.Inv($"    <line x1=\"{plotX:0.#}\" y1=\"{plotY:0.#}\" x2=\"{plotX:0.#}\" y2=\"{plotY + plotH:0.#}\" stroke=\"#666666\"/>\n"));
        svg.Append(Units.Inv($"    <line x1=\"{plotX:0.#}\" y1=\"{plotY + plotH:0.#}\" x2=\"{plotX + plotW:0.#}\" y2=\"{plotY + plotH:0.#}\" stroke=\"#666666\"/>\n"));

        var max = MaxValue(data);
        svg.Append(Units.Inv($"    <text x=\"{plotX - 4:0.#}\" y=\"{plotY + 4:0.#}\" font-size=\"9\" text-anchor=\"end\" fill=\"#666666\">{max:0.##}</text>\n"));
        svg.Append(Units.Inv($"    <text x=\"{plotX - 4:0.#}\" y=\"{plotY + plotH:0.#}\" font-size=\"9\" text-anchor=\"end\" fill=\"#666666\">0</text>\n"));
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
