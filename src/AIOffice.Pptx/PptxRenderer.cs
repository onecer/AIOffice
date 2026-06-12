using System.Text;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// Approximate-but-inspectable rendering: each slide becomes an SVG of
/// positioned rectangles and text runs (the pptx render-look-fix MVP).
/// </summary>
internal static class PptxRenderer
{
    private const double DefaultFontPt = 18;

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
        var svg = new StringBuilder();
        svg.Append(Units.Inv($"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 {width:0.#} {height:0.#}\" "));
        svg.Append(Units.Inv($"width=\"{width:0.#}\" height=\"{height:0.#}\" font-family=\"Helvetica, Arial, sans-serif\" "));
        svg.Append(Units.Inv($"data-slide=\"{slideIndex}\">\n"));
        svg.Append(Units.Inv($"  <rect x=\"0\" y=\"0\" width=\"{width:0.#}\" height=\"{height:0.#}\" fill=\"#ffffff\" stroke=\"#cccccc\"/>\n"));

        foreach (var shape in PptxDoc.Shapes(slidePart))
        {
            AppendShape(svg, shape, slideIndex);
        }

        svg.Append("</svg>");
        return svg.ToString();
    }

    /// <summary>
    /// Emits one shape wrapped in a group carrying the data-aio-path render
    /// contract: a browser click on any child maps back to the canonical
    /// stable-id document path.
    /// </summary>
    private static void AppendShape(StringBuilder svg, ShapeView shape, int slideIndex)
    {
        var geometry = PptxDoc.Geometry(shape.Element) ?? new GeometryEmu(0, 0, Units.CmToEmu(4), Units.CmToEmu(1.5));
        var x = Units.EmuToPx(geometry.X);
        var y = Units.EmuToPx(geometry.Y);
        var w = Units.EmuToPx(geometry.Cx);
        var h = Units.EmuToPx(geometry.Cy);
        var fill = PptxDoc.FillHex(shape.Element);
        var dash = shape.Kind == "shape" ? string.Empty : " stroke-dasharray=\"4 3\"";

        svg.Append(Units.Inv($"  <g data-aio-path=\"{Escape(shape.CanonicalPath(slideIndex))}\" data-name=\"{Escape(shape.Name)}\">\n"));
        svg.Append(Units.Inv($"    <rect x=\"{x:0.#}\" y=\"{y:0.#}\" width=\"{w:0.#}\" height=\"{h:0.#}\" "));
        svg.Append(Units.Inv($"fill=\"{(fill is null ? "none" : "#" + fill)}\" stroke=\"#999999\"{dash}/>\n"));

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
