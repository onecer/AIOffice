using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;

namespace AIOffice.Render;

/// <summary>
/// The cross-format <c>render --to pdf</c> orchestration shared by the CLI and
/// the MCP server (M3): ask the format handler for its inspectable artifact
/// (html for docx/xlsx, svg for pptx) and print it to a paged PDF with a
/// headless Chromium via <see cref="PdfRenderer"/>. For pptx the WHOLE deck
/// becomes one PDF — every slide is its own page (print CSS pins the page
/// size to the slide size); <c>--scope /slide[N]</c> narrows to one slide.
/// PDF is binary, so the result is always written to a file (default: next to
/// the source, extension swapped to .pdf) and the envelope carries the path.
/// </summary>
public static class PdfRenderVerb
{
    /// <summary>Renders <c>ctx.File</c> to PDF. Reads <c>ctx.Args.scope</c> and <c>ctx.Args.output</c> (pre-resolved).</summary>
    public static Envelope Execute(IFormatHandler handler, CommandContext ctx)
    {
        var file = ctx.File ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "render --to pdf needs a document file.",
            "Pass the document path, e.g. 'aioffice render report.docx --to pdf'.");

        var scope = Str(ctx.Args, "scope");
        var intermediate = handler.Kind == DocumentKind.Pptx ? "svg" : "html";
        var inner = new CommandContext
        {
            Workspace = ctx.Workspace,
            File = file,
            Args = new JsonObject { ["to"] = intermediate, ["scope"] = scope },
        };

        var rendered = handler.Render(inner);
        if (!rendered.IsOk)
        {
            return rendered; // typed failure (invalid scope, corrupt file, ...) passes through untouched
        }

        var outPdf = Str(ctx.Args, "output") ?? Path.ChangeExtension(file, ".pdf");
        string written;
        int pages;
        if (handler.Kind == DocumentKind.Pptx)
        {
            written = PrintDeck(rendered, outPdf, out pages);
        }
        else
        {
            written = PrintHtmlContent(rendered, outPdf);
            pages = 0; // unknown until the browser paginates; 0 = not reported
        }

        return Envelope.Ok(
            new
            {
                format = "pdf",
                scope,
                written,
                sizeBytes = new FileInfo(written).Length,
                pages = pages > 0 ? (int?)pages : null,
            },
            rendered.Meta);
    }

    /// <summary>docx/xlsx: wrap the handler's HTML fragment in a printable A4 page and print it.</summary>
    private static string PrintHtmlContent(Envelope rendered, string outPdf)
    {
        var content = DataNode(rendered)?["content"]?.GetValue<string>();
        if (string.IsNullOrEmpty(content))
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                "The format handler returned no HTML content to print.",
                "Re-run with --to html to inspect the intermediate output; report this if it persists.");
        }

        var page =
            "<!DOCTYPE html><html><head><meta charset=\"utf-8\"><style>" +
            "@page{size:A4;margin:14mm}" +
            "body{margin:0;background:#ffffff;font-family:-apple-system,'Segoe UI',Helvetica,Arial,sans-serif;color:#111}" +
            "table{border-collapse:collapse;margin:8px 0}td,th{border:1px solid #ccc;padding:2px 8px;min-width:2em}" +
            "header,footer{color:#555;border-bottom:1px solid #ddd;margin-bottom:8px;padding-bottom:4px}" +
            "footer{border-bottom:none;border-top:1px solid #ddd;margin-top:8px;padding-top:4px}" +
            "</style></head><body>" + content + "</body></html>";
        return PdfRenderer.HtmlStringToPdf(page, outPdf);
    }

    /// <summary>
    /// pptx: one PDF page per rendered slide. The print CSS pins the @page box
    /// to the slide's pixel size, so each SVG fills its page exactly.
    /// </summary>
    private static string PrintDeck(Envelope rendered, string outPdf, out int pages)
    {
        var data = DataNode(rendered);
        var slides = data?["slides"]?.AsArray();
        if (slides is null || slides.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                "The pptx handler returned no slides to print.",
                "Re-run with --to svg to inspect the intermediate output; report this if it persists.");
        }

        var firstSvg = slides[0]!["svg"]!.GetValue<string>();
        var (width, height) = PngRenderVerb.SvgPixelSize(firstSvg);

        var html = new StringBuilder();
        html.Append("<!DOCTYPE html><html><head><meta charset=\"utf-8\"><style>");
        html.Append(CultureInfo.InvariantCulture, $"@page{{size:{width}px {height}px;margin:0}}");
        html.Append("html,body{margin:0;padding:0}");
        html.Append(CultureInfo.InvariantCulture, $".slide{{width:{width}px;height:{height}px;overflow:hidden;break-after:page}}");
        html.Append(".slide:last-child{break-after:auto}svg{display:block}");
        html.Append("</style></head><body>");
        foreach (var slide in slides)
        {
            html.Append("<div class=\"slide\">");
            html.Append(slide!["svg"]!.GetValue<string>());
            html.Append("</div>");
        }

        html.Append("</body></html>");

        pages = slides.Count;
        return PdfRenderer.HtmlStringToPdf(html.ToString(), outPdf);
    }

    private static JsonObject? DataNode(Envelope envelope) =>
        JsonSerializer.SerializeToNode(envelope.Data, JsonDefaults.Options) as JsonObject;

    private static string? Str(JsonObject args, string name) =>
        args[name] is JsonValue v && v.TryGetValue<string>(out var s) ? s : null;
}
