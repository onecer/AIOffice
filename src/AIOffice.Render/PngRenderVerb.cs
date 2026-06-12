using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;

namespace AIOffice.Render;

/// <summary>
/// The cross-format <c>render --to png</c> orchestration shared by the CLI and
/// the MCP server: ask the format handler for its inspectable artifact (html
/// for docx/xlsx, svg for pptx) and screenshot it with a headless Chromium via
/// <see cref="PngRenderer"/>. PNG is binary, so the result is always written
/// to a file (default: next to the source, extension swapped to .png) and the
/// envelope carries the path instead of inlined content.
/// </summary>
public static class PngRenderVerb
{
    /// <summary>Default viewport width for paged (docx/xlsx) renders.</summary>
    public const int DefaultWidthPx = 1280;

    /// <summary>
    /// Renders <c>ctx.File</c> to PNG. Reads <c>ctx.Args.scope</c> (optional)
    /// and <c>ctx.Args.output</c> (optional absolute path, pre-resolved by the
    /// caller's workspace). For pptx without a scope, only the first slide is
    /// rendered and a <c>scope_defaulted</c> warning names the fix.
    /// </summary>
    public static Envelope Execute(IFormatHandler handler, CommandContext ctx)
    {
        var file = ctx.File ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "render --to png needs a document file.",
            "Pass the document path, e.g. 'aioffice render report.docx --to png'.");

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

        var outPng = Str(ctx.Args, "output") ?? Path.ChangeExtension(file, ".png");
        var warnings = new List<Warning>();
        string written;
        if (handler.Kind == DocumentKind.Pptx)
        {
            written = RenderPptxSlide(rendered, scope, outPng, warnings, out scope);
        }
        else
        {
            written = RenderHtmlContent(rendered, outPng);
        }

        var meta = rendered.Meta with
        {
            Warnings = warnings.Count > 0
                ? [.. rendered.Meta.Warnings ?? [], .. warnings]
                : rendered.Meta.Warnings,
        };

        return Envelope.Ok(
            new
            {
                format = "png",
                scope,
                written,
                sizeBytes = new FileInfo(written).Length,
            },
            meta);
    }

    /// <summary>docx/xlsx: wrap the handler's HTML fragment in a minimal UTF-8 page and screenshot it.</summary>
    private static string RenderHtmlContent(Envelope rendered, string outPng)
    {
        var content = DataNode(rendered)?["content"]?.GetValue<string>();
        if (string.IsNullOrEmpty(content))
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                "The format handler returned no HTML content to rasterize.",
                "Re-run with --to html to inspect the intermediate output; report this if it persists.");
        }

        var page =
            "<!DOCTYPE html><html><head><meta charset=\"utf-8\"><style>" +
            "body{margin:16px;background:#ffffff;font-family:-apple-system,'Segoe UI',Helvetica,Arial,sans-serif;color:#111}" +
            "table{border-collapse:collapse;margin:8px 0}td,th{border:1px solid #ccc;padding:2px 8px;min-width:2em}" +
            "header,footer{color:#555;border-bottom:1px solid #ddd;margin-bottom:8px;padding-bottom:4px}" +
            "footer{border-bottom:none;border-top:1px solid #ddd;margin-top:8px;padding-top:4px}" +
            "</style></head><body>" + content + "</body></html>";
        return PngRenderer.HtmlStringToPng(page, outPng, DefaultWidthPx);
    }

    /// <summary>pptx: screenshot one slide's SVG at its natural pixel size.</summary>
    private static string RenderPptxSlide(
        Envelope rendered, string? scope, string outPng, List<Warning> warnings, out string slidePath)
    {
        var data = DataNode(rendered);
        var slides = data?["slides"]?.AsArray();
        if (slides is null || slides.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                "The pptx handler returned no slides to rasterize.",
                "Re-run with --to svg to inspect the intermediate output; report this if it persists.");
        }

        if (scope is null && slides.Count > 1)
        {
            warnings.Add(new Warning(
                "scope_defaulted",
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"The deck has {slides.Count} slides; only /slide[1] was rendered. Pass --scope /slide[N] for another slide.")));
        }

        var first = slides[0]!;
        slidePath = first["path"]!.GetValue<string>();
        var svg = first["svg"]!.GetValue<string>();
        var (width, height) = SvgPixelSize(svg);
        return PngRenderer.SvgToPng(svg, outPng, width, height);
    }

    /// <summary>Reads width/height off the svg root so the viewport matches the slide exactly.</summary>
    internal static (int WidthPx, int HeightPx) SvgPixelSize(string svg)
    {
        return (Dimension("width", 1280), Dimension("height", 720));

        int Dimension(string name, int fallback)
        {
            var marker = name + "=\"";
            var at = svg.IndexOf(marker, StringComparison.Ordinal);
            if (at < 0)
            {
                return fallback;
            }

            var start = at + marker.Length;
            var end = svg.IndexOf('"', start);
            return end > start &&
                   double.TryParse(svg[start..end], NumberStyles.Float, CultureInfo.InvariantCulture, out var value) &&
                   value is > 0 and < 100_000
                ? (int)Math.Ceiling(value)
                : fallback;
        }
    }

    private static JsonObject? DataNode(Envelope envelope) =>
        JsonSerializer.SerializeToNode(envelope.Data, JsonDefaults.Options) as JsonObject;

    private static string? Str(JsonObject args, string name) =>
        args[name] is JsonValue v && v.TryGetValue<string>(out var s) ? s : null;
}
