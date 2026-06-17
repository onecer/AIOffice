using AIOffice.Core;

namespace AIOffice.Render;

/// <summary>
/// The single render entry point shared by the CLI <c>render</c> verb and the
/// MCP <c>office_render</c> tool, so the optional <c>--engine</c> selector
/// behaves identically on both surfaces. It owns the cross-format png/pdf
/// plumbing and the engine-aware fallback for svg/html/text:
/// <list type="bullet">
/// <item>png/pdf → <see cref="PngRenderVerb"/> / <see cref="PdfRenderVerb"/>
/// (which read <c>ctx.Args.engine</c> themselves and pick chromium or soffice).</item>
/// <item>svg/html/text → the native <see cref="IFormatHandler.Render"/>. The
/// soffice engine supports only png+pdf, so an explicit (or auto-resolved)
/// soffice request for these targets falls back to the native engine with an
/// <c>engine_fallback</c> warning.</item>
/// </list>
/// The <c>--engine</c> value is OPTIONAL and additive: omitting it (or passing
/// <c>chromium</c>) leaves every existing render path byte-for-byte unchanged.
/// </summary>
public static class RenderDispatch
{
    /// <summary>The <c>engine_fallback</c> warning code (soffice → native for svg/html/text).</summary>
    public const string EngineFallbackCode = "engine_fallback";

    /// <summary>
    /// Runs the render. <paramref name="to"/> is the (already-normalized)
    /// <c>--to</c> value; null/empty defaults to the handler's own default
    /// (html). The engine is read from <c>ctx.Args["engine"]</c>.
    /// </summary>
    public static Envelope Execute(IFormatHandler handler, CommandContext ctx, string? to)
    {
        if (to is "png")
        {
            return PngRenderVerb.Execute(handler, ctx);
        }

        if (to is "pdf")
        {
            return PdfRenderVerb.Execute(handler, ctx);
        }

        // svg / html / text (and the default): native handler. Validate the
        // engine value up-front so a bad --engine is invalid_args even here,
        // then fall back with a warning when soffice was requested.
        var engine = RenderEngineSelector.Parse(EngineArg(ctx));
        var resolved = RenderEngineSelector.Resolve(engine);
        var native = handler.Render(ctx);
        if (resolved != RenderEngine.Soffice || !native.IsOk)
        {
            return native;
        }

        var warning = new Warning(
            EngineFallbackCode,
            $"The soffice engine renders png+pdf only; '{to ?? "html"}' was produced by the native engine. " +
            "Use --to png|pdf for soffice fidelity.");
        var meta = native.Meta with
        {
            Warnings = [.. native.Meta.Warnings ?? [], warning],
        };
        return native with { Meta = meta };
    }

    private static string? EngineArg(CommandContext ctx) =>
        ctx.Args["engine"] is System.Text.Json.Nodes.JsonValue v && v.TryGetValue<string>(out var s) ? s : null;
}
