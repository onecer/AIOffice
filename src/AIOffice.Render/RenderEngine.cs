using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;

namespace AIOffice.Render;

/// <summary>The render engine a caller asked for via <c>render --engine</c>.</summary>
public enum RenderEngine
{
    /// <summary>The default: screenshot/print aioffice's own HTML/SVG projection with a headless Chromium. 100% unchanged.</summary>
    Chromium,

    /// <summary>TRUE-fidelity: hand the original document to a headless LibreOffice (png+pdf only).</summary>
    Soffice,

    /// <summary>soffice when the doctor finds it on this machine, else chromium.</summary>
    Auto,
}

/// <summary>
/// Shared helpers for the optional <c>--engine</c> render selector. Parses the
/// flag, resolves <c>auto</c> against the live soffice probe, and pulls a
/// 1-based page index out of a <c>--scope</c> like <c>/slide[3]</c> (the soffice
/// engine converts the whole document, so it needs the page number rather than
/// a pre-scoped fragment).
/// </summary>
public static class RenderEngineSelector
{
    /// <summary>The accepted <c>--engine</c> values, for schema/help and the invalid-value error.</summary>
    public static readonly IReadOnlyList<string> Names = ["chromium", "soffice", "auto"];

    private static readonly Regex SlideScope =
        new(@"/slide\[(\d+)\]", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>
    /// Parses the optional <c>--engine</c> value (null/empty/"chromium" → the
    /// default chromium engine). An unknown value is <c>invalid_args</c>.
    /// </summary>
    public static RenderEngine Parse(string? engine)
    {
        if (string.IsNullOrWhiteSpace(engine))
        {
            return RenderEngine.Chromium;
        }

        return engine.ToLowerInvariant() switch
        {
            "chromium" => RenderEngine.Chromium,
            "soffice" => RenderEngine.Soffice,
            "auto" => RenderEngine.Auto,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown render engine: '{engine}'.",
                "Use --engine chromium (default), soffice (LibreOffice true-fidelity), or auto.",
                candidates: Names),
        };
    }

    /// <summary>
    /// Resolves <c>auto</c> to a concrete engine against the live soffice probe:
    /// soffice when LibreOffice is present, else chromium. chromium/soffice pass
    /// through unchanged (the explicit soffice error path is handled downstream).
    /// </summary>
    public static RenderEngine Resolve(RenderEngine requested, SofficeInfo? soffice = null)
    {
        if (requested != RenderEngine.Auto)
        {
            return requested;
        }

        var info = soffice ?? SofficeLocator.Probe();
        return info.Found ? RenderEngine.Soffice : RenderEngine.Chromium;
    }

    /// <summary>
    /// The 1-based page/slide number to rasterize for the soffice PNG path.
    /// A scope of <c>/slide[N]</c> selects slide N; anything else (or no scope)
    /// is page 1. The soffice engine converts the whole document, so the page
    /// number — not a pre-scoped fragment — is what selects the output page.
    /// </summary>
    public static int PageFromScope(string? scope)
    {
        if (string.IsNullOrEmpty(scope))
        {
            return 1;
        }

        var match = SlideScope.Match(scope);
        return match.Success &&
               int.TryParse(match.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var n) &&
               n >= 1
            ? n
            : 1;
    }
}
