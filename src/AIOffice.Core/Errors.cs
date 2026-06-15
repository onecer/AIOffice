namespace AIOffice.Core;

/// <summary>
/// The closed set of machine-readable error codes used in every envelope.
/// Codes are snake_case strings (stable wire contract), not a C# enum, so
/// serialization can never drift from the documented surface.
/// </summary>
public static class ErrorCodes
{
    public const string InvalidArgs = "invalid_args";
    public const string FileNotFound = "file_not_found";
    public const string SandboxDenied = "sandbox_denied";

    /// <summary>Address did not resolve. Producers MUST attach nearest-match candidates.</summary>
    public const string InvalidPath = "invalid_path";

    /// <summary>Capability not implemented yet. Suggestion MUST name the workaround.</summary>
    public const string UnsupportedFeature = "unsupported_feature";

    public const string FormatCorrupt = "format_corrupt";
    public const string StaleAddress = "stale_address";

    /// <summary>File exceeds the size guard. Suggestion MUST name AIOFFICE_MAX_FILE_MB.</summary>
    public const string FileTooLarge = "file_too_large";

    /// <summary>Warning-level: a formula could not be evaluated; cached value returned.</summary>
    public const string FormulaNotEvaluated = "formula_not_evaluated";

    /// <summary>
    /// (1.4) A dynamic-array formula's computed result would spill over cells that
    /// already hold content. Nothing was written. Suggestion MUST name clearing
    /// the target range.
    /// </summary>
    public const string SpillBlocked = "spill_blocked";

    /// <summary>Warning-level: a replace op matched nothing (replacements = 0); the edit still succeeds.</summary>
    public const string FindNoMatch = "find_no_match";

    public const string InternalError = "internal_error";

    /// <summary>No live preview server for the file. Suggestion MUST name the open command.</summary>
    public const string PreviewNotRunning = "preview_not_running";

    /// <summary>All known codes, useful for validation and the schema command.</summary>
    public static readonly IReadOnlyList<string> All =
    [
        InvalidArgs, FileNotFound, SandboxDenied, InvalidPath, UnsupportedFeature,
        FormatCorrupt, StaleAddress, FileTooLarge, FormulaNotEvaluated, SpillBlocked, FindNoMatch, InternalError,
        PreviewNotRunning,
    ];
}

/// <summary>
/// The closed set of non-fatal warning codes that ride in <c>meta.warnings</c>.
/// Unlike <see cref="ErrorCodes"/> these never set a non-zero exit code — the
/// command still succeeds. Centralizing them here (rather than as scattered
/// string literals) is the single source of truth: it keeps the wire codes from
/// drifting (e.g. the truncation code is one name everywhere) and lets
/// <c>schema</c> advertise the full vocabulary alongside the error codes.
/// </summary>
public static class WarningCodes
{
    /// <summary>A live preview server was already running for the file; the existing one is reused.</summary>
    public const string AlreadyRunning = "already_running";

    /// <summary>A bibliography's entries/numbers were pre-rendered from the cached sources; Word finalizes format on open/refresh.</summary>
    public const string BibliographyCached = "bibliography_cached";

    /// <summary>Caption/cross-reference numbers were served from the cached SEQ values (Word does not recompute fields headlessly).</summary>
    public const string CaptionNumbersCached = "caption_numbers_cached";

    /// <summary>A cross-format convert dropped content the neutral model cannot carry; see <c>data.dropped</c>.</summary>
    public const string ConvertLossy = "convert_lossy";

    /// <summary>A csv import/export produced no rows (the source range was empty).</summary>
    public const string CsvEmpty = "csv_empty";

    /// <summary>A diff change list exceeded its cap and was truncated; narrow the comparison.</summary>
    public const string DiffTruncated = "diff_truncated";

    /// <summary>A table of figures' entries were pre-rendered from the cached captions; Word repaginates page numbers on open/refresh.</summary>
    public const string FiguresCached = "figures_cached";

    /// <summary>An index's entries were pre-rendered/alphabetized from the marked XE fields; Word computes page numbers on open/refresh.</summary>
    public const string IndexCached = "index_cached";

    /// <summary>(1.5) A table formula's value was computed from a static grid, but one or more input cells are themselves fields; Word may recompute to a different number on refresh (F9).</summary>
    public const string TableFormulaCached = "table_formula_cached";

    /// <summary>An equation carried a LaTeX construct the converter only partially rendered.</summary>
    public const string EquationPartial = "equation_partial";

    /// <summary>A replace op matched nothing (replacements = 0); the edit still succeeded.</summary>
    public const string FindNoMatch = "find_no_match";

    /// <summary>A formula could not be evaluated headlessly; its cached value was returned.</summary>
    public const string FormulaNotEvaluated = "formula_not_evaluated";

    /// <summary>(1.5) A goal-seek solve did not converge; the changing cell was left unchanged.</summary>
    public const string GoalSeekNoSolution = "goal_seek_no_solution";

    /// <summary>A markdown block had no neutral equivalent and was skipped on import.</summary>
    public const string MdBlockSkipped = "md_block_skipped";

    /// <summary>(1.3) A pptx 3D model was embedded as a 3D part behind a poster picture fallback; PowerPoint 2019+ renders the model.</summary>
    public const string Model3DAsMedia = "model3d_as_media";

    /// <summary>Raw HTML in a markdown source was skipped (not rendered into the document).</summary>
    public const string MdHtmlSkipped = "md_html_skipped";

    /// <summary>An image reference in a markdown source was skipped (not embedded).</summary>
    public const string MdImageSkipped = "md_image_skipped";

    /// <summary>A link in a markdown source was flattened to plain text.</summary>
    public const string MdLinkSkipped = "md_link_skipped";

    /// <summary>A response payload exceeded its byte cap and was cut off; narrow the scope or raise --max-bytes.</summary>
    public const string ResultTruncated = "result_truncated";

    /// <summary>A render/png scope was defaulted (e.g. a multi-slide deck rendered slide 1).</summary>
    public const string ScopeDefaulted = "scope_defaulted";

    /// <summary>A view that streams (stats/text) fell back to loading the whole workbook.</summary>
    public const string StreamFallback = "stream_fallback";

    /// <summary>One or more {{placeholders}} in a template were left unresolved by the merge map.</summary>
    public const string TemplateUnresolved = "template_unresolved";

    /// <summary>A table-of-contents page count is unknown headlessly (Word repaginates on open).</summary>
    public const string TocPagesUnknown = "toc_pages_unknown";

    /// <summary>A table of contents was refreshed from the current headings.</summary>
    public const string TocRefreshed = "toc_refreshed";

    /// <summary>The complete, closed set of warning codes, for the schema command.</summary>
    public static readonly IReadOnlyList<string> All =
    [
        AlreadyRunning, BibliographyCached, CaptionNumbersCached, ConvertLossy, CsvEmpty, DiffTruncated,
        EquationPartial, FiguresCached, FindNoMatch, FormulaNotEvaluated, GoalSeekNoSolution, IndexCached,
        MdBlockSkipped, MdHtmlSkipped, MdImageSkipped, MdLinkSkipped, Model3DAsMedia, ResultTruncated,
        ScopeDefaulted, StreamFallback, TableFormulaCached, TemplateUnresolved, TocPagesUnknown, TocRefreshed,
    ];
}

/// <summary>Process exit codes for the CLI.</summary>
public static class ExitCodes
{
    public const int Ok = 0;
    public const int UserError = 2;
    public const int InternalError = 3;
    public const int SandboxDenied = 4;
    public const int UnsupportedFeature = 5;

    /// <summary>Maps an envelope error code to the documented process exit code.</summary>
    public static int ForErrorCode(string errorCode) => errorCode switch
    {
        ErrorCodes.SandboxDenied => SandboxDenied,
        ErrorCodes.UnsupportedFeature => UnsupportedFeature,
        ErrorCodes.FormatCorrupt or ErrorCodes.InternalError => InternalError,
        _ => UserError,
    };
}

/// <summary>
/// The one exception type that crosses command boundaries. A non-empty,
/// actionable <see cref="Suggestion"/> is REQUIRED by construction — an error
/// without a way forward is a bug.
/// </summary>
public sealed class AiofficeException : Exception
{
    public AiofficeException(
        string code,
        string message,
        string suggestion,
        IReadOnlyList<string>? candidates = null,
        Exception? innerException = null)
        : base(message, innerException)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(code);
        ArgumentException.ThrowIfNullOrWhiteSpace(message);
        ArgumentException.ThrowIfNullOrWhiteSpace(suggestion);
        Code = code;
        Suggestion = suggestion;
        Candidates = candidates;
    }

    /// <summary>One of <see cref="ErrorCodes"/>.</summary>
    public string Code { get; }

    /// <summary>Actionable next step for the caller. Never empty.</summary>
    public string Suggestion { get; }

    /// <summary>Nearest-match alternatives (required for invalid_path resolution failures).</summary>
    public IReadOnlyList<string>? Candidates { get; }

    public ErrorBody ToErrorBody() => new(Code, Message, Suggestion, Candidates);
}
