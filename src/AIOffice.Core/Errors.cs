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

    /// <summary>Warning-level: a formula could not be evaluated; cached value returned.</summary>
    public const string FormulaNotEvaluated = "formula_not_evaluated";

    public const string InternalError = "internal_error";

    /// <summary>No live preview server for the file. Suggestion MUST name the open command.</summary>
    public const string PreviewNotRunning = "preview_not_running";

    /// <summary>All known codes, useful for validation and the schema command.</summary>
    public static readonly IReadOnlyList<string> All =
    [
        InvalidArgs, FileNotFound, SandboxDenied, InvalidPath, UnsupportedFeature,
        FormatCorrupt, StaleAddress, FormulaNotEvaluated, InternalError, PreviewNotRunning,
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
