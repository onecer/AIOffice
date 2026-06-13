using System.Text.Json;
using System.Text.Json.Serialization;

namespace AIOffice.Core;

/// <summary>A non-fatal warning attached to envelope meta.</summary>
public sealed record Warning(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("message")] string Message);

/// <summary>The error half of an envelope. Suggestion is always present.</summary>
public sealed record ErrorBody(
    [property: JsonPropertyName("code")] string Code,
    [property: JsonPropertyName("message")] string Message,
    [property: JsonPropertyName("suggestion")] string Suggestion,
    [property: JsonPropertyName("candidates")] IReadOnlyList<string>? Candidates = null);

/// <summary>Metadata attached to every envelope.</summary>
public sealed record Meta
{
    [JsonPropertyName("file")]
    public string? File { get; init; }

    /// <summary>First 12 hex chars of the SHA-256 of the file bytes.</summary>
    [JsonPropertyName("rev")]
    public string? Rev { get; init; }

    [JsonPropertyName("elapsedMs")]
    public long ElapsedMs { get; init; }

    [JsonPropertyName("version")]
    public string Version { get; init; } = ToolVersion;

    [JsonPropertyName("warnings")]
    public IReadOnlyList<Warning>? Warnings { get; init; }

    /// <summary>The aioffice tool version reported in every envelope.</summary>
    public static string ToolVersion { get; } =
        typeof(Meta).Assembly.GetName().Version is { } v ? $"{v.Major}.{v.Minor}.{v.Build}" : "0.0.0";

    /// <summary>
    /// The AI-facing CONTRACT version — the stability promise for the JSON
    /// envelope shape, the ErrorCode list, the addressing grammar, exit codes, and
    /// the op/view/tool vocabularies (see CONTRACT.md). It is independent of the
    /// tool's package version (<see cref="ToolVersion"/>): the tool can ship many
    /// releases while the contract stays "1.0-rc". Bumps to this value signal a
    /// breaking change to the AI surface. Surfaced in office_schema and doctor.
    /// </summary>
    public const string SurfaceVersion = "1.0-rc";
}

/// <summary>
/// The single response shape every command prints to stdout: exactly one JSON
/// object with ok/data/error/meta. Success has data and a null error; failure
/// has error (always with a suggestion) and null data.
/// </summary>
public sealed record Envelope
{
    private static readonly JsonSerializerOptions CompactJson = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false,
    };

    private static readonly JsonSerializerOptions PrettyJson = new(CompactJson)
    {
        WriteIndented = true,
    };

    [JsonPropertyName("ok")]
    public required bool IsOk { get; init; }

    [JsonPropertyName("data")]
    public object? Data { get; init; }

    [JsonPropertyName("error")]
    public ErrorBody? Error { get; init; }

    [JsonPropertyName("meta")]
    public required Meta Meta { get; init; }

    /// <summary>Builds a success envelope.</summary>
    public static Envelope Ok(object? data, Meta? meta = null) =>
        new() { IsOk = true, Data = data, Error = null, Meta = meta ?? new Meta() };

    /// <summary>Builds a failure envelope. Suggestion is mandatory by contract.</summary>
    public static Envelope Fail(
        string code,
        string message,
        string suggestion,
        IReadOnlyList<string>? candidates = null,
        Meta? meta = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(suggestion);
        return new Envelope
        {
            IsOk = false,
            Data = null,
            Error = new ErrorBody(code, message, suggestion, candidates),
            Meta = meta ?? new Meta(),
        };
    }

    /// <summary>Builds a failure envelope from any exception, never leaking a crash.</summary>
    public static Envelope FromException(Exception exception, Meta? meta = null) => exception switch
    {
        AiofficeException ax => new Envelope
        {
            IsOk = false,
            Data = null,
            Error = ax.ToErrorBody(),
            Meta = meta ?? new Meta(),
        },
        _ => Fail(
            ErrorCodes.InternalError,
            exception.Message,
            "This is a bug in aioffice. Re-run with the same input and report the issue with this message.",
            meta: meta),
    };

    /// <summary>The exit code the CLI should use for this envelope. Not part of the JSON wire shape.</summary>
    [JsonIgnore]
    public int ExitCode => IsOk ? ExitCodes.Ok : ExitCodes.ForErrorCode(Error!.Code);

    /// <summary>Serializes the envelope to its canonical JSON wire form.</summary>
    public string ToJson(bool pretty = false) =>
        JsonSerializer.Serialize(this, pretty ? PrettyJson : CompactJson);
}
