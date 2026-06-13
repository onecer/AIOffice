using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace AIOffice.Core;

/// <summary>The document formats aioffice understands.</summary>
public enum DocumentKind
{
    Docx,
    Xlsx,
    Pptx,
}

/// <summary>
/// One operation inside an atomic <c>edit --ops</c> batch:
/// <c>[{op:set|add|remove|move, path, type?, props?, position?}]</c>.
/// </summary>
public sealed record EditOp
{
    public static readonly IReadOnlyList<string> Kinds = ["set", "add", "remove", "move", "replace", "accept", "reject"];

    /// <summary>set | add | remove | move.</summary>
    [JsonPropertyName("op")]
    public required string Op { get; init; }

    /// <summary>Target address (addressing grammar).</summary>
    [JsonPropertyName("path")]
    public required string Path { get; init; }

    /// <summary>Element type for add (e.g. p, table, slide, shape).</summary>
    [JsonPropertyName("type")]
    public string? Type { get; init; }

    /// <summary>Properties to set/add, as a free-form JSON object.</summary>
    [JsonPropertyName("props")]
    public JsonObject? Props { get; init; }

    /// <summary>Placement for add/move: before | after | inside | a target path.</summary>
    [JsonPropertyName("position")]
    public string? Position { get; init; }

    /// <summary>Parses a JSON array of edit ops, validating op kinds and paths.</summary>
    public static IReadOnlyList<EditOp> ParseBatch(string json)
    {
        List<EditOp>? ops;
        try
        {
            ops = JsonSerializer.Deserialize<List<EditOp>>(json, JsonDefaults.Options);
        }
        catch (JsonException ex)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"--ops is not valid JSON: {ex.Message}",
                "Pass a JSON array like [{\"op\":\"set\",\"path\":\"/body/p[1]\",\"props\":{\"text\":\"Hi\"}}] or @ops.json.",
                innerException: ex);
        }

        if (ops is null || ops.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "--ops must be a non-empty JSON array of operations.",
                "Pass at least one operation, e.g. [{\"op\":\"set\",\"path\":\"/body/p[1]\",\"props\":{\"text\":\"Hi\"}}].");
        }

        for (var i = 0; i < ops.Count; i++)
        {
            var op = ops[i];
            if (!Kinds.Contains(op.Op, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{i}].op is '{op.Op}' but must be one of: set, add, remove, move, replace, accept, reject.",
                    "Use set to change properties, add to insert, remove to delete, move to reposition, " +
                    "replace for find/replace, accept/reject to resolve tracked revisions.",
                    candidates: Kinds);
            }

            if (string.IsNullOrWhiteSpace(op.Path))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{i}].path is missing.",
                    "Every operation needs a target path, e.g. /body/p[3] or /Sheet1/A1.");
            }

            // "/" is the document-wide replace scope (M4): the command layer
            // expands it into the format's default scopes before handlers run.
            if (op.Op != "replace" || op.Path != "/")
            {
                _ = DocPath.Parse(op.Path); // throws invalid_path with grammar hint
            }
        }

        return ops;
    }
}

/// <summary>Shared JSON settings for the wire surface (camelCase, no nulls).</summary>
public static class JsonDefaults
{
    public static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
    };
}

/// <summary>Everything a handler needs to execute one command.</summary>
public sealed class CommandContext
{
    public required Workspace Workspace { get; init; }

    /// <summary>Sandbox-resolved absolute path of the target file, when the verb takes one.</summary>
    public string? File { get; init; }

    /// <summary>Verb-specific arguments as parsed JSON (selector, view, data, …).</summary>
    public JsonObject Args { get; init; } = [];
}

/// <summary>
/// A format handler implements the whole verb surface for one document kind.
/// Unimplemented capabilities must throw <c>unsupported_feature</c> (with a
/// workaround in the suggestion) — never crash, never silently no-op.
/// </summary>
public interface IFormatHandler
{
    DocumentKind Kind { get; }

    Envelope Create(CommandContext ctx);

    /// <summary>
    /// M5 import hook: build a NEW document from a foreign source file
    /// (markdown → docx, csv → xlsx). <paramref name="sourcePath"/> arrives
    /// sandbox-resolved by the command layer; handlers re-resolve defensively.
    /// Additive default: formats without an importer refuse with
    /// <c>unsupported_feature</c> and name the workaround.
    /// </summary>
    Envelope CreateFrom(CommandContext ctx, string sourcePath) =>
        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"{Kind.ToString().ToLowerInvariant()} cannot be created from a source file yet.",
            "create --from supports .md/.markdown → .docx and .csv/.tsv → .xlsx. " +
            "Create an empty document instead and add content with 'aioffice edit'.",
            candidates: [".md", ".markdown", ".csv", ".tsv"]);

    Envelope Read(CommandContext ctx);
    Envelope Get(CommandContext ctx);
    Envelope Query(CommandContext ctx);
    Envelope Edit(CommandContext ctx, IReadOnlyList<EditOp> ops);
    Envelope Render(CommandContext ctx);
    Envelope Validate(CommandContext ctx);
    Envelope Template(CommandContext ctx);
}

/// <summary>
/// One finding from an <c>aioffice audit</c> pass. <see cref="Id"/> is stable
/// within a run (e.g. <c>"a11y_no_alt_text#/slide[2]/shape[@id=5]"</c>) so
/// <c>--fix</c> can target specific findings; <see cref="Autofixable"/> says
/// whether the default <c>--fix</c> will resolve it.
/// </summary>
public sealed record AuditFinding
{
    /// <summary>Stable-within-a-run id, conventionally <c>code#path</c>.</summary>
    [JsonPropertyName("id")]
    public required string Id { get; init; }

    /// <summary>error | warning | info.</summary>
    [JsonPropertyName("severity")]
    public required string Severity { get; init; }

    /// <summary>accessibility | quality.</summary>
    [JsonPropertyName("category")]
    public required string Category { get; init; }

    /// <summary>The finding code namespace, e.g. a11y_no_alt_text, quality_broken_ref.</summary>
    [JsonPropertyName("code")]
    public required string Code { get; init; }

    /// <summary>The canonical path the finding points at, when it has one.</summary>
    [JsonPropertyName("path")]
    public string? Path { get; init; }

    [JsonPropertyName("message")]
    public required string Message { get; init; }

    [JsonPropertyName("suggestion")]
    public required string Suggestion { get; init; }

    /// <summary>True when <c>--fix</c> can safely resolve this finding.</summary>
    [JsonPropertyName("autofixable")]
    public required bool Autofixable { get; init; }
}

/// <summary>The result of an audit pass: every finding plus a severity tally.</summary>
public sealed record AuditResult
{
    [JsonPropertyName("findings")]
    public required IReadOnlyList<AuditFinding> Findings { get; init; }

    [JsonPropertyName("summary")]
    public required AuditSummary Summary { get; init; }
}

/// <summary>Per-severity counts over an <see cref="AuditResult"/>.</summary>
public sealed record AuditSummary(
    [property: JsonPropertyName("errors")] int Errors,
    [property: JsonPropertyName("warnings")] int Warnings,
    [property: JsonPropertyName("infos")] int Infos);

/// <summary>Knobs for one audit pass.</summary>
public sealed record AuditOptions
{
    /// <summary>accessibility | quality | all.</summary>
    [JsonPropertyName("category")]
    public string Category { get; init; } = "all";

    /// <summary>error | warning | info — the lowest severity to report.</summary>
    [JsonPropertyName("minSeverity")]
    public string MinSeverity { get; init; } = "info";

    /// <summary>True when the caller wants safe autofixes applied.</summary>
    [JsonPropertyName("fix")]
    public bool Fix { get; init; }

    /// <summary>The audit categories, severities (low → high) and their ordering helpers.</summary>
    public static readonly IReadOnlyList<string> Categories = ["accessibility", "quality", "all"];

    public static readonly IReadOnlyList<string> Severities = ["info", "warning", "error"];

    /// <summary>Severity rank (info=0, warning=1, error=2); unknown sorts lowest.</summary>
    public static int SeverityRank(string severity) => severity switch
    {
        "error" => 2,
        "warning" => 1,
        _ => 0,
    };
}

/// <summary>
/// The M7 audit surface: a handler inspects a document for accessibility and
/// quality findings and can apply the safe subset of autofixes. Every format
/// handler implements this; <see cref="Fix"/> is never destructive.
/// </summary>
public interface IAuditor
{
    /// <summary>Runs the audit checks selected by <paramref name="opts"/>.</summary>
    AuditResult Audit(CommandContext ctx, AuditOptions opts);

    /// <summary>Applies the autofixes for the given finding ids and reports how many landed.</summary>
    int Fix(CommandContext ctx, IReadOnlyList<string> findingIds);
}

/// <summary>
/// One semantic difference between a baseline document and the current one,
/// surfaced by <c>aioffice diff</c>. <see cref="Path"/> is the canonical path
/// in the CURRENT document (or the baseline path for a <c>removed</c> change,
/// which no longer exists in the current one). <see cref="Before"/>/<see cref="After"/>
/// carry concise human-readable values for a <c>modified</c> change.
/// </summary>
public sealed record DiffChange
{
    /// <summary>added | removed | modified | moved.</summary>
    [JsonPropertyName("kind")]
    public required string Kind { get; init; }

    /// <summary>Canonical path in the current document (baseline path for <c>removed</c>).</summary>
    [JsonPropertyName("path")]
    public required string Path { get; init; }

    /// <summary>The baseline value for a <c>modified</c> change (e.g. old text / "Normal").</summary>
    [JsonPropertyName("before")]
    public string? Before { get; init; }

    /// <summary>The current value for a <c>modified</c> change (e.g. new text / "Heading1").</summary>
    [JsonPropertyName("after")]
    public string? After { get; init; }

    /// <summary>A short note disambiguating the change (e.g. "style", "text", "moved from /body/p[5]").</summary>
    [JsonPropertyName("detail")]
    public string? Detail { get; init; }
}

/// <summary>Per-kind tally over a <see cref="DiffResult"/>.</summary>
public sealed record DiffSummary(
    [property: JsonPropertyName("added")] int Added,
    [property: JsonPropertyName("removed")] int Removed,
    [property: JsonPropertyName("modified")] int Modified,
    [property: JsonPropertyName("moved")] int Moved);

/// <summary>The result of a diff pass: every change plus a per-kind tally.</summary>
public sealed record DiffResult
{
    [JsonPropertyName("changes")]
    public required IReadOnlyList<DiffChange> Changes { get; init; }

    [JsonPropertyName("summary")]
    public required DiffSummary Summary { get; init; }

    /// <summary>
    /// Non-fatal diff warnings (e.g. <c>diff_truncated</c> when the per-pass
    /// change cap was hit on a very large diff). Null when there are none; the
    /// verb layer copies these into the envelope's <c>meta.warnings</c>.
    /// </summary>
    [JsonPropertyName("warnings")]
    public IReadOnlyList<Warning>? Warnings { get; init; }

    /// <summary>An empty result (identical documents): no changes, all-zero summary.</summary>
    public static DiffResult Empty { get; } = new()
    {
        Changes = [],
        Summary = new DiffSummary(0, 0, 0, 0),
    };

    /// <summary>
    /// Sorts the changes by (Path, Kind) with an ordinal comparer and tallies the
    /// summary, so the same two documents diff identically on every platform.
    /// Optional <paramref name="warnings"/> ride along to the envelope meta.
    /// </summary>
    public static DiffResult FromChanges(
        IEnumerable<DiffChange> changes, IReadOnlyList<Warning>? warnings = null)
    {
        var sorted = changes
            .OrderBy(c => c.Path, StringComparer.Ordinal)
            .ThenBy(c => c.Kind, StringComparer.Ordinal)
            .ThenBy(c => c.Detail ?? string.Empty, StringComparer.Ordinal)
            .ToList();

        return new DiffResult
        {
            Changes = sorted,
            Summary = new DiffSummary(
                Added: sorted.Count(c => c.Kind == "added"),
                Removed: sorted.Count(c => c.Kind == "removed"),
                Modified: sorted.Count(c => c.Kind == "modified"),
                Moved: sorted.Count(c => c.Kind == "moved")),
            Warnings = warnings is { Count: > 0 } ? warnings : null,
        };
    }
}

/// <summary>The detail level of a diff projection.</summary>
public static class DiffView
{
    public const string Summary = "summary";

    public const string Detailed = "detailed";

    public static readonly IReadOnlyList<string> All = [Summary, Detailed];
}

/// <summary>
/// The M8 diff surface: a handler semantically compares the current document
/// (<see cref="CommandContext.File"/>) against a same-format baseline and
/// returns the ordered change list. Both documents must be the same kind; a
/// format mismatch is <c>invalid_args</c> naming the mismatch.
/// </summary>
public interface IDiffer
{
    /// <summary>Diffs <paramref name="baselineFile"/> (the OLD document) against <c>ctx.File</c> (the NEW one).</summary>
    DiffResult Diff(CommandContext ctx, string baselineFile);
}

/// <summary>Maps file extensions to format handlers.</summary>
public sealed class HandlerRegistry
{
    private readonly Dictionary<string, IFormatHandler> _byExtension = new(StringComparer.OrdinalIgnoreCase);

    public void Register(IFormatHandler handler, params string[] extensions)
    {
        foreach (var extension in extensions)
        {
            var key = extension.StartsWith('.') ? extension : "." + extension;
            _byExtension[key] = handler;
        }
    }

    public IReadOnlyCollection<string> KnownExtensions => _byExtension.Keys;

    /// <summary>Resolves the handler for a file path or throws <c>unsupported_feature</c>.</summary>
    public IFormatHandler Resolve(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        if (extension.Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cannot infer the document kind: '{filePath}' has no extension.",
                "Use a file name ending in .docx, .xlsx or .pptx, or pass --kind explicitly.");
        }

        if (_byExtension.TryGetValue(extension, out var handler))
        {
            return handler;
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"No handler is registered for '{extension}' files.",
            "Supported formats: " + string.Join(", ", _byExtension.Keys.Order(StringComparer.Ordinal)) +
            ". Convert the file to one of these first.",
            candidates: [.. _byExtension.Keys.Order(StringComparer.Ordinal)]);
    }
}
