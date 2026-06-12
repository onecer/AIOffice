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
    public static readonly IReadOnlyList<string> Kinds = ["set", "add", "remove", "move", "accept", "reject"];

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
                    $"ops[{i}].op is '{op.Op}' but must be one of: set, add, remove, move, accept, reject.",
                    "Use set to change properties, add to insert, remove to delete, move to reposition, accept/reject to resolve tracked revisions.",
                    candidates: Kinds);
            }

            if (string.IsNullOrWhiteSpace(op.Path))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{i}].path is missing.",
                    "Every operation needs a target path, e.g. /body/p[3] or /Sheet1/A1.");
            }

            _ = DocPath.Parse(op.Path); // throws invalid_path with grammar hint
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
    Envelope Read(CommandContext ctx);
    Envelope Get(CommandContext ctx);
    Envelope Query(CommandContext ctx);
    Envelope Edit(CommandContext ctx, IReadOnlyList<EditOp> ops);
    Envelope Render(CommandContext ctx);
    Envelope Validate(CommandContext ctx);
    Envelope Template(CommandContext ctx);
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
