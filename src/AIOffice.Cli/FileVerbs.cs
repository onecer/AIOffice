using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using AIOffice.Core.Cli;

namespace AIOffice.Cli;

/// <summary>
/// Dispatches the eight file-taking verbs into the format handlers and owns the
/// cross-format edit pipeline: --expect-rev is checked BEFORE any write and
/// every (non dry-run) edit snapshots the pre-image first.
/// </summary>
public sealed class FileVerbs
{
    private readonly Workspace _workspace;
    private readonly DiscoveredHandlers _handlers;
    private readonly SnapshotStore _snapshots;

    public FileVerbs(Workspace workspace, DiscoveredHandlers handlers, SnapshotStore snapshots)
    {
        _workspace = workspace;
        _handlers = handlers;
        _snapshots = snapshots;
    }

    public Envelope Create(ParsedArgs args)
    {
        var file = RequireFile(args, mustExist: false);
        var handler = ResolveHandler(file, args.GetOption("kind"));
        var ctx = Context(file, new JsonObject
        {
            ["kind"] = handler.Kind.ToString().ToLowerInvariant(),
            ["title"] = args.GetOption("title"),
        });
        return handler.Create(ctx);
    }

    public Envelope Read(ParsedArgs args)
    {
        var file = RequireFile(args, mustExist: true);
        var ctx = Context(file, new JsonObject
        {
            ["view"] = args.GetOption("view"),
            ["range"] = args.GetOption("range"),
            ["maxBytes"] = ParseOptionalInt(args, "max-bytes"),
        });
        return ResolveHandler(file, kindOverride: null).Read(ctx);
    }

    public Envelope Query(ParsedArgs args)
    {
        var file = RequireFile(args, mustExist: true);
        var selectorText = RequirePositional(args, 1, "selector",
            "aioffice query <file> <selector> — e.g. aioffice query report.docx \"p[style=Heading1]\".");
        _ = Selector.Parse(selectorText); // fail fast with the grammar hint
        var ctx = Context(file, new JsonObject { ["selector"] = selectorText });
        return ResolveHandler(file, kindOverride: null).Query(ctx);
    }

    public Envelope Get(ParsedArgs args)
    {
        var file = RequireFile(args, mustExist: true);
        var pathText = RequirePositional(args, 1, "path",
            "aioffice get <file> <path> — e.g. aioffice get report.docx /body/p[3].");
        _ = DocPath.Parse(pathText); // fail fast with the grammar hint
        var ctx = Context(file, new JsonObject { ["path"] = pathText });
        return ResolveHandler(file, kindOverride: null).Get(ctx);
    }

    public Envelope Edit(ParsedArgs args)
    {
        var file = RequireFile(args, mustExist: true);
        var ops = CollectOps(args); // validate input before anything else
        var handler = ResolveHandler(file, kindOverride: null);
        var dryRun = args.HasFlag("dry-run");

        // Optimistic concurrency: verified BEFORE any write or snapshot.
        if (args.GetOption("expect-rev") is { } expected)
        {
            var actual = Rev.OfFile(file);
            if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
            {
                throw new AiofficeException(
                    ErrorCodes.StaleAddress,
                    $"File rev is {actual} but --expect-rev was {expected}; the document changed since you last read it.",
                    "Re-run 'aioffice read' or 'aioffice query' to refresh paths, then retry with the new rev.");
            }
        }

        // Pre-image snapshot, so every edit is undoable. Handlers that received
        // a SnapshotStore at construction own this themselves (and snapshot
        // only on success); for the rest the CLI snapshots up front.
        if (!dryRun && !_handlers.HandlerManagedSnapshots.Contains(handler.Kind))
        {
            _snapshots.Save(file);
        }

        var ctx = Context(file, new JsonObject { ["dryRun"] = dryRun });
        return handler.Edit(ctx, ops);
    }

    public Envelope Render(ParsedArgs args)
    {
        var file = RequireFile(args, mustExist: true);
        var scope = args.GetOption("scope");
        if (scope is not null)
        {
            _ = DocPath.Parse(scope);
        }

        var output = args.GetOption("o") ?? args.GetOption("out");
        var ctx = Context(file, new JsonObject
        {
            ["to"] = args.GetOption("to"),
            ["scope"] = scope,
            ["output"] = output is null ? null : _workspace.Resolve(output),
        });
        return ResolveHandler(file, kindOverride: null).Render(ctx);
    }

    public Envelope Validate(ParsedArgs args)
    {
        var file = RequireFile(args, mustExist: true);
        var ctx = Context(file, new JsonObject());
        return ResolveHandler(file, kindOverride: null).Validate(ctx);
    }

    public Envelope Template(ParsedArgs args)
    {
        var file = RequireFile(args, mustExist: true);
        var dataText = args.GetOption("data") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "template requires --data <json|@file>.",
            "Pass the merge values as JSON, e.g. --data '{\"customer\":\"ACME\"}' or --data @values.json.");

        JsonNode? data;
        try
        {
            data = JsonNode.Parse(ArgParser.ExpandAtFile(dataText, _workspace));
        }
        catch (JsonException ex)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"--data is not valid JSON: {ex.Message}",
                "Pass a JSON object like {\"customer\":\"ACME\",\"total\":42} or @values.json.",
                innerException: ex);
        }

        var output = args.GetOption("o") ?? args.GetOption("out");
        var ctx = Context(file, new JsonObject
        {
            ["data"] = data,
            ["output"] = output is null ? null : _workspace.Resolve(output),
        });
        return ResolveHandler(file, kindOverride: null).Template(ctx);
    }

    // ----- shared plumbing -------------------------------------------------

    private CommandContext Context(string file, JsonObject verbArgs) =>
        new() { Workspace = _workspace, File = file, Args = verbArgs };

    private string RequireFile(ParsedArgs args, bool mustExist)
    {
        var userPath = RequirePositional(args, 0, "file",
            "Every document verb starts with the file, e.g. aioffice read report.docx.");
        return _workspace.Resolve(userPath, mustExist);
    }

    private static string RequirePositional(ParsedArgs args, int index, string name, string suggestion)
    {
        if (args.Positionals.Count <= index)
        {
            throw new AiofficeException(ErrorCodes.InvalidArgs, $"Missing required argument: <{name}>.", suggestion);
        }

        return args.Positionals[index];
    }

    private IFormatHandler ResolveHandler(string file, string? kindOverride)
    {
        if (kindOverride is not null)
        {
            if (!Enum.TryParse<DocumentKind>(kindOverride, ignoreCase: true, out var kind))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown --kind: '{kindOverride}'.",
                    "Use one of: docx, xlsx, pptx.",
                    candidates: ["docx", "xlsx", "pptx"]);
            }

            if (_handlers.ByKind.TryGetValue(kind, out var byKind))
            {
                return byKind;
            }

            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"The {kindOverride} handler is not available in this build yet.",
                "Run 'aioffice doctor' to see handler status; use one of the ready formats meanwhile.",
                candidates: [.. _handlers.ByKind.Keys.Select(k => k.ToString().ToLowerInvariant())]);
        }

        if (_handlers.ByKind.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                "No format handlers are available in this build yet.",
                "Run 'aioffice doctor' to see per-format status; the docx/xlsx/pptx handlers land with their packages.");
        }

        return _handlers.Registry.Resolve(file);
    }

    private static int? ParseOptionalInt(ParsedArgs args, string option)
    {
        var text = args.GetOption(option);
        if (text is null)
        {
            return null;
        }

        if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var value) || value <= 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"--{option} must be a positive integer, got '{text}'.",
                $"Pass a number, e.g. --{option} 65536.");
        }

        return value;
    }

    /// <summary>Builds the op batch from --ops or from the --set/--add/--remove sugar.</summary>
    private IReadOnlyList<EditOp> CollectOps(ParsedArgs args)
    {
        if (args.GetOption("ops") is { } opsText)
        {
            return EditOp.ParseBatch(ArgParser.ExpandAtFile(opsText, _workspace));
        }

        var setPath = args.GetOption("set");
        var addPath = args.GetOption("add");
        var removePaths = args.GetOptionValues("remove");

        if (setPath is null && addPath is null && removePaths.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "edit needs --ops, or one of the sugar forms --set/--add/--remove.",
                "Example: aioffice edit report.docx --set /body/p[1] text='Hello' — or pass a batch with --ops '[...]'.");
        }

        if (setPath is not null && addPath is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "--set and --add cannot be combined in one sugar invocation (the trailing k=v pairs would be ambiguous).",
                "Use --ops for multi-op batches: [{\"op\":\"set\",...},{\"op\":\"add\",...}].");
        }

        var props = ParseProps(args.Positionals.Skip(1));
        var ops = new List<EditOp>();

        if (setPath is not null)
        {
            if (props.Count == 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "--set needs at least one trailing k=v property.",
                    "Example: aioffice edit report.docx --set /body/p[1] text='New text' style=Heading1.");
            }

            ops.Add(new EditOp { Op = "set", Path = setPath, Props = props });
        }

        if (addPath is not null)
        {
            var type = args.GetOption("type") ?? throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "--add needs --type <element> (e.g. p, table, slide, shape).",
                "Example: aioffice edit deck.pptx --add /slide[1] --type slide --position after.");
            ops.Add(new EditOp
            {
                Op = "add",
                Path = addPath,
                Type = type,
                Props = props.Count > 0 ? props : null,
                Position = args.GetOption("position"),
            });
        }

        foreach (var removePath in removePaths)
        {
            ops.Add(new EditOp { Op = "remove", Path = removePath });
        }

        foreach (var op in ops)
        {
            _ = DocPath.Parse(op.Path); // same fail-fast validation --ops gets
        }

        return ops;
    }

    /// <summary>Parses trailing k=v pairs; values that parse as JSON literals keep their type.</summary>
    private static JsonObject ParseProps(IEnumerable<string> pairs)
    {
        var props = new JsonObject();
        foreach (var pair in pairs)
        {
            var equals = pair.IndexOf('=', StringComparison.Ordinal);
            if (equals <= 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Expected k=v, got '{pair}'.",
                    "Properties look like text='Hello world', style=Heading1 or value=42.");
            }

            var key = pair[..equals];
            var raw = pair[(equals + 1)..];
            props[key] = ParseScalar(raw);
        }

        return props;
    }

    private static JsonNode? ParseScalar(string raw)
    {
        if (raw.Length == 0)
        {
            return string.Empty;
        }

        // true/false/null/numbers/quoted strings/objects/arrays parse as JSON;
        // everything else is a plain string.
        try
        {
            return JsonNode.Parse(raw);
        }
        catch (JsonException)
        {
            return raw;
        }
    }
}
