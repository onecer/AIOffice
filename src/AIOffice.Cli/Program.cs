using System.Diagnostics;
using AIOffice.Core;
using AIOffice.Core.Cli;

namespace AIOffice.Cli;

/// <summary>
/// aioffice entry point: parses argv, builds the sandboxed workspace, dispatches
/// to format handlers or native verbs, and prints exactly one JSON envelope to
/// stdout with the documented exit code. Errors never crash — they become typed
/// envelopes.
/// </summary>
internal static class Program
{
    private static readonly string[] FileVerbNames =
    [
        "create", "read", "query", "get", "edit", "render", "validate", "template", "audit",
    ];

    private static async Task<int> Main(string[] argv)
    {
        var stopwatch = Stopwatch.StartNew();

        ParsedArgs parsed;
        try
        {
            parsed = ArgParser.Parse(argv);
        }
        catch (Exception ex)
        {
            return Print(Envelope.FromException(ex), pretty: !Console.IsOutputRedirected, quiet: false);
        }

        var pretty = parsed.HasFlag("pretty") || (!parsed.HasFlag("json") && !Console.IsOutputRedirected);
        var quiet = parsed.HasFlag("quiet");

        // The MCP server owns stdio (JSON-RPC), so it bypasses envelope printing.
        if (parsed.Verb == "mcp")
        {
            var entry = HandlerDiscovery.FindMcpEntryPoint(out var status);
            if (entry is null)
            {
                var unavailable = Envelope.Fail(
                    ErrorCodes.UnsupportedFeature,
                    $"The MCP server is not available in this build: {status}.",
                    "Use the CLI verbs directly meanwhile — the MCP tools mirror them 1:1 " +
                    "(office_create=create, office_read=read, … office_status=doctor).");
                return Print(Stamp(unavailable, file: null, workspace: null, stopwatch.ElapsedMilliseconds), pretty, quiet);
            }

            try
            {
                var workspaceRoot = parsed.GetOption("workspace")
                    ?? Environment.GetEnvironmentVariable("AIOFFICE_WORKSPACE");
                return await entry(workspaceRoot, argv, CancellationToken.None).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                return Print(Envelope.FromException(ex), pretty, quiet);
            }
        }

        // preview owns its own printing order: the open action must print (and
        // flush) the startup envelope BEFORE it blocks on the server.
        if (parsed.Verb == "preview")
        {
            try
            {
                return PreviewVerbs.Run(parsed, BuildWorkspace(parsed), stopwatch, env => Print(env, pretty, quiet));
            }
            catch (Exception ex)
            {
                return Print(
                    Stamp(Envelope.FromException(ex), MetaFileFor(parsed), workspace: null, stopwatch.ElapsedMilliseconds),
                    pretty, quiet);
            }
        }

        // Resolved up front so error envelopes still carry file + rev meta
        // (a stale_address response should tell the agent the current rev).
        var metaFile = MetaFileFor(parsed);
        Workspace? workspace = null;
        Envelope envelope;
        try
        {
            if (NeedsWorkspace(parsed.Verb))
            {
                workspace = BuildWorkspace(parsed);
            }

            envelope = Execute(parsed, workspace);
        }
        catch (Exception ex)
        {
            envelope = Envelope.FromException(ex);
        }

        envelope = Stamp(envelope, metaFile, workspace, stopwatch.ElapsedMilliseconds);
        return Print(envelope, pretty, quiet);
    }

    private static bool NeedsWorkspace(string? verb) =>
        verb is "doctor" or "snapshot" || (verb is not null && FileVerbNames.Contains(verb, StringComparer.Ordinal));

    /// <summary>The positional that names the file this command touches, if any.</summary>
    private static string? MetaFileFor(ParsedArgs parsed) => parsed.Verb switch
    {
        "snapshot" or "preview" when parsed.Positionals.Count > 1 => parsed.Positionals[1],
        { } verb when FileVerbNames.Contains(verb, StringComparer.Ordinal) && parsed.Positionals.Count > 0 =>
            parsed.Positionals[0],
        _ => null,
    };

    private static Envelope Execute(ParsedArgs parsed, Workspace? workspace)
    {
        var verb = parsed.Verb;
        if (verb is null)
        {
            if (parsed.HasFlag("version"))
            {
                return NativeVerbs.Version();
            }

            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "No command given.",
                "Run 'aioffice help' for the surface or 'aioffice schema' for the machine-readable version.",
                candidates: CommandSurface.VerbNames);
        }

        switch (verb)
        {
            case "version":
                return NativeVerbs.Version();
            case "schema":
                return NativeVerbs.Schema(parsed);
            case "help":
                return NativeVerbs.Help(parsed);
            default:
                break;
        }

        if (workspace is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown command: '{verb}'.",
                "Run 'aioffice help' to list commands.",
                candidates: CommandSurface.NearestVerbs(verb));
        }

        var snapshots = new SnapshotStore();

        if (verb == "doctor")
        {
            var discovered = HandlerDiscovery.Discover(snapshots);
            return new NativeVerbs(workspace, snapshots, discovered.Statuses).Doctor();
        }

        if (verb == "snapshot")
        {
            return new NativeVerbs(workspace, snapshots, []).Snapshot(parsed);
        }

        var fileVerbs = new FileVerbs(workspace, HandlerDiscovery.Discover(snapshots), snapshots);
        return verb switch
        {
            "create" => fileVerbs.Create(parsed),
            "read" => fileVerbs.Read(parsed),
            "query" => fileVerbs.Query(parsed),
            "get" => fileVerbs.Get(parsed),
            "edit" => fileVerbs.Edit(parsed),
            "render" => fileVerbs.Render(parsed),
            "validate" => fileVerbs.Validate(parsed),
            "template" => fileVerbs.Template(parsed),
            "audit" => fileVerbs.Audit(parsed),
            _ => throw new UnreachableException(verb),
        };
    }

    private static Workspace BuildWorkspace(ParsedArgs parsed)
    {
        var root = parsed.GetOption("workspace")
            ?? Environment.GetEnvironmentVariable("AIOFFICE_WORKSPACE")
            ?? Directory.GetCurrentDirectory();
        return new Workspace(root);
    }

    /// <summary>
    /// Stamps authoritative meta on the way out: elapsed time, plus file + rev
    /// (recomputed from the post-command bytes) for verbs that touch a file.
    /// </summary>
    private static Envelope Stamp(Envelope envelope, string? file, Workspace? workspace, long elapsedMs)
    {
        var meta = envelope.Meta;
        var displayFile = meta.File;
        var rev = meta.Rev;

        if (file is not null)
        {
            try
            {
                var resolved = workspace?.Resolve(file) ?? Path.GetFullPath(file);
                displayFile = Display(resolved, workspace);
                if (File.Exists(resolved))
                {
                    rev = Rev.OfFile(resolved);
                }
            }
            catch (Exception ex) when (ex is AiofficeException or IOException or UnauthorizedAccessException)
            {
                // Meta stamping is best-effort; the envelope already carries the real error.
            }
        }

        return envelope with { Meta = meta with { File = displayFile, Rev = rev, ElapsedMs = elapsedMs } };
    }

    /// <summary>Reports files workspace-relative when possible, absolute otherwise.</summary>
    private static string Display(string resolved, Workspace? workspace)
    {
        if (workspace is null)
        {
            return resolved;
        }

        var relative = Path.GetRelativePath(workspace.Root, resolved);
        return relative.StartsWith("..", StringComparison.Ordinal) || Path.IsPathRooted(relative)
            ? resolved
            : relative;
    }

    private static int Print(Envelope envelope, bool pretty, bool quiet)
    {
        if (!quiet || !envelope.IsOk)
        {
            Console.WriteLine(envelope.ToJson(pretty));
        }

        return envelope.ExitCode;
    }
}
