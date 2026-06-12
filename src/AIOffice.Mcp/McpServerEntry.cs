using AIOffice.Core;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

namespace AIOffice.Mcp;

/// <summary>
/// Entry point for <c>aioffice mcp</c>: a stdio MCP server exposing the 12 v0
/// tools (see <see cref="ToolCatalog"/>) over the exact same
/// <see cref="CommandService"/> the CLI verbs use — one command layer, two doors.
/// </summary>
public static class McpServerEntry
{
    /// <summary>The server name advertised during the MCP handshake.</summary>
    public const string ServerName = "aioffice";

    /// <summary>
    /// Runs the stdio MCP server until the client disconnects or the token is
    /// cancelled. Workspace root comes from <c>AIOFFICE_WORKSPACE</c>, falling
    /// back to the current directory. This is the overload the CLI's
    /// <c>mcp</c> verb calls.
    /// </summary>
    public static Task RunAsync(CancellationToken cancellationToken) =>
        RunAsync(workspaceRoot: null, cancellationToken);

    /// <summary>
    /// Runs the stdio MCP server with an explicit workspace root
    /// (<c>aioffice mcp --workspace dir</c>).
    /// </summary>
    public static async Task RunAsync(string? workspaceRoot, CancellationToken cancellationToken)
    {
        var root = workspaceRoot
            ?? Environment.GetEnvironmentVariable("AIOFFICE_WORKSPACE")
            ?? Directory.GetCurrentDirectory();
        // One snapshot ring (default ~/.aioffice/snapshots) shared by the service
        // and the handlers it injects into; kinds that self-snapshot are reported
        // so the service never snapshots the same pre-image twice.
        var snapshots = new SnapshotStore();
        var registry = HandlerDiscovery.CreateDefaultRegistry(snapshots, out var handlerManaged);
        var service = new CommandService(
            new Workspace(root), registry, handlerManagedSnapshots: handlerManaged);

        await using var transport = new StdioServerTransport(ServerName);
        await using var server = McpServer.Create(transport, BuildOptions(service));
        await server.RunAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Builds the server options around a command service. Public so tests (and
    /// embedders) can host the identical server over an in-memory transport.
    /// </summary>
    public static McpServerOptions BuildOptions(CommandService service)
    {
        ArgumentNullException.ThrowIfNull(service);
        return new McpServerOptions
        {
            ServerInfo = new Implementation
            {
                Name = ServerName,
                Title = "AIOffice — AI-native Office document tools",
                Version = Meta.ToolVersion,
            },
            ServerInstructions =
                "Tools for real .docx/.xlsx/.pptx files. Every result is one JSON envelope " +
                "{ok,data,error,meta}; on ok=false read error.suggestion and error.candidates before retrying. " +
                "Never guess property names or selectors - call office_help first; office_schema returns the whole surface.",
            Capabilities = new ServerCapabilities
            {
                Tools = new ToolsCapability(),
            },
            Handlers = new McpServerHandlers
            {
                ListToolsHandler = static (_, _) =>
                    ValueTask.FromResult(new ListToolsResult { Tools = [.. ToolCatalog.Tools] }),
                CallToolHandler = (request, _) =>
                    ValueTask.FromResult(ToolRouter.Call(service, request.Params)),
            },
        };
    }
}
