using System.IO.Pipelines;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using AIOffice.Mcp;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>A disposable sandbox root + snapshot dir + command service factory.</summary>
internal sealed class TempWorkspace : IDisposable
{
    private readonly string _baseDir;

    public TempWorkspace()
    {
        _baseDir = Path.Combine(Path.GetTempPath(), "aioffice-mcp-tests", Guid.NewGuid().ToString("N"));
        Root = Path.Combine(_baseDir, "ws");
        SnapshotDir = Path.Combine(_baseDir, "snapshots");
        Directory.CreateDirectory(Root);
        Directory.CreateDirectory(SnapshotDir);
        Workspace = new Workspace(Root);
    }

    public string Root { get; }

    public string SnapshotDir { get; }

    public Workspace Workspace { get; }

    public CommandService NewService(params IFormatHandler[] handlers)
    {
        var registry = new HandlerRegistry();
        foreach (var handler in handlers)
        {
            registry.Register(handler, ExtensionFor(handler.Kind));
        }

        return NewService(registry);
    }

    public CommandService NewService(HandlerRegistry registry, IReadOnlySet<DocumentKind>? handlerManagedSnapshots = null) =>
        new(Workspace, registry, SnapshotDir, handlerManagedSnapshots);

    /// <summary>Writes a file under the workspace root and returns its workspace-relative name.</summary>
    public string WriteFile(string name, string content)
    {
        File.WriteAllText(Path.Combine(Root, name), content);
        return name;
    }

    public string ReadFile(string name) => File.ReadAllText(Path.Combine(Root, name));

    public void Dispose()
    {
        try
        {
            Directory.Delete(_baseDir, recursive: true);
        }
        catch (IOException)
        {
            // Best effort cleanup.
        }
    }

    private static string ExtensionFor(DocumentKind kind) => kind switch
    {
        DocumentKind.Docx => ".docx",
        DocumentKind.Xlsx => ".xlsx",
        DocumentKind.Pptx => ".pptx",
        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
    };
}

/// <summary>
/// A minimal in-test docx handler exercising the MCP plumbing end to end
/// (sandbox, rev guard, snapshots, envelopes) without depending on the real
/// Word package, which lands in parallel.
/// </summary>
internal sealed class FakeDocxHandler : IFormatHandler
{
    public DocumentKind Kind => DocumentKind.Docx;

    public Envelope Create(CommandContext ctx)
    {
        var title = ctx.Args["title"] is JsonValue v && v.TryGetValue<string>(out var t) ? t : string.Empty;
        File.WriteAllText(ctx.File!, "blank:" + title);
        return Envelope.Ok(new { created = ctx.File, kind = "docx" });
    }

    public Envelope Edit(CommandContext ctx, IReadOnlyList<EditOp> ops)
    {
        var dryRun = ctx.Args["dryRun"] is JsonValue dv && dv.TryGetValue<bool>(out var d) && d;
        if (!dryRun)
        {
            var text = string.Join('\n', ops.Select(o =>
                o.Props?["text"] is JsonValue tv && tv.TryGetValue<string>(out var s) ? s : o.Op));
            File.WriteAllText(ctx.File!, text);
        }

        int? snapshot = ctx.Args["snapshot"] is JsonValue sv && sv.TryGetValue<int>(out var n) ? n : null;
        return Envelope.Ok(new
        {
            applied = ops.Count,
            results = ops.Select(o => new { op = o.Op, path = o.Path, ok = true }).ToArray(),
            snapshot,
            dryRun,
        });
    }

    public Envelope Read(CommandContext ctx) => throw NotImplemented("read");

    public Envelope Get(CommandContext ctx) => throw NotImplemented("get");

    public Envelope Query(CommandContext ctx) => throw NotImplemented("query");

    public Envelope Render(CommandContext ctx) => throw NotImplemented("render");

    public Envelope Validate(CommandContext ctx) => throw NotImplemented("validate");

    public Envelope Template(CommandContext ctx) => throw NotImplemented("template");

    private static AiofficeException NotImplemented(string verb) => new(
        ErrorCodes.UnsupportedFeature,
        $"The fake test handler does not implement '{verb}'.",
        "Use the real format handler for this verb.");
}

/// <summary>
/// Boots the real aioffice MCP server in-process over an in-memory duplex pipe
/// and connects the official SDK client to it.
/// </summary>
internal sealed class McpTestServer : IAsyncDisposable
{
    private readonly McpServer _server;
    private readonly Task _run;
    private readonly CancellationTokenSource _cts;

    private McpTestServer(McpServer server, Task run, CancellationTokenSource cts, McpClient client)
    {
        _server = server;
        _run = run;
        _cts = cts;
        Client = client;
    }

    public McpClient Client { get; }

    public static async Task<McpTestServer> StartAsync(CommandService service)
    {
        var clientToServer = new Pipe();
        var serverToClient = new Pipe();

        var server = McpServer.Create(
            new StreamServerTransport(
                clientToServer.Reader.AsStream(),
                serverToClient.Writer.AsStream(),
                serverName: McpServerEntry.ServerName),
            McpServerEntry.BuildOptions(service));

        var cts = new CancellationTokenSource(TimeSpan.FromSeconds(60));
        var run = server.RunAsync(cts.Token);

        var client = await McpClient.CreateAsync(
            new StreamClientTransport(
                serverInput: clientToServer.Writer.AsStream(),
                serverOutput: serverToClient.Reader.AsStream()),
            cancellationToken: cts.Token);

        return new McpTestServer(server, run, cts, client);
    }

    /// <summary>Calls a tool and returns the raw MCP result (content blocks + isError).</summary>
    public async Task<CallToolResult> CallRawAsync(string tool, IReadOnlyDictionary<string, object?>? args = null) =>
        await Client.CallToolAsync(tool, args, cancellationToken: _cts.Token);

    /// <summary>Calls a tool and parses the envelope from its first text content block.</summary>
    public async Task<JsonElement> CallAsync(string tool, IReadOnlyDictionary<string, object?>? args = null)
    {
        var result = await CallRawAsync(tool, args);
        var text = Assert.IsType<TextContentBlock>(result.Content[0]).Text;
        using var doc = JsonDocument.Parse(text);
        return doc.RootElement.Clone();
    }

    public async ValueTask DisposeAsync()
    {
        await Client.DisposeAsync();
        await _cts.CancelAsync();
        try
        {
            await _run.WaitAsync(TimeSpan.FromSeconds(10));
        }
        catch (OperationCanceledException)
        {
            // Normal shutdown path.
        }
        catch (TimeoutException)
        {
            // Do not hang the test run on teardown.
        }

        await _server.DisposeAsync();
        _cts.Dispose();
    }
}

/// <summary>Shared assertions over envelope JSON.</summary>
internal static class EnvelopeAssert
{
    public static JsonElement Ok(JsonElement envelope)
    {
        Assert.True(envelope.GetProperty("ok").GetBoolean(),
            $"expected ok envelope, got: {envelope.GetRawText()}");
        return envelope.GetProperty("data");
    }

    public static JsonElement Fail(JsonElement envelope, string expectedCode)
    {
        Assert.False(envelope.GetProperty("ok").GetBoolean(),
            $"expected failure envelope, got: {envelope.GetRawText()}");
        var error = envelope.GetProperty("error");
        Assert.Equal(expectedCode, error.GetProperty("code").GetString());
        Assert.False(string.IsNullOrWhiteSpace(error.GetProperty("suggestion").GetString()),
            "error.suggestion must never be empty");
        return error;
    }
}
