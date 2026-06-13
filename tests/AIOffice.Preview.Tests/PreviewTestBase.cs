using System.Net.Sockets;
using System.Text;
using System.Text.Json.Nodes;
using AIOffice.Core;
using AIOffice.Excel;
using AIOffice.Pptx;
using AIOffice.Word;
using Xunit;

namespace AIOffice.Preview.Tests;

/// <summary>
/// Per-test temp workspace (canonical root, so paths match server output), an
/// isolated lockfile directory and snapshot ring, plus fixture builders that
/// create real documents through the format handlers.
/// </summary>
public abstract class PreviewTestBase : IDisposable
{
    /// <summary>Shared client for request/response routes (SSE tests use their own, untimed client).</summary>
    protected static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(30) };

    protected PreviewTestBase()
    {
        var raw = Path.Combine(Path.GetTempPath(), "aioffice-preview-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(raw);
        Workspace = new Workspace(raw);
        Dir = Workspace.Root;
        LockDir = Path.Combine(Dir, ".preview-locks");
        Snapshots = new SnapshotStore(Path.Combine(Dir, ".snapshots"));
    }

    protected string Dir { get; }

    protected Workspace Workspace { get; }

    protected string LockDir { get; }

    protected SnapshotStore Snapshots { get; }

    public void Dispose()
    {
        if (Directory.Exists(Dir))
        {
            Directory.Delete(Dir, recursive: true);
        }

        GC.SuppressFinalize(this);
    }

    protected PreviewServer StartServer(string file, int port = 0)
    {
        var server = PreviewServer.Start(file, Workspace, port, LockDir);

        // Wait until the listener actually answers a port probe before handing the
        // server back. On Windows the HttpListener (http.sys) registers its prefix
        // slightly after Start() returns, so an immediate Close/probe can race and
        // see the server as not-running; poll for readiness so every preview socket
        // test is deterministic on all platforms.
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(5);
        while (!PreviewLock.IsPortAlive(server.Port) && DateTime.UtcNow < deadline)
        {
            System.Threading.Thread.Sleep(25);
        }

        return server;
    }

    protected CommandContext Ctx(string file, params (string Key, JsonNode? Value)[] args)
    {
        var obj = new JsonObject();
        foreach (var (key, value) in args)
        {
            obj[key] = value;
        }

        return new CommandContext { Workspace = Workspace, File = file, Args = obj };
    }

    protected static void AssertOk(Envelope envelope) => Assert.True(envelope.IsOk, envelope.ToJson());

    /// <summary>A docx with a Heading1 paragraph, one body paragraph and a 2x2 table.</summary>
    protected string CreateDocx(string name = "doc.docx")
    {
        var handler = new WordHandler(Snapshots);
        var file = Path.Combine(Dir, name);
        AssertOk(handler.Create(Ctx(file, ("title", "Preview Heading"))));
        AssertOk(handler.Edit(Ctx(file), EditOp.ParseBatch("""
            [
              {"op":"set","path":"/body/p[2]","props":{"text":"Hello preview"}},
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}}
            ]
            """)));
        return file;
    }

    /// <summary>An xlsx with A1="Name" and B2=42 on Sheet1.</summary>
    protected string CreateXlsx(string name = "book.xlsx")
    {
        var handler = new ExcelHandler(Snapshots);
        var file = Path.Combine(Dir, name);
        AssertOk(handler.Create(Ctx(file)));
        AssertOk(handler.Edit(Ctx(file), EditOp.ParseBatch("""
            [
              {"op":"set","path":"/Sheet1/A1","props":{"value":"Name"}},
              {"op":"set","path":"/Sheet1/B2","props":{"value":42}}
            ]
            """)));
        return file;
    }

    /// <summary>A pptx whose single slide carries a title shape.</summary>
    protected string CreatePptx(string name = "deck.pptx")
    {
        var handler = new PptxHandler();
        var file = Path.Combine(Dir, name);
        AssertOk(handler.Create(Ctx(file, ("title", "Preview Deck"))));
        return file;
    }

    protected static Task<string> GetStringAsync(PreviewServer server, string route = "") =>
        Http.GetStringAsync(server.Url + route);

    protected static Task<HttpResponseMessage> PostJsonAsync(PreviewServer server, string route, string json) =>
        Http.PostAsync(server.Url + route, new StringContent(json, Encoding.UTF8, "application/json"));

    /// <summary>A loopback port that was just released, i.e. (almost certainly) dead.</summary>
    protected static int DeadPort()
    {
        var probe = new TcpListener(System.Net.IPAddress.Loopback, 0);
        probe.Start();
        var port = ((System.Net.IPEndPoint)probe.LocalEndpoint).Port;
        probe.Stop();
        return port;
    }
}
