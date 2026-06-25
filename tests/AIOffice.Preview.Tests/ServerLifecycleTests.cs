using AIOffice.Core;
using Xunit;

namespace AIOffice.Preview.Tests;

public sealed class ServerLifecycleTests : PreviewTestBase
{
    [Fact]
    public async Task Start_binds_a_port_in_the_preview_range_and_serves_html()
    {
        var file = CreateDocx();
        using var server = StartServer(file);

        Assert.InRange(server.Port, PreviewServer.PortRangeStart, PreviewServer.PortRangeEnd);

        var response = await Http.GetAsync(server.Url);
        Assert.Equal(System.Net.HttpStatusCode.OK, response.StatusCode);
        Assert.Equal("text/html", response.Content.Headers.ContentType?.MediaType);
        Assert.StartsWith("<!DOCTYPE html>", await response.Content.ReadAsStringAsync(), StringComparison.Ordinal);
    }

    [Fact]
    public void Start_writes_the_lockfile_and_dispose_deletes_it()
    {
        var file = CreateDocx();
        var server = StartServer(file);

        var lockfile = PreviewLock.TryRead(server.LockfilePath);
        Assert.NotNull(lockfile);
        Assert.Equal(server.Port, lockfile.Port);
        Assert.Equal(Environment.ProcessId, lockfile.Pid);
        Assert.Equal(file, lockfile.File);

        server.Dispose();
        Assert.False(File.Exists(server.LockfilePath), "lockfile must be deleted on shutdown");
    }

    [Fact]
    public void Start_overwrites_a_stale_lockfile_whose_port_is_dead()
    {
        var file = CreateDocx();
        var lockPath = PreviewLock.PathFor(file, LockDir);
        PreviewLock.Write(lockPath, new PreviewLockfile(DeadPort(), 99999, file, DateTimeOffset.UtcNow));

        using var server = StartServer(file);

        Assert.Equal(lockPath, server.LockfilePath);
        var rewritten = PreviewLock.TryRead(lockPath);
        Assert.NotNull(rewritten);
        Assert.Equal(server.Port, rewritten.Port);
        Assert.Equal(Environment.ProcessId, rewritten.Pid);
    }

    [Fact]
    public void Start_rejects_a_second_server_while_the_first_is_alive()
    {
        var file = CreateDocx();
        using var first = StartServer(file);

        var ex = Assert.Throws<AiofficeException>(() => StartServer(file));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("already running", ex.Message, StringComparison.Ordinal);
        Assert.False(string.IsNullOrWhiteSpace(ex.Suggestion));
    }

    [Fact]
    public void Start_rejects_unsupported_extensions_with_a_workaround()
    {
        var file = Path.Combine(Dir, "notes.txt");
        File.WriteAllText(file, "not an office file");

        var ex = Assert.Throws<AiofficeException>(() => StartServer(file));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.False(string.IsNullOrWhiteSpace(ex.Suggestion));
        Assert.Contains(".docx", ex.Candidates ?? []);
    }

    [Fact]
    public async Task Shutdown_route_stops_the_server_and_deletes_the_lockfile()
    {
        var file = CreateDocx();
        using var server = StartServer(file);

        var response = await Http.PostAsync(server.Url + "shutdown", content: null);
        Assert.Equal(System.Net.HttpStatusCode.OK, response.StatusCode);
        Assert.Contains("\"ok\":true", await response.Content.ReadAsStringAsync(), StringComparison.Ordinal);

        // WaitForShutdownAsync completes only after the server ran its full stop sequence
        // (close the listener, delete the lockfile, signal). That — plus the lockfile being
        // gone — IS the "shutdown route stops the server" contract. We deliberately do NOT
        // probe raw port liveness or re-request the URL: on windows-latest the http.sys
        // request queue drains asynchronously after HttpListener.Close(), so a follow-up GET
        // can still be served (and a TCP probe can still connect) for a brief, OS-controlled
        // window. That teardown timing is not AIOffice's contract — the preview port-scanner
        // already tolerates a not-yet-released port — and asserting it here was flaky.
        await server.WaitForShutdownAsync().WaitAsync(TimeSpan.FromSeconds(10));
        Assert.False(File.Exists(server.LockfilePath), "lockfile must be deleted on shutdown");
    }

    [Fact]
    public async Task Unknown_routes_return_an_envelope_with_a_suggestion()
    {
        var file = CreateDocx();
        using var server = StartServer(file);

        var response = await Http.GetAsync(server.Url + "nope");
        Assert.Equal(System.Net.HttpStatusCode.NotFound, response.StatusCode);
        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("\"suggestion\":", body, StringComparison.Ordinal);
    }
}
