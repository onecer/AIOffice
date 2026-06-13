using AIOffice.Core;
using Xunit;

namespace AIOffice.Preview.Tests;

public sealed class PreviewClientTests : PreviewTestBase
{
    [Fact]
    public void GetSelection_without_a_lockfile_is_preview_not_running()
    {
        var file = CreateDocx();

        var ex = Assert.Throws<AiofficeException>(() => PreviewClient.GetSelection(file, LockDir));

        Assert.Equal(ErrorCodes.PreviewNotRunning, ex.Code);
        Assert.Contains("aioffice preview open", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void GetSelection_with_a_dead_port_lockfile_is_preview_not_running_and_cleans_up()
    {
        var file = CreateDocx();
        var lockPath = PreviewLock.PathFor(file, LockDir);
        PreviewLock.Write(lockPath, new PreviewLockfile(DeadPort(), 99999, file, DateTimeOffset.UtcNow));

        var ex = Assert.Throws<AiofficeException>(() => PreviewClient.GetSelection(file, LockDir));

        Assert.Equal(ErrorCodes.PreviewNotRunning, ex.Code);
        Assert.False(File.Exists(lockPath), "a dead lockfile must be cleaned up");
    }

    [Fact]
    public async Task GetSelection_returns_what_the_browser_posted()
    {
        var file = CreateDocx();
        using var server = StartServer(file);
        await PostJsonAsync(server, "selection", """{"paths":["/body/p[1]","/body/table[1]"]}""");

        var snapshot = PreviewClient.GetSelection(file, LockDir);

        Assert.Equal(["/body/p[1]", "/body/table[1]"], snapshot.Paths);
        Assert.Equal(Rev.OfFile(file), snapshot.Rev);
    }

    [Fact]
    public async Task Close_stops_the_server_and_a_second_close_is_preview_not_running()
    {
        var file = CreateDocx();
        using var server = StartServer(file);

        PreviewClient.Close(file, LockDir);

        await server.WaitForShutdownAsync().WaitAsync(TimeSpan.FromSeconds(5));

        // The lockfile deletion is eventually-consistent: on Windows a transient
        // sharing violation (AV/indexer) can briefly hold the file, so the server
        // retries and the stale-port path cleans up. Poll rather than demand an
        // instantaneous delete so the test is deterministic on every platform.
        await WaitUntil(() => !File.Exists(server.LockfilePath), TimeSpan.FromSeconds(5));
        Assert.False(File.Exists(server.LockfilePath), "lockfile must be deleted on close");

        var ex = Assert.Throws<AiofficeException>(() => PreviewClient.Close(file, LockDir));
        Assert.Equal(ErrorCodes.PreviewNotRunning, ex.Code);
    }

    private static async Task WaitUntil(Func<bool> condition, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (!condition() && DateTime.UtcNow < deadline)
        {
            await Task.Delay(25);
        }
    }
}
