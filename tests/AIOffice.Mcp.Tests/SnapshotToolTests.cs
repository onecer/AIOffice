using Xunit;

namespace AIOffice.Mcp.Tests;

public sealed class SnapshotToolTests
{
    [Fact]
    public async Task SnapshotListAndRestore_RoundTripsThePreImage()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));
        ws.WriteFile("test.docx", "original");

        // Mutate once (auto-snapshots the pre-image).
        EnvelopeAssert.Ok(await srv.CallAsync("office_edit", new Dictionary<string, object?>
        {
            ["file"] = "test.docx",
            ["ops"] = new[]
            {
                new Dictionary<string, object?>
                {
                    ["op"] = "set",
                    ["path"] = "/body/p[1]",
                    ["props"] = new Dictionary<string, string> { ["text"] = "edited" },
                },
            },
        }));
        Assert.Equal("edited", ws.ReadFile("test.docx"));

        // list shows exactly the auto snapshot.
        var listData = EnvelopeAssert.Ok(await srv.CallAsync("file_snapshot",
            new Dictionary<string, object?> { ["file"] = "test.docx" }));
        var entry = Assert.Single(listData.GetProperty("snapshots").EnumerateArray().ToList());
        Assert.Equal(1, entry.GetProperty("n").GetInt32());
        Assert.Equal("auto", entry.GetProperty("trigger").GetString());
        Assert.True(entry.GetProperty("bytes").GetInt64() > 0);

        // restore rolls the file back...
        var restoreData = EnvelopeAssert.Ok(await srv.CallAsync("file_snapshot",
            new Dictionary<string, object?> { ["file"] = "test.docx", ["action"] = "restore", ["n"] = 1 }));
        Assert.Equal(1, restoreData.GetProperty("restored").GetInt32());
        Assert.Equal("original", ws.ReadFile("test.docx"));

        // ...and is itself undoable: the pre-restore state was snapshotted first.
        var after = EnvelopeAssert.Ok(await srv.CallAsync("file_snapshot",
            new Dictionary<string, object?> { ["file"] = "test.docx" }));
        Assert.Equal(2, after.GetProperty("snapshots").GetArrayLength());
    }

    [Fact]
    public async Task Restore_UnknownNumber_IsInvalidArgsWithExistingNumbersAsCandidates()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));
        ws.WriteFile("test.docx", "original");

        var envelope = await srv.CallAsync("file_snapshot",
            new Dictionary<string, object?> { ["file"] = "test.docx", ["action"] = "restore", ["n"] = 99 });
        EnvelopeAssert.Fail(envelope, "invalid_args");
    }

    [Fact]
    public async Task UnknownAction_IsInvalidArgsWithCandidates()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));
        ws.WriteFile("test.docx", "original");

        var envelope = await srv.CallAsync("file_snapshot",
            new Dictionary<string, object?> { ["file"] = "test.docx", ["action"] = "rewind" });
        var error = EnvelopeAssert.Fail(envelope, "invalid_args");
        Assert.Equal(["list", "restore"],
            error.GetProperty("candidates").EnumerateArray().Select(c => c.GetString()).ToList());
    }
}
