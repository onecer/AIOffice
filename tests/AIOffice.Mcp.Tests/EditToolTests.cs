using AIOffice.Core;
using Xunit;

namespace AIOffice.Mcp.Tests;

public sealed class EditToolTests
{
    private static Dictionary<string, object?> SetTextOps(string file, string text) => new()
    {
        ["file"] = file,
        ["ops"] = new[]
        {
            new Dictionary<string, object?>
            {
                ["op"] = "set",
                ["path"] = "/body/p[1]",
                ["props"] = new Dictionary<string, string> { ["text"] = text },
            },
        },
    };

    [Fact]
    public async Task Edit_HappyPath_WritesFile_Snapshots_AndStampsRev()
    {
        using var ws = new TempWorkspace();
        var service = ws.NewService(new FakeDocxHandler());
        await using var srv = await McpTestServer.StartAsync(service);
        ws.WriteFile("test.docx", "original");

        var envelope = await srv.CallAsync("office_edit", SetTextOps("test.docx", "Hello from MCP"));

        var data = EnvelopeAssert.Ok(envelope);
        Assert.Equal(1, data.GetProperty("applied").GetInt32());
        Assert.Equal(1, data.GetProperty("snapshot").GetInt32());
        Assert.False(data.GetProperty("dryRun").GetBoolean());
        Assert.True(data.GetProperty("results")[0].GetProperty("ok").GetBoolean());

        // The mutation really happened, through the same handler layer the CLI uses.
        Assert.Equal("Hello from MCP", ws.ReadFile("test.docx"));

        // meta stamps the user-supplied file name and the post-edit rev.
        var meta = envelope.GetProperty("meta");
        Assert.Equal("test.docx", meta.GetProperty("file").GetString());
        var resolved = service.Workspace.Resolve("test.docx");
        Assert.Equal(Rev.OfFile(resolved), meta.GetProperty("rev").GetString());

        // The pre-image landed in the snapshot ring.
        var ring = new SnapshotStore(ws.SnapshotDir).List(resolved);
        var entry = Assert.Single(ring);
        Assert.Equal(1, entry.Number);
        Assert.Equal("original", File.ReadAllText(entry.Path));
    }

    [Fact]
    public async Task Edit_ExpectRevMismatch_FailsBeforeAnyWrite()
    {
        using var ws = new TempWorkspace();
        var service = ws.NewService(new FakeDocxHandler());
        await using var srv = await McpTestServer.StartAsync(service);
        ws.WriteFile("test.docx", "original");

        var args = SetTextOps("test.docx", "clobbered");
        args["expect_rev"] = "000000000000";
        var envelope = await srv.CallAsync("office_edit", args);

        EnvelopeAssert.Fail(envelope, "stale_address");
        Assert.Equal("original", ws.ReadFile("test.docx"));
        Assert.Empty(new SnapshotStore(ws.SnapshotDir).List(service.Workspace.Resolve("test.docx")));
    }

    [Fact]
    public async Task Edit_MatchingExpectRev_Succeeds()
    {
        using var ws = new TempWorkspace();
        var service = ws.NewService(new FakeDocxHandler());
        await using var srv = await McpTestServer.StartAsync(service);
        ws.WriteFile("test.docx", "original");

        var args = SetTextOps("test.docx", "guarded write");
        args["expect_rev"] = Rev.OfFile(service.Workspace.Resolve("test.docx"));
        var envelope = await srv.CallAsync("office_edit", args);

        EnvelopeAssert.Ok(envelope);
        Assert.Equal("guarded write", ws.ReadFile("test.docx"));
    }

    [Fact]
    public async Task Edit_DryRun_ValidatesWithoutWritingOrSnapshotting()
    {
        using var ws = new TempWorkspace();
        var service = ws.NewService(new FakeDocxHandler());
        await using var srv = await McpTestServer.StartAsync(service);
        ws.WriteFile("test.docx", "original");

        var args = SetTextOps("test.docx", "never written");
        args["dry_run"] = true;
        var envelope = await srv.CallAsync("office_edit", args);

        var data = EnvelopeAssert.Ok(envelope);
        Assert.True(data.GetProperty("dryRun").GetBoolean());
        Assert.False(data.TryGetProperty("snapshot", out var snapshot) && snapshot.ValueKind != System.Text.Json.JsonValueKind.Null,
            "dry run must not create a snapshot");
        Assert.Equal("original", ws.ReadFile("test.docx"));
        Assert.Empty(new SnapshotStore(ws.SnapshotDir).List(service.Workspace.Resolve("test.docx")));
    }

    [Fact]
    public async Task Edit_BadOpKind_IsInvalidArgsWithCandidates()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));
        ws.WriteFile("test.docx", "original");

        var envelope = await srv.CallAsync("office_edit", new Dictionary<string, object?>
        {
            ["file"] = "test.docx",
            ["ops"] = new[]
            {
                new Dictionary<string, object?> { ["op"] = "explode", ["path"] = "/body/p[1]" },
            },
        });

        var error = EnvelopeAssert.Fail(envelope, "invalid_args");
        Assert.Contains("set", error.GetProperty("candidates").EnumerateArray().Select(c => c.GetString()));
        Assert.Equal("original", ws.ReadFile("test.docx"));
    }

    [Fact]
    public async Task Edit_MissingOps_IsInvalidArgs()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));
        ws.WriteFile("test.docx", "original");

        var envelope = await srv.CallAsync("office_edit", new Dictionary<string, object?> { ["file"] = "test.docx" });
        EnvelopeAssert.Fail(envelope, "invalid_args");
    }

    [Fact]
    public async Task Edit_MissingFile_IsFileNotFound()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var envelope = await srv.CallAsync("office_edit", SetTextOps("ghost.docx", "x"));
        EnvelopeAssert.Fail(envelope, "file_not_found");
    }

    [Fact]
    public async Task Edit_PathOutsideWorkspace_IsSandboxDenied()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var envelope = await srv.CallAsync("office_edit", SetTextOps("../outside.docx", "x"));
        EnvelopeAssert.Fail(envelope, "sandbox_denied");
    }

    [Fact]
    public async Task Edit_TrackAndAuthor_ReachTheHandlerContext()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));
        ws.WriteFile("test.docx", "original");

        var args = SetTextOps("test.docx", "tracked text");
        args["track"] = true;
        args["author"] = "Reviewer";

        var data = EnvelopeAssert.Ok(await srv.CallAsync("office_edit", args));

        // The fake handler echoes what it saw in ctx.Args — proves the M2 wiring.
        Assert.True(data.GetProperty("track").GetBoolean());
        Assert.Equal("Reviewer", data.GetProperty("author").GetString());
    }

    [Fact]
    public async Task Edit_WithoutTrack_DefaultsToUntracked()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));
        ws.WriteFile("test.docx", "original");

        var previousAuthor = Environment.GetEnvironmentVariable("AIOFFICE_AUTHOR");
        Environment.SetEnvironmentVariable("AIOFFICE_AUTHOR", null);
        try
        {
            var data = EnvelopeAssert.Ok(await srv.CallAsync("office_edit", SetTextOps("test.docx", "plain")));

            Assert.False(data.GetProperty("track").GetBoolean());
            Assert.False(data.TryGetProperty("author", out _)); // no author arg, no env -> omitted from the wire
        }
        finally
        {
            Environment.SetEnvironmentVariable("AIOFFICE_AUTHOR", previousAuthor);
        }
    }
}
