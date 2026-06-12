using AIOffice.Core;
using AIOffice.Mcp;
using Xunit;

namespace AIOffice.Mcp.Tests;

public sealed class HandlerDiscoveryTests
{
    private static readonly string[] AllowedExtensions = [".docx", ".xlsx", ".pptx"];

    [Fact]
    public void DefaultRegistry_DiscoversAllThreeFormatHandlers()
    {
        // AIOffice.Mcp references the three format packages directly, so all
        // three handlers must be discovered — and nothing else invented.
        var registry = HandlerDiscovery.CreateDefaultRegistry();
        Assert.All(registry.KnownExtensions, e => Assert.Contains(e, AllowedExtensions));
        foreach (var extension in AllowedExtensions)
        {
            Assert.Contains(extension, registry.KnownExtensions, StringComparer.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void SnapshotInjection_MarksEveryDiscoveredKindHandlerManaged()
    {
        using var ws = new TempWorkspace();
        var snapshots = new SnapshotStore(ws.SnapshotDir);

        _ = HandlerDiscovery.CreateDefaultRegistry(snapshots, out var managed);

        // All three shipped handlers accept a SnapshotStore, so all three own
        // their pre-image snapshotting once one is injected.
        Assert.Equal(3, managed.Count);
        Assert.Contains(DocumentKind.Docx, managed);
        Assert.Contains(DocumentKind.Xlsx, managed);
        Assert.Contains(DocumentKind.Pptx, managed);
    }

    [Theory]
    [InlineData("deck.pptx", "pptx")]
    [InlineData("book.xlsx", "xlsx")]
    public async Task OfficeCreate_SucceedsThroughDiscoveredHandlers(string file, string kind)
    {
        using var ws = new TempWorkspace();
        var registry = HandlerDiscovery.CreateDefaultRegistry(new SnapshotStore(ws.SnapshotDir), out var managed);
        await using var srv = await McpTestServer.StartAsync(ws.NewService(registry, managed));

        // Regression guard: discovery must accept the shipped handlers' ctor
        // shape (a single optional SnapshotStore) — a parameterless-only
        // convention would leave the registry empty and every office_* tool
        // answering unsupported_feature.
        var data = EnvelopeAssert.Ok(await srv.CallAsync("office_create",
            new Dictionary<string, object?> { ["file"] = file }));

        Assert.Equal(kind, data.GetProperty("kind").GetString());
        Assert.True(File.Exists(Path.Combine(ws.Root, file)), $"{file} must exist after office_create");
    }

    [Fact]
    public async Task OfficeEdit_HandlerManagedKind_SnapshotsExactlyOnce()
    {
        using var ws = new TempWorkspace();
        var snapshots = new SnapshotStore(ws.SnapshotDir);
        var registry = HandlerDiscovery.CreateDefaultRegistry(snapshots, out var managed);
        await using var srv = await McpTestServer.StartAsync(ws.NewService(registry, managed));

        EnvelopeAssert.Ok(await srv.CallAsync("office_create",
            new Dictionary<string, object?> { ["file"] = "book.xlsx" }));
        EnvelopeAssert.Ok(await srv.CallAsync("office_edit", new Dictionary<string, object?>
        {
            ["file"] = "book.xlsx",
            ["ops"] = new[]
            {
                new Dictionary<string, object?>
                {
                    ["op"] = "set",
                    ["path"] = "/Sheet1/A1",
                    ["props"] = new Dictionary<string, object?> { ["value"] = 1 },
                },
            },
        }));

        // The handler received the store and snapshots the pre-image itself;
        // the service must not add a second copy of the same edit to the ring.
        Assert.Single(snapshots.List(ws.Workspace.Resolve("book.xlsx")));
    }

    [Fact]
    public async Task RealDocxHandler_AnswersWithTypedEnvelopesOnly()
    {
        using var ws = new TempWorkspace();
        var registry = HandlerDiscovery.CreateDefaultRegistry(new SnapshotStore(ws.SnapshotDir), out var managed);
        await using var srv = await McpTestServer.StartAsync(ws.NewService(registry, managed));

        // Whatever the real handler supports today, the contract is: a typed
        // envelope (ok or coded error with suggestion), never internal_error.
        var create = await srv.CallAsync("office_create",
            new Dictionary<string, object?> { ["file"] = "real.docx", ["title"] = "Discovery smoke" });
        AssertTypedEnvelope(create);

        if (create.GetProperty("ok").GetBoolean())
        {
            var edit = await srv.CallAsync("office_edit", new Dictionary<string, object?>
            {
                ["file"] = "real.docx",
                ["ops"] = new[]
                {
                    new Dictionary<string, object?>
                    {
                        ["op"] = "set",
                        ["path"] = "/body/p[1]",
                        ["props"] = new Dictionary<string, string> { ["text"] = "Hello, real handler" },
                    },
                },
            });
            AssertTypedEnvelope(edit);
        }
    }

    private static void AssertTypedEnvelope(System.Text.Json.JsonElement envelope)
    {
        if (envelope.GetProperty("ok").GetBoolean())
        {
            return;
        }

        var error = envelope.GetProperty("error");
        Assert.NotEqual("internal_error", error.GetProperty("code").GetString());
        Assert.False(string.IsNullOrWhiteSpace(error.GetProperty("suggestion").GetString()));
    }
}
