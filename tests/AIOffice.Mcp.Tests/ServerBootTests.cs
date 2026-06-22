using Xunit;

namespace AIOffice.Mcp.Tests;

public sealed class ServerBootTests
{
    private static readonly string[] ExpectedTools =
    [
        "office_create", "office_read", "office_query", "office_get",
        "office_edit", "office_render", "office_validate", "office_audit", "office_diff", "office_template",
        "office_convert", "file_snapshot", "office_status", "office_help", "office_schema",
        "preview_open", "preview_selection", "preview_mark", "preview_goto",
    ];

    [Fact]
    public async Task Boot_AdvertisesAioffice_AndExactlyTheNineteenTools()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        Assert.Equal("aioffice", srv.Client.ServerInfo.Name);

        var tools = await srv.Client.ListToolsAsync();
        Assert.Equal(ExpectedTools.Length, tools.Count);
        Assert.Equal(19, tools.Count); // 17 verb-tools + preview_mark/preview_goto (preview watch parity)
        Assert.Equal(ExpectedTools.ToHashSet(StringComparer.Ordinal),
            tools.Select(t => t.Name).ToHashSet(StringComparer.Ordinal));

        Assert.All(tools, t =>
        {
            Assert.False(string.IsNullOrWhiteSpace(t.Description), $"{t.Name} has no description");
            Assert.Equal("object", t.JsonSchema.GetProperty("type").GetString());
        });
    }

    [Fact]
    public async Task UnknownTool_ReturnsInvalidArgsEnvelope_WithToolNameCandidates()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var raw = await srv.CallRawAsync("office_explode");
        Assert.True(raw.IsError);

        var envelope = await srv.CallAsync("office_explode");
        var error = EnvelopeAssert.Fail(envelope, "invalid_args");
        var candidates = error.GetProperty("candidates").EnumerateArray().Select(c => c.GetString()).ToList();
        Assert.Contains("office_edit", candidates);
        Assert.Equal(ExpectedTools.Length, candidates.Count);
    }

    [Fact]
    public async Task EveryEnvelope_CarriesVersionAndElapsedMs()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var envelope = await srv.CallAsync("office_status");
        var meta = envelope.GetProperty("meta");
        Assert.False(string.IsNullOrWhiteSpace(meta.GetProperty("version").GetString()));
        Assert.True(meta.GetProperty("elapsedMs").GetInt64() >= 0);
    }
}
