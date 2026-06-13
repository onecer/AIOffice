using AIOffice.Core;
using AIOffice.Mcp;
using Xunit;

namespace AIOffice.Mcp.Tests;

public sealed class SchemaHelpStatusTests
{
    [Fact]
    public async Task OfficeSchema_ReturnsTheFullSixteenVerbSurface()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var data = EnvelopeAssert.Ok(await srv.CallAsync("office_schema"));
        Assert.Equal(Meta.ToolVersion, data.GetProperty("version").GetString());

        var verbs = data.GetProperty("verbs").EnumerateArray().ToList();
        Assert.Equal(17, verbs.Count); // M1 preview; M7 audit; M8 diff; M9 convert
        Assert.Equal(
            SurfaceSchema.VerbNames.ToHashSet(StringComparer.Ordinal),
            verbs.Select(v => v.GetProperty("name").GetString()!).ToHashSet(StringComparer.Ordinal));

        Assert.All(verbs, v =>
        {
            Assert.False(string.IsNullOrWhiteSpace(v.GetProperty("summary").GetString()));
            Assert.True(v.GetProperty("errors").GetArrayLength() > 0);
            Assert.True(v.GetProperty("examples").GetArrayLength() > 0);
        });
    }

    [Fact]
    public async Task OfficeSchema_FiltersToOneVerb_AndMapsItToItsMcpTool()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var data = EnvelopeAssert.Ok(await srv.CallAsync("office_schema",
            new Dictionary<string, object?> { ["verb"] = "edit" }));
        var verb = Assert.Single(data.GetProperty("verbs").EnumerateArray().ToList());
        Assert.Equal("edit", verb.GetProperty("name").GetString());
        Assert.Equal("office_edit", verb.GetProperty("mcpTool").GetString());
    }

    [Fact]
    public async Task OfficeSchema_UnknownVerb_ListsAllVerbsAsCandidates()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var envelope = await srv.CallAsync("office_schema", new Dictionary<string, object?> { ["verb"] = "transmogrify" });
        var error = EnvelopeAssert.Fail(envelope, "invalid_args");
        Assert.Equal(17, error.GetProperty("candidates").GetArrayLength());
    }

    [Fact]
    public async Task OfficeHelp_IndexAndTopics_Work()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var index = EnvelopeAssert.Ok(await srv.CallAsync("office_help"));
        Assert.Equal("index", index.GetProperty("topic").GetString());
        Assert.Contains("addressing", index.GetProperty("doc").GetString(), StringComparison.Ordinal);

        var addressing = EnvelopeAssert.Ok(await srv.CallAsync("office_help",
            new Dictionary<string, object?> { ["topic"] = "addressing" }));
        Assert.Contains("/body/p[3]", addressing.GetProperty("doc").GetString(), StringComparison.Ordinal);
        Assert.True(addressing.GetProperty("related").GetArrayLength() > 0);

        var paragraph = EnvelopeAssert.Ok(await srv.CallAsync("office_help",
            new Dictionary<string, object?> { ["topic"] = "docx/paragraph" }));
        Assert.Contains("style", paragraph.GetProperty("doc").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task OfficeHelp_UnknownTopic_SuggestsCandidates()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var envelope = await srv.CallAsync("office_help", new Dictionary<string, object?> { ["topic"] = "docx" });
        var error = EnvelopeAssert.Fail(envelope, "invalid_args");
        Assert.Contains("docx/paragraph",
            error.GetProperty("candidates").EnumerateArray().Select(c => c.GetString()));
    }

    [Fact]
    public async Task OfficeStatus_ReportsHealthyWorkspaceAndChecks()
    {
        using var ws = new TempWorkspace();
        var service = ws.NewService(new FakeDocxHandler());
        await using var srv = await McpTestServer.StartAsync(service);

        var data = EnvelopeAssert.Ok(await srv.CallAsync("office_status"));
        Assert.True(data.GetProperty("healthy").GetBoolean());
        Assert.Equal(service.Workspace.Root, data.GetProperty("workspace").GetString());
        Assert.Equal(Meta.ToolVersion, data.GetProperty("version").GetString());

        var checks = data.GetProperty("checks").EnumerateArray().ToList();
        Assert.Equal(3, checks.Count);
        Assert.All(checks, c => Assert.True(c.GetProperty("ok").GetBoolean(), c.GetRawText()));
    }

    [Fact]
    public async Task OfficeCreate_WritesThroughTheHandler_AndRefusesSilentOverwrite()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var data = EnvelopeAssert.Ok(await srv.CallAsync("office_create",
            new Dictionary<string, object?> { ["file"] = "out/new.docx", ["title"] = "T" }));
        Assert.Equal("docx", data.GetProperty("kind").GetString());
        Assert.Equal("blank:T", ws.ReadFile(Path.Combine("out", "new.docx")));

        var again = await srv.CallAsync("office_create", new Dictionary<string, object?> { ["file"] = "out/new.docx" });
        EnvelopeAssert.Fail(again, "invalid_args");
    }

    [Fact]
    public async Task UnregisteredFormat_IsTypedUnsupportedFeature_NeverACrash()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var envelope = await srv.CallAsync("office_create", new Dictionary<string, object?> { ["file"] = "data.xlsx" });
        var error = EnvelopeAssert.Fail(envelope, "unsupported_feature");
        Assert.Contains(".docx", error.GetProperty("suggestion").GetString(), StringComparison.Ordinal);
    }
}
