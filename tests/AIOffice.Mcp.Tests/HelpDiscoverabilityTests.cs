using AIOffice.Mcp;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// Drift guard: office_help must SERVE the authoritative settable-property
/// references (properties-docx/xlsx/pptx). Their bodies are the .md files owned
/// on-disk by AIOffice.Cli and LINKED into AIOffice.Mcp as embedded resources
/// (single source of truth). These tests fail if the link is dropped/renamed, if
/// serving breaks, or if a future milestone's vocabulary silently stops flowing
/// through to the MCP surface — they assert the CURRENT vocabulary is present.
/// </summary>
public sealed class HelpDiscoverabilityTests
{
    [Fact]
    public async Task OfficeHelp_ServesPropertiesPptx_WithTheCurrent3DAndEffectVocabulary()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var data = EnvelopeAssert.Ok(await srv.CallAsync("office_help",
            new Dictionary<string, object?> { ["topic"] = "properties-pptx" }));

        Assert.Equal("properties-pptx", data.GetProperty("topic").GetString());
        var doc = data.GetProperty("doc").GetString()!;
        // Vocabulary added across 1.23–1.26 — this is the whole point of the link.
        Assert.Contains("bevel", doc, StringComparison.Ordinal);
        Assert.Contains("scene3d", doc, StringComparison.Ordinal);
        Assert.Contains("innerShadow", doc, StringComparison.Ordinal);
        Assert.Contains("softEdge", doc, StringComparison.Ordinal);
        // It is the FULL reference, not a stub.
        Assert.True(doc.Length > 5000, $"properties-pptx served only {doc.Length} chars — link likely broken");
        Assert.True(data.GetProperty("related").GetArrayLength() > 0);
    }

    [Fact]
    public async Task OfficeHelp_ServesPropertiesXlsx_WithTabColorAndMargins()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var data = EnvelopeAssert.Ok(await srv.CallAsync("office_help",
            new Dictionary<string, object?> { ["topic"] = "properties-xlsx" }));

        var doc = data.GetProperty("doc").GetString()!;
        Assert.Contains("tabColor", doc, StringComparison.Ordinal);
        Assert.Contains("margins", doc, StringComparison.Ordinal);
        Assert.True(doc.Length > 5000, $"properties-xlsx served only {doc.Length} chars — link likely broken");
    }

    [Fact]
    public async Task OfficeHelp_ServesPropertiesDocx_WithEmphasisMark()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var data = EnvelopeAssert.Ok(await srv.CallAsync("office_help",
            new Dictionary<string, object?> { ["topic"] = "properties-docx" }));

        var doc = data.GetProperty("doc").GetString()!;
        Assert.Contains("emphasisMark", doc, StringComparison.Ordinal);
        Assert.Contains("underline", doc, StringComparison.Ordinal);
        Assert.True(doc.Length > 5000, $"properties-docx served only {doc.Length} chars — link likely broken");
    }

    [Fact]
    public async Task OfficeHelp_Index_ListsThePropertiesTopics()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        var index = EnvelopeAssert.Ok(await srv.CallAsync("office_help"));
        var doc = index.GetProperty("doc").GetString()!;
        Assert.Contains("properties-docx", doc, StringComparison.Ordinal);
        Assert.Contains("properties-xlsx", doc, StringComparison.Ordinal);
        Assert.Contains("properties-pptx", doc, StringComparison.Ordinal);
    }

    [Fact]
    public async Task OfficeHelp_RefreshedEffectTopics_AreNoLongerStuckAt1Point1()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        // pptx/effect must now mention the newer effects AND cross-reference the full list.
        var pptxEffect = EnvelopeAssert.Ok(await srv.CallAsync("office_help",
            new Dictionary<string, object?> { ["topic"] = "pptx/effect" }));
        var pe = pptxEffect.GetProperty("doc").GetString()!;
        Assert.Contains("scene3d", pe, StringComparison.Ordinal);
        Assert.Contains("bevel", pe, StringComparison.Ordinal);
        Assert.Contains("properties-pptx", pe, StringComparison.Ordinal);

        // docx/effect must cross-reference the full run vocabulary.
        var docxEffect = EnvelopeAssert.Ok(await srv.CallAsync("office_help",
            new Dictionary<string, object?> { ["topic"] = "docx/effect" }));
        Assert.Contains("properties-docx", docxEffect.GetProperty("doc").GetString(), StringComparison.Ordinal);

        // pptx/shape must point at the complete property reference.
        var pptxShape = EnvelopeAssert.Ok(await srv.CallAsync("office_help",
            new Dictionary<string, object?> { ["topic"] = "pptx/shape" }));
        Assert.Contains("properties-pptx", pptxShape.GetProperty("doc").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task OfficeHelp_PropertiesTopics_AreCandidatesForNearMisses()
    {
        using var ws = new TempWorkspace();
        await using var srv = await McpTestServer.StartAsync(ws.NewService(new FakeDocxHandler()));

        // "properties" alone is not a topic; the three references must surface as candidates.
        var envelope = await srv.CallAsync("office_help",
            new Dictionary<string, object?> { ["topic"] = "properties" });
        var error = EnvelopeAssert.Fail(envelope, "invalid_args");
        var candidates = error.GetProperty("candidates").EnumerateArray().Select(c => c.GetString()).ToList();
        Assert.Contains("properties-docx", candidates);
        Assert.Contains("properties-xlsx", candidates);
        Assert.Contains("properties-pptx", candidates);
    }
}
