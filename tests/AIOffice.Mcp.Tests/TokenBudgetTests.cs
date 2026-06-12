using System.Text.Json;
using AIOffice.Mcp;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// docs/MCP.md §2.1: the always-resident tool surface is a context tax on every
/// agent. The whole serialized surface (name + description + inputSchema per
/// tool) must stay ≤ 3500 tokens, estimated as chars/4.
/// </summary>
public sealed class TokenBudgetTests
{
    private const double BudgetTokens = 3500;

    [Fact]
    public void WholeToolSurface_StaysInsideTheTokenBudget()
    {
        var totalChars = ToolCatalog.Tools.Sum(t =>
            t.Name.Length +
            (t.Description?.Length ?? 0) +
            JsonSerializer.Serialize(t.InputSchema).Length);

        var estimatedTokens = totalChars / 4.0;
        Assert.True(
            estimatedTokens <= BudgetTokens,
            $"Serialized tool surface is ~{estimatedTokens:F0} tokens (budget {BudgetTokens}). Move details into office_help.");
    }

    [Fact]
    public void EverySchema_IsACompactObjectSchema()
    {
        Assert.All(ToolCatalog.Tools, t =>
        {
            Assert.Equal(JsonValueKind.Object, t.InputSchema.ValueKind);
            Assert.Equal("object", t.InputSchema.GetProperty("type").GetString());
        });
    }
}
