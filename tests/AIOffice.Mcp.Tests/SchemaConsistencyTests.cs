using System.Text.Json;
using AIOffice.Cli;
using AIOffice.Core;
using AIOffice.Mcp;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// Guards the 1.0 contract that the CLI introspection surface
/// (<see cref="CommandSurface"/>, behind <c>aioffice schema</c>) and the MCP
/// introspection surface (<see cref="SurfaceSchema"/> / <see cref="ToolCatalog"/>,
/// behind <c>office_schema</c> / <c>tools/list</c>) never drift on the shared
/// vocabularies an agent depends on: the verb set, the <c>read --view</c> enum,
/// the <c>edit</c> op list, and the MCP tool count. These two surfaces are
/// separate definitions by design (the CLI one carries usage/positionals/options,
/// the MCP one is token-budgeted); this test is what keeps them honest.
/// </summary>
public sealed class SchemaConsistencyTests
{
    [Fact]
    public void Cli_and_mcp_surfaces_expose_the_same_verb_set()
    {
        var cliVerbs = CommandSurface.VerbNames.ToHashSet(StringComparer.Ordinal);
        var mcpVerbs = SurfaceSchema.VerbNames.ToHashSet(StringComparer.Ordinal);

        // The CLI surface carries two extra verbs with no MCP tool: `version`
        // (the package version rides in every envelope's meta.version) and
        // `plugin` (a host-install utility that edits local host config files,
        // not a document operation). Otherwise the two verb sets are identical.
        var cliOnly = new HashSet<string>(["version", "plugin"], StringComparer.Ordinal);
        Assert.Equal(17, mcpVerbs.Count);
        Assert.Equal(19, cliVerbs.Count);
        Assert.Contains("version", cliVerbs);
        Assert.Contains("plugin", cliVerbs);
        Assert.True(mcpVerbs.IsSubsetOf(cliVerbs));
        Assert.Equal(mcpVerbs, cliVerbs.Where(v => !cliOnly.Contains(v)).ToHashSet(StringComparer.Ordinal));
    }

    [Fact]
    public void Read_view_enum_agrees_across_cli_and_mcp_and_lists_embeds()
    {
        var cliViews = CliReadViewEnum();
        var mcpViews = McpReadViewEnum();

        // The full universal + format-specific view vocabulary, in one place.
        string[] expected =
        [
            "outline", "text", "stats", "structure", "properties", "embeds",
            "revisions", "comments", "styles", "fields", "markdown", "csv",
        ];

        Assert.Equal(expected.ToHashSet(StringComparer.Ordinal), cliViews.ToHashSet(StringComparer.Ordinal));
        Assert.Equal(expected.ToHashSet(StringComparer.Ordinal), mcpViews.ToHashSet(StringComparer.Ordinal));

        // embeds is a real view on all three formats — it must be discoverable
        // from BOTH introspection surfaces (the M10 drift this test exists for).
        Assert.Contains("embeds", cliViews);
        Assert.Contains("embeds", mcpViews);
    }

    [Fact]
    public void Edit_op_vocabulary_agrees_across_cli_mcp_and_the_runtime()
    {
        // The runtime gate (EditOp.Kinds) is the source of truth.
        Assert.Contains("extract", EditOp.Kinds);
        Assert.Equal(8, EditOp.Kinds.Count);

        var cliOps = CliEditOpEnum();
        var mcpOps = McpEditOpEnum();

        foreach (var kind in EditOp.Kinds)
        {
            Assert.Contains(kind, cliOps);
            Assert.Contains(kind, mcpOps);
        }
    }

    [Fact]
    public void Mcp_tool_catalog_has_nineteen_tools_including_file_snapshot()
    {
        // 17 verb-tools + preview_mark + preview_goto (the preview verb carries
        // preview_open/selection/mark/goto, so the verb set stays 17 — see above).
        Assert.Equal(19, ToolCatalog.Names.Count);
        Assert.Contains("file_snapshot", ToolCatalog.Names);
        Assert.Contains("preview_mark", ToolCatalog.Names);
        Assert.Contains("preview_goto", ToolCatalog.Names);

        // The snapshot tool is the one deliberate exception to the office_* /
        // preview_* naming (CONTRACT §7): assert it is exactly file_snapshot, and
        // that no office_snapshot ever sneaks in, so a rename can't slip past 1.0.
        Assert.DoesNotContain("office_snapshot", ToolCatalog.Names);
    }

    // ----- helpers: pull the live enums out of each surface --------------------

    private static IReadOnlyList<string> CliReadViewEnum()
    {
        var read = CommandSurface.Find("read")!;
        var view = read.Options.Single(o => o.Name == "view");
        return view.Value!.Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    }

    private static IReadOnlyList<string> McpReadViewEnum()
    {
        var tool = ToolCatalog.Tools.Single(t => t.Name == "office_read");
        var schema = tool.InputSchema;
        var enumNode = schema.GetProperty("properties").GetProperty("view").GetProperty("enum");
        return [.. enumNode.EnumerateArray().Select(e => e.GetString()!)];
    }

    private static IReadOnlyList<string> CliEditOpEnum()
    {
        // The CLI describes the op list in the --ops option text, e.g.
        // "[{op:set|add|...|extract, path, ...}]". Pull out the op alternation.
        var edit = CommandSurface.Find("edit")!;
        var ops = edit.Options.Single(o => o.Name == "ops");
        return ExtractAlternation(ops.Description, "op:");
    }

    private static IReadOnlyList<string> McpEditOpEnum()
    {
        var tool = ToolCatalog.Tools.Single(t => t.Name == "office_edit");
        var enumNode = tool.InputSchema
            .GetProperty("properties").GetProperty("ops")
            .GetProperty("items").GetProperty("properties").GetProperty("op")
            .GetProperty("enum");
        return [.. enumNode.EnumerateArray().Select(e => e.GetString()!)];
    }

    /// <summary>Pulls a <c>a|b|c</c> alternation that follows <paramref name="marker"/> out of free text.</summary>
    private static IReadOnlyList<string> ExtractAlternation(string text, string marker)
    {
        var start = text.IndexOf(marker, StringComparison.Ordinal);
        Assert.True(start >= 0, $"marker '{marker}' not found in: {text}");
        start += marker.Length;
        var end = start;
        while (end < text.Length && (char.IsLetter(text[end]) || text[end] == '|'))
        {
            end++;
        }

        return text[start..end].Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    }
}
