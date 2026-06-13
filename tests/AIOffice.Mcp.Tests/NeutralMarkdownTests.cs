using AIOffice.Core;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// The command-layer <see cref="NeutralMarkdown"/> serializer that lets any
/// office format reach markdown through the neutral model: writing a
/// <see cref="NeutralDoc"/> to GFM and parsing GFM back must preserve the block
/// structure and inline formatting markdown can carry.
/// </summary>
public sealed class NeutralMarkdownTests
{
    [Fact]
    public void Write_renders_each_block_kind()
    {
        var doc = new NeutralDoc("Title", [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Heading One")]),
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [
                new NeutralRun("plain "),
                new NeutralRun("bold", Bold: true),
                new NeutralRun(" and "),
                new NeutralRun("link", Href: "https://example.com"),
            ]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 0, Runs: [new NeutralRun("first")]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 1, Ordered: true, Runs: [new NeutralRun("nested")]),
            new NeutralBlock(NeutralBlockKind.Table, HeaderRow: true, Rows: [["A", "B"], ["1", "2"]]),
            new NeutralBlock(NeutralBlockKind.Image, Source: "pic.png", Alt: "a picture"),
        ]);

        var md = NeutralMarkdown.Write(doc);

        Assert.Contains("# Heading One", md, StringComparison.Ordinal);
        Assert.Contains("**bold**", md, StringComparison.Ordinal);
        Assert.Contains("[link](https://example.com)", md, StringComparison.Ordinal);
        Assert.Contains("- first", md, StringComparison.Ordinal);
        Assert.Contains("  1. nested", md, StringComparison.Ordinal);
        Assert.Contains("| A | B |", md, StringComparison.Ordinal);
        Assert.Contains("| --- | --- |", md, StringComparison.Ordinal);
        Assert.Contains("![a picture](pic.png)", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Roundtrip_preserves_headings_lists_table_and_inline()
    {
        var original = new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 2, Runs: [new NeutralRun("Section")]),
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [
                new NeutralRun("an "),
                new NeutralRun("italic", Italic: true),
                new NeutralRun(" word"),
            ]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 0, Runs: [new NeutralRun("alpha")]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 1, Runs: [new NeutralRun("beta")]),
            new NeutralBlock(NeutralBlockKind.Table, HeaderRow: true, Rows: [["Name", "Qty"], ["Apple", "3"]]),
        ]);

        var parsed = NeutralMarkdown.Parse(NeutralMarkdown.Write(original));

        var heading = parsed.Blocks.Single(b => b.Kind == NeutralBlockKind.Heading);
        Assert.Equal(2, heading.Level);
        Assert.Equal("Section", Text(heading));

        var paragraph = parsed.Blocks.First(b => b.Kind == NeutralBlockKind.Paragraph);
        Assert.Contains(paragraph.Runs!, r => r.Text == "italic" && r.Italic);

        var items = parsed.Blocks.Where(b => b.Kind == NeutralBlockKind.ListItem).ToList();
        Assert.Contains(items, i => Text(i) == "alpha" && i.Level == 0);
        Assert.Contains(items, i => Text(i) == "beta" && i.Level == 1);

        var table = parsed.Blocks.Single(b => b.Kind == NeutralBlockKind.Table);
        Assert.True(table.HeaderRow);
        Assert.Equal(["Name", "Qty"], table.Rows![0]);
        Assert.Equal(["Apple", "3"], table.Rows![1]);
    }

    [Fact]
    public void Parse_first_h1_becomes_the_title()
    {
        var parsed = NeutralMarkdown.Parse("# Doc Title\n\nbody text\n");
        Assert.Equal("Doc Title", parsed.Title);
    }

    [Fact]
    public void Write_emits_the_title_as_a_leading_h1_when_it_is_not_already_the_first_heading()
    {
        // The title is a distinct core property (here "Quarterly Report"), separate
        // from the first body heading ("Sales"); Write must not drop it.
        var doc = new NeutralDoc("Quarterly Report", [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Sales")]),
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("body")]),
        ]);

        var md = NeutralMarkdown.Write(doc);

        Assert.StartsWith("# Quarterly Report\n", md, StringComparison.Ordinal);
        Assert.Contains("# Sales", md, StringComparison.Ordinal);
        // The emitted title round-trips: Parse reads it back off the leading H1.
        Assert.Equal("Quarterly Report", NeutralMarkdown.Parse(md).Title);
    }

    [Fact]
    public void Write_does_not_duplicate_a_title_that_is_already_the_first_heading()
    {
        // When the title already opens the document as its first H1, Write must not
        // emit a second copy — otherwise every round trip would grow a duplicate.
        var doc = new NeutralDoc("Sales", [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Sales")]),
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("body")]),
        ]);

        var md = NeutralMarkdown.Write(doc);

        // Exactly one "# Sales" line, and Write is a fixed point under re-Write.
        Assert.Equal(1, md.Split('\n').Count(l => l == "# Sales"));
        Assert.Equal(md, NeutralMarkdown.Write(NeutralMarkdown.Parse(md)));
    }

    [Fact]
    public void Write_without_a_title_emits_no_leading_heading()
    {
        var md = NeutralMarkdown.Write(new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("just a paragraph")]),
        ]));

        Assert.StartsWith("just a paragraph", md, StringComparison.Ordinal);
        Assert.DoesNotContain("#", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Parse_headerless_table_keeps_headerrow_false()
    {
        // A GFM table whose header row is blank parses back as a headerless table.
        var md = "|  |  |\n| --- | --- |\n| x | y |\n";
        var table = NeutralMarkdown.Parse(md).Blocks.Single(b => b.Kind == NeutralBlockKind.Table);
        Assert.False(table.HeaderRow);
        Assert.Equal(["x", "y"], table.Rows![0]);
    }

    [Fact]
    public void Write_uses_lf_line_endings_only()
    {
        var md = NeutralMarkdown.Write(new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("X")]),
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("y")]),
        ]));

        Assert.DoesNotContain('\r', md);
    }

    private static string Text(NeutralBlock block) =>
        block.Runs is { } runs ? string.Concat(runs.Select(r => r.Text)) : string.Empty;
}
