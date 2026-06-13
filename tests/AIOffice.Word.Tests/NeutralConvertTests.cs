using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// M9 flagship: the format-neutral conversion surface for docx
/// (<see cref="WordHandler.ExportNeutral"/> + <see cref="WordHandler.ImportNeutral"/>).
/// The round-trip law here is structural, not byte-exact: a rich docx exported
/// to the neutral model, imported into a fresh docx, then re-exported must carry
/// the same headings, list nesting, run formatting, tables and hyperlinks.
/// </summary>
public sealed class NeutralConvertTests : WordTestBase
{
    /// <summary>Builds a content-rich docx covering every neutral block kind.</summary>
    private string RichDoc(string name = "rich.docx")
    {
        var file = CreateDoc(name, title: "Quarterly Report");

        // Structural content first: Create leaves p[1]=Heading1 title + p[2] empty.
        var build = Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Highlights","style":"Heading2"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Bold lead","bold":true,"color":"FF0000"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Apples","list":"bullet"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Pears","list":"bullet","level":1}},
              {"op":"add","path":"/body","type":"p","props":{"text":"First","list":"number"}},
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"H1"}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"H2"}},
              {"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"text":"a"}},
              {"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"text":"b"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"see site"}}
            ]
            """);
        Assert.True(build.IsOk, build.ToJson());

        // Link into the "see site" paragraph by its canonical path (query resolves it).
        var matches = Data(Handler.Query(Ctx(name, new JsonObject { ["selector"] = "p:contains('see site')" })));
        var path = matches["matches"]!.AsArray()[0]!["path"]!.GetValue<string>();
        var link = Edit(file,
            "[{\"op\":\"add\",\"path\":\"" + path +
            "\",\"type\":\"link\",\"props\":{\"text\":\"AIOffice\",\"url\":\"https://example.com/x\"}}]");
        Assert.True(link.IsOk, link.ToJson());
        return file;
    }

    private NeutralDoc Export(string file) => Handler.ExportNeutral(Ctx(file));

    /// <summary>Imports a neutral doc into a freshly created docx and returns its path.</summary>
    private string Import(NeutralDoc doc, string name = "out.docx")
    {
        var file = CreateDoc(name);
        var result = Handler.ImportNeutral(Ctx(name), doc);
        Assert.True(result.BlocksWritten > 0);
        return file;
    }

    // ------------------------------------------------------------------ export

    [Fact]
    public void Export_maps_headings_paragraphs_lists_tables_and_links()
    {
        var neutral = Export(RichDoc());

        Assert.Equal("Quarterly Report", neutral.Title);

        var headings = neutral.Blocks.Where(b => b.Kind == NeutralBlockKind.Heading).ToList();
        Assert.Contains(headings, h => h.Level == 1 && BlockText(h) == "Quarterly Report");
        Assert.Contains(headings, h => h.Level == 2 && BlockText(h) == "Highlights");

        var bold = neutral.Blocks.First(b => BlockText(b) == "Bold lead");
        var boldRun = bold.Runs!.First();
        Assert.True(boldRun.Bold);
        Assert.Equal("FF0000", boldRun.Color);

        var items = neutral.Blocks.Where(b => b.Kind == NeutralBlockKind.ListItem).ToList();
        Assert.Contains(items, i => !i.Ordered && i.Level == 0 && BlockText(i) == "Apples");
        Assert.Contains(items, i => !i.Ordered && i.Level == 1 && BlockText(i) == "Pears");
        Assert.Contains(items, i => i.Ordered && BlockText(i) == "First");

        var table = neutral.Blocks.Single(b => b.Kind == NeutralBlockKind.Table);
        Assert.Equal(2, table.Rows!.Count);
        Assert.Equal(["H1", "H2"], table.Rows[0]);
        Assert.Equal(["a", "b"], table.Rows[1]);

        var linkRun = neutral.Blocks
            .Where(b => b.Runs is not null)
            .SelectMany(b => b.Runs!)
            .First(r => r.Text == "AIOffice");
        Assert.Equal("https://example.com/x", linkRun.Href);
    }

    [Fact]
    public void Export_title_falls_back_to_first_heading1_when_core_title_absent()
    {
        var file = CreateDoc("untitled.docx"); // no title => no core Title property
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"p","props":{"text":"Doc Heading","style":"Heading1"}}]""");

        Assert.Equal("Doc Heading", Export(file).Title);
    }

    [Fact]
    public void Export_uses_standardized_core_title_and_skips_the_create_blank_paragraph()
    {
        // 'create' leaves a trailing blank paragraph behind the Heading1 it writes;
        // it must not leak into the neutral model as a leading/junk empty block.
        var file = CreateDoc("titled.docx", title: "Q3 Report"); // body: [Heading1 "Q3 Report", blank p]
        Edit(file, """
            [
              {"op":"set","path":"/properties","props":{"title":"Q3 Report"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Overview","style":"Heading2"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Body text here."}}
            ]
            """);

        var neutral = Export(file);

        // The title comes from the standardized core property (docProps/core.xml).
        Assert.Equal("Q3 Report", neutral.Title);

        // The FIRST block is the first heading — no leading empty Paragraph artifact.
        Assert.Equal(NeutralBlockKind.Heading, neutral.Blocks[0].Kind);
        Assert.Equal("Q3 Report", BlockText(neutral.Blocks[0]));

        // No empty Paragraph block survived from the create-time blank paragraph.
        Assert.DoesNotContain(
            neutral.Blocks,
            b => b.Kind == NeutralBlockKind.Paragraph && BlockText(b).Trim().Length == 0);

        // The real paragraph content still round-trips.
        Assert.Contains(neutral.Blocks, b => b.Kind == NeutralBlockKind.Paragraph && BlockText(b) == "Body text here.");
    }

    // --------------------------------------------------------------- roundtrip

    [Fact]
    public void Roundtrip_through_neutral_preserves_structure_and_formatting()
    {
        var neutral = Export(RichDoc());
        var rebuilt = Import(neutral);

        AssertValidatesClean(rebuilt);

        // Re-export the rebuilt doc and compare the neutral projections structurally.
        var again = Export(rebuilt);

        Assert.Equal("Quarterly Report", again.Title);

        var headings = again.Blocks.Where(b => b.Kind == NeutralBlockKind.Heading).Select(BlockText).ToList();
        Assert.Contains("Quarterly Report", headings);
        Assert.Contains("Highlights", headings);

        var bold = again.Blocks.First(b => BlockText(b) == "Bold lead").Runs!.First();
        Assert.True(bold.Bold);
        Assert.Equal("FF0000", bold.Color);

        var items = again.Blocks.Where(b => b.Kind == NeutralBlockKind.ListItem).ToList();
        Assert.Contains(items, i => !i.Ordered && i.Level == 0 && BlockText(i) == "Apples");
        Assert.Contains(items, i => !i.Ordered && i.Level == 1 && BlockText(i) == "Pears");
        Assert.Contains(items, i => i.Ordered && BlockText(i) == "First");

        var table = again.Blocks.Single(b => b.Kind == NeutralBlockKind.Table);
        Assert.Equal(["H1", "H2"], table.Rows![0]);
        Assert.Equal(["a", "b"], table.Rows[1]);

        var linkRun = again.Blocks
            .Where(b => b.Runs is not null)
            .SelectMany(b => b.Runs!)
            .First(r => r.Text == "AIOffice");
        Assert.Equal("https://example.com/x", linkRun.Href);
    }

    [Fact]
    public void Import_preserves_italic_and_underline_runs()
    {
        var doc = new NeutralDoc("Styled", [
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [
                new NeutralRun("plain "),
                new NeutralRun("emph", Italic: true),
                new NeutralRun(" and "),
                new NeutralRun("under", Underline: true),
            ]),
        ]);

        var file = Import(doc, "styled.docx");
        AssertValidatesClean(file);

        var runs = Export(file).Blocks.First(b => b.Kind == NeutralBlockKind.Paragraph).Runs!;
        Assert.Contains(runs, r => r.Text == "emph" && r.Italic);
        Assert.Contains(runs, r => r.Text == "under" && r.Underline);
        Assert.Equal("Styled", Export(file).Title);
    }

    [Fact]
    public void Import_header_row_table_is_bold_and_repeats()
    {
        var doc = new NeutralDoc(null, [
            new NeutralBlock(
                NeutralBlockKind.Table,
                Rows: [["Name", "Qty"], ["Apples", "12"]],
                HeaderRow: true),
        ]);

        var file = Import(doc, "tbl.docx");
        AssertValidatesClean(file);

        // The first row must round-trip back as a HeaderRow (bold cells / tblHeader).
        var table = Export(file).Blocks.Single(b => b.Kind == NeutralBlockKind.Table);
        Assert.True(table.HeaderRow);
        Assert.Equal(["Name", "Qty"], table.Rows![0]);
    }

    // --------------------------------------------------------------- lossiness

    [Fact]
    public void Import_drops_image_with_embedded_export_token_and_records_it()
    {
        // An Image block straight off ExportNeutral carries an embedded: token, not a
        // disk path; importing it cannot recover the bytes, so it must be dropped.
        var doc = new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Image, Source: "embedded:image1", Alt: "a chart"),
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("after")]),
        ]);

        var file = CreateDoc("dropimg.docx");
        var result = Handler.ImportNeutral(Ctx("dropimg.docx"), doc);

        Assert.Equal(1, result.BlocksWritten); // only the paragraph landed
        Assert.Single(result.Dropped);
        Assert.Contains("a chart", result.Dropped[0]);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Import_embeds_image_when_source_resolves_in_workspace()
    {
        var png = WritePng("logo.png", width: 8, height: 6);
        var doc = new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Image, Source: png, Alt: "the logo"),
        ]);

        var file = CreateDoc("withimg.docx");
        var result = Handler.ImportNeutral(Ctx("withimg.docx"), doc);

        Assert.Equal(1, result.BlocksWritten);
        Assert.Empty(result.Dropped);
        AssertValidatesClean(file);

        // The embedded image re-exports as an Image block carrying its alt text.
        var image = Export(file).Blocks.Single(b => b.Kind == NeutralBlockKind.Image);
        Assert.Equal("the logo", image.Alt);
        Assert.StartsWith("embedded:", image.Source);
    }

    [Fact]
    public void Import_drops_missing_image_source_without_failing()
    {
        var doc = new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Image, Source: "nope.png", Alt: "ghost"),
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("body")]),
        ]);

        var file = CreateDoc("missimg.docx");
        var result = Handler.ImportNeutral(Ctx("missimg.docx"), doc);

        Assert.Equal(1, result.BlocksWritten);
        Assert.Single(result.Dropped);
        AssertValidatesClean(file);
        Assert.Contains("body", BodyTexts(file));
    }

    // -------------------------------------------------------------- determinism

    [Fact]
    public void Import_is_deterministic_across_runs()
    {
        var neutral = Export(RichDoc("seed.docx"));

        var a = Import(neutral, "a.docx");
        var b = Import(neutral, "b.docx");

        // Same neutral input → identical body text projection on every run/platform.
        Assert.Equal(NormalizedBody(a), NormalizedBody(b));
    }

    [Fact]
    public void Import_into_empty_doc_clears_placeholder_paragraph()
    {
        var doc = new NeutralDoc("One", [
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("only line")]),
        ]);

        var file = Import(doc, "single.docx");
        var lines = BodyTexts(file).Where(l => l.Length > 0).ToList();
        Assert.Equal(["only line"], lines);
    }

    [Fact]
    public void Import_internal_anchor_href_becomes_anchor_link()
    {
        var doc = new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [
                new NeutralRun("jump", Href: "#section2"),
            ]),
        ]);

        var file = Import(doc, "anchor.docx");
        AssertValidatesClean(file);

        var run = Export(file).Blocks.First(b => b.Kind == NeutralBlockKind.Paragraph).Runs!.Single();
        Assert.Equal("#section2", run.Href);
    }

    // ----------------------------------------------------------------- helpers

    private static string BlockText(NeutralBlock block) =>
        block.Runs is { } runs ? string.Concat(runs.Select(r => r.Text)) : string.Empty;

    private string NormalizedBody(string file) =>
        string.Join("\n", BodyTexts(file)).Replace("\r\n", "\n", StringComparison.Ordinal);
}
