using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>M5 flagship: the markdown bridge (CreateFrom import + read --view markdown export).</summary>
public sealed class MarkdownBridgeTests : WordTestBase
{
    /// <summary>One fixture covering every supported markdown construct.</summary>
    private const string RichMarkdown = """
        ---
        title: front matter is ignored
        ---

        # Quarterly Report

        ## Highlights

        Revenue grew **strongly** with *some* caveats and ~~no~~ `inline_code` notes.

        A [link to the site](https://example.com) and <https://aioffice.dev> too.

        ### Numbers

        | Item | Qty | Price |
        | :--- | :---: | ---: |
        | Apples \| pears | 12 | 3.50 |
        | Bananas | 7 | 1.20 |

        - bullet one
        - bullet two
            - nested bullet

        1. first
        2. second
            1. second.one

        > Quoted wisdom here.

        ```
        let x = 1;
        x += 1;
        ```

        ---

        Closing paragraph.
        """;

    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    private string WriteMd(string name, string content)
    {
        var path = Path.Combine(Dir, name);
        File.WriteAllText(path, content);
        return path;
    }

    private string ImportMd(string markdown, string docName = "imported.docx")
    {
        WriteMd("source.md", markdown);
        var envelope = Handler.CreateFrom(Ctx(docName), "source.md");
        Assert.True(envelope.IsOk, envelope.ToJson());
        return Path.Combine(Dir, docName);
    }

    private string ExportMd(string file) =>
        Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "markdown" })))["markdown"]!.GetValue<string>();

    // ----------------------------------------------------------------- import

    [Fact]
    public void Rich_fixture_imports_with_validator_zero()
    {
        var file = ImportMd(RichMarkdown);

        AssertValidatesClean(file);

        // Headings land as styled paragraphs and drive the outline tree.
        var outline = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "outline" })));
        var h1 = outline["headings"]!.AsArray().Single()!;
        Assert.Equal("Quarterly Report", h1["text"]!.GetValue<string>());
        var h2 = h1["children"]!.AsArray().Single()!;
        Assert.Equal("Highlights", h2["text"]!.GetValue<string>());
        Assert.Equal("Numbers", h2["children"]!.AsArray().Single()!["text"]!.GetValue<string>());
    }

    [Fact]
    public void Import_maps_inline_formatting_to_runs()
    {
        var file = ImportMd("Revenue grew **strongly** with *some* caveats and ~~no~~ `inline_code` notes.");

        var matches = Data(Handler.Query(Ctx(file, new JsonObject { ["selector"] = "run[bold=true]" })));
        Assert.Equal("strongly", matches["matches"]!.AsArray().Single()!["snippet"]!.GetValue<string>());

        var italic = Data(Handler.Query(Ctx(file, new JsonObject { ["selector"] = "run[italic=true]" })));
        Assert.Equal("some", italic["matches"]!.AsArray().Single()!["snippet"]!.GetValue<string>());

        Assert.Equal(
            "Revenue grew strongly with some caveats and no inline_code notes.",
            Get(file, "/body/p[1]")["properties"]!["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Import_builds_real_nested_lists_on_the_numbering_machinery()
    {
        var file = ImportMd("""
            - bullet one
            - bullet two
                - nested bullet

            1. first
            2. second
            """);

        var first = Get(file, "/body/p[1]")["properties"]!;
        Assert.Equal("bullet", first["list"]!.GetValue<string>());
        Assert.Equal(0, first["level"]!.GetValue<int>());

        var nested = Get(file, "/body/p[3]")["properties"]!;
        Assert.Equal("bullet", nested["list"]!.GetValue<string>());
        Assert.Equal(1, nested["level"]!.GetValue<int>());

        var second = Get(file, "/body/p[5]")["properties"]!;
        Assert.Equal("number", second["list"]!.GetValue<string>());
        Assert.Equal(2, second["number"]!.GetValue<int>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Import_numbered_lists_restart_per_list()
    {
        var file = ImportMd("""
            1. alpha

            separator paragraph

            1. beta
            """);

        Assert.Equal(1, Get(file, "/body/p[1]")["properties"]!["number"]!.GetValue<int>());
        Assert.Equal(1, Get(file, "/body/p[3]")["properties"]!["number"]!.GetValue<int>());
    }

    [Fact]
    public void Import_builds_real_tables_with_alignment_and_bold_header()
    {
        var file = ImportMd("""
            | Item | Qty | Price |
            | :--- | :---: | ---: |
            | Apples \| pears | 12 | 3.50 |
            """);

        var table = Get(file, "/body/table[1]")["properties"]!;
        Assert.Equal(2, table["rows"]!.GetValue<int>());
        Assert.Equal(3, table["columns"]!.GetValue<int>());
        Assert.True(table["headerRow"]!.GetValue<bool>());

        Assert.True(Get(file, "/body/table[1]/tr[1]/tc[1]/p[1]")["properties"]!["bold"]!.GetValue<bool>());
        Assert.Equal("center", Get(file, "/body/table[1]/tr[2]/tc[2]/p[1]")["properties"]!["alignment"]!.GetValue<string>());
        Assert.Equal("right", Get(file, "/body/table[1]/tr[2]/tc[3]/p[1]")["properties"]!["alignment"]!.GetValue<string>());
        Assert.Equal("Apples | pears", Get(file, "/body/table[1]/tr[2]/tc[1]")["properties"]!["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Import_links_become_real_hyperlinks()
    {
        var file = ImportMd("See [the site](https://example.com) for details.");

        var link = Get(file, "/body/p[1]/link[1]")["properties"]!;
        Assert.Equal("https://example.com", link["url"]!.GetValue<string>());
        Assert.Equal("the site", link["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Import_relative_link_keeps_text_with_warning()
    {
        WriteMd("source.md", "See [other doc](./other.md).");
        var envelope = Handler.CreateFrom(Ctx("out.docx"), "source.md");

        var json = JsonNode.Parse(envelope.ToJson())!;
        var codes = json["meta"]!["warnings"]!.AsArray().Select(w => w!["code"]!.GetValue<string>()).ToList();
        Assert.Contains("md_link_skipped", codes);
        Assert.Equal("See other doc.", Get(Path.Combine(Dir, "out.docx"), "/body/p[1]")["properties"]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void Import_embeds_images_through_the_sandbox()
    {
        File.WriteAllBytes(
            Path.Combine(Dir, "dot.png"),
            Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="));
        var file = ImportMd("Before\n\n![a dot](dot.png)\n\nAfter");

        var props = Get(file, "/body/p[2]")["properties"]!;
        Assert.Equal("image", props["kind"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Import_missing_image_warns_and_does_not_fail()
    {
        WriteMd("source.md", "![missing](nope.png)\n\nStill here.");
        var envelope = Handler.CreateFrom(Ctx("out.docx"), "source.md");

        Assert.True(envelope.IsOk);
        var codes = JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"]!.AsArray()
            .Select(w => w!["code"]!.GetValue<string>()).ToList();
        Assert.Contains("md_image_skipped", codes);
        AssertValidatesClean(Path.Combine(Dir, "out.docx"));
    }

    [Fact]
    public void Import_raw_html_is_skipped_with_warning()
    {
        WriteMd("source.md", "<div>raw</div>\n\nText with <b>inline html</b> here.");
        var envelope = Handler.CreateFrom(Ctx("out.docx"), "source.md");

        Assert.True(envelope.IsOk);
        var codes = JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"]!.AsArray()
            .Select(w => w!["code"]!.GetValue<string>()).ToList();
        Assert.Contains("md_html_skipped", codes);

        // The inline html's surrounding text survives.
        Assert.Equal(
            "Text with inline html here.",
            Get(Path.Combine(Dir, "out.docx"), "/body/p[1]")["properties"]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void Import_code_block_preserves_lines_and_monospace_style()
    {
        var file = ImportMd("```\nlet x = 1;\nx += 1;\n```");

        var props = Get(file, "/body/p[1]")["properties"]!;
        Assert.Equal("Code", props["style"]!.GetValue<string>());
        Assert.Equal("let x = 1;x += 1;", props["text"]!.GetValue<string>()); // w:br between lines
        AssertValidatesClean(file);
    }

    [Fact]
    public void Import_front_matter_blockquote_and_rule_shape_correctly()
    {
        var file = ImportMd("""
            ---
            ignored: yes
            ---

            > Quoted wisdom.

            ---
            """);

        var texts = BodyTexts(file);
        Assert.Equal("Quoted wisdom.", texts[0]); // front matter gone, quote is p[1]
        Assert.Equal(2, texts.Count); // quote + the empty horizontal-rule paragraph
        AssertValidatesClean(file);
    }

    // ------------------------------------------------------ import refusals

    [Fact]
    public void Import_into_an_existing_file_is_invalid_args()
    {
        var existing = CreateDoc("exists.docx");
        WriteMd("source.md", "# Hi");

        var ex = Assert.Throws<AiofficeException>(() => Handler.CreateFrom(Ctx("exists.docx"), "source.md"));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.True(File.Exists(existing));
    }

    [Fact]
    public void Import_from_a_non_markdown_source_is_unsupported_feature()
    {
        File.WriteAllText(Path.Combine(Dir, "notes.txt"), "plain");

        var ex = Assert.Throws<AiofficeException>(() => Handler.CreateFrom(Ctx("out.docx"), "notes.txt"));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains(".md", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Import_from_a_missing_source_is_file_not_found()
    {
        var ex = Assert.Throws<AiofficeException>(() => Handler.CreateFrom(Ctx("out.docx"), "ghost.md"));

        Assert.Equal(ErrorCodes.FileNotFound, ex.Code);
    }

    [Fact]
    public void Import_source_outside_the_sandbox_is_denied()
    {
        var ex = Assert.Throws<AiofficeException>(() => Handler.CreateFrom(Ctx("out.docx"), "../escape.md"));

        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
    }

    // ----------------------------------------------------------------- export

    [Fact]
    public void Export_walks_headings_lists_links_and_bold()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"Title","style":"Heading1"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Plain with "}},
              {"op":"add","path":"/body/p[2]","type":"link","props":{"text":"a link","url":"https://example.com"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"bold words","bold":true}},
              {"op":"add","path":"/body","type":"p","props":{"text":"item one","list":"bullet"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"item two","list":"bullet"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"step one","list":"number"}}
            ]
            """);

        var md = ExportMd(file);

        Assert.Contains("# Title", md, StringComparison.Ordinal);
        Assert.Contains("[a link](https://example.com)", md, StringComparison.Ordinal);
        Assert.Contains("**bold words**", md, StringComparison.Ordinal);
        Assert.Contains("- item one\n- item two", md, StringComparison.Ordinal);
        Assert.Contains("1. step one", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Export_emits_pipe_tables_with_escaped_pipes()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"Name"}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"Qty"}},
              {"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"text":"a|b"}},
              {"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"text":"7"}}
            ]
            """);

        var md = ExportMd(file);

        Assert.Contains("| Name | Qty |", md, StringComparison.Ordinal);
        Assert.Contains("| --- | --- |", md, StringComparison.Ordinal);
        Assert.Contains("a\\|b", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Export_footnotes_use_gfm_footnote_syntax()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"Claim."}},
              {"op":"add","path":"/body/p[1]","type":"footnote","props":{"text":"the source"}}
            ]
            """);

        var md = ExportMd(file);

        Assert.Contains("Claim.[^1]", md, StringComparison.Ordinal);
        Assert.Contains("[^1]: the source", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Export_notes_omitted_headers_in_a_leading_comment()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """
            [
              {"op":"add","path":"/header[1]","type":"header","props":{"text":"Confidential"}},
              {"op":"add","path":"/body","type":"watermark","props":{"text":"DRAFT"}}
            ]
            """);

        var md = ExportMd(file);

        Assert.StartsWith("<!-- omitted: headers", md, StringComparison.Ordinal);
        Assert.Contains("watermark", md[..md.IndexOf('\n')], StringComparison.Ordinal);
        Assert.DoesNotContain("Confidential", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Export_without_headers_has_no_omission_comment()
    {
        var file = CreateDoc(title: "Doc");

        Assert.DoesNotContain("<!--", ExportMd(file), StringComparison.Ordinal);
    }

    // ------------------------------------------------------------ round trip

    /// <summary>
    /// The bridge sanity law: md -> docx -> md preserves heading, list, table,
    /// link and bold STRUCTURE (whitespace and cosmetic spelling normalized).
    /// </summary>
    [Fact]
    public void Markdown_round_trip_preserves_structure()
    {
        var file = ImportMd(RichMarkdown);
        var exported = ExportMd(file);

        // Headings survive with their levels.
        Assert.Contains("# Quarterly Report", exported, StringComparison.Ordinal);
        Assert.Contains("## Highlights", exported, StringComparison.Ordinal);
        Assert.Contains("### Numbers", exported, StringComparison.Ordinal);

        // Inline formatting survives.
        Assert.Contains("**strongly**", exported, StringComparison.Ordinal);
        Assert.Contains("*some*", exported, StringComparison.Ordinal);
        Assert.Contains("~~no~~", exported, StringComparison.Ordinal);
        Assert.Contains("`inline_code`", exported, StringComparison.Ordinal);

        // Links survive as links.
        Assert.Contains("[link to the site](https://example.com)", exported, StringComparison.Ordinal);

        // Lists survive with nesting and numbering.
        Assert.Contains("- bullet one", exported, StringComparison.Ordinal);
        Assert.Contains("    - nested bullet", exported, StringComparison.Ordinal);
        Assert.Contains("1. first", exported, StringComparison.Ordinal);
        Assert.Contains("2. second", exported, StringComparison.Ordinal);
        Assert.Contains("    1. second.one", exported, StringComparison.Ordinal);

        // The table survives as a pipe table with its alignment row.
        Assert.Contains("| Item | Qty | Price |", exported, StringComparison.Ordinal);
        Assert.Contains("| :--- | :---: | ---: |", exported, StringComparison.Ordinal);
        Assert.Contains("Apples \\| pears", exported, StringComparison.Ordinal);

        // Quote, code and rule survive.
        Assert.Contains("> Quoted wisdom here.", exported, StringComparison.Ordinal);
        Assert.Contains("```\nlet x = 1;\nx += 1;\n```", exported, StringComparison.Ordinal);
        Assert.Contains("\n---", exported, StringComparison.Ordinal);

        // And the exported markdown re-imports cleanly: the full circle.
        WriteMd("roundtrip.md", exported);
        var envelope = Handler.CreateFrom(Ctx("roundtrip.docx"), "roundtrip.md");
        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatesClean(Path.Combine(Dir, "roundtrip.docx"));
    }

    [Fact]
    public void Markdown_view_is_listed_and_unknown_views_still_fail()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Read(Ctx(file, new JsonObject { ["view"] = "yaml" })));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("markdown", ex.Candidates!);
    }
}
