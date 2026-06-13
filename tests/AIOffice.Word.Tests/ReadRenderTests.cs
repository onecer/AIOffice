using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class ReadRenderTests : WordTestBase
{
    private string CreateSample()
    {
        var file = CreateDoc(title: "Doc Title");
        Edit(file, """
            [
              {"op":"set","path":"/body/p[2]","props":{"text":"Intro words here."}},
              {"op":"add","path":"/body","props":{"text":"Chapter One","style":"Heading2"}},
              {"op":"add","path":"/body","props":{"text":"Strong claim","bold":true}},
              {"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":2}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"K"}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"V"}}
            ]
            """);
        return file;
    }

    private JsonNode Read(string file, JsonObject args) => Data(Handler.Read(Ctx(file, args)));

    [Fact]
    public void Text_view_prefixes_every_paragraph_with_its_path()
    {
        var file = CreateSample();

        var lines = Read(file, new JsonObject { ["view"] = "text" })["lines"]!.AsArray();

        Assert.Equal("/body/p[1]", lines[0]!["path"]!.GetValue<string>());
        Assert.Equal("Doc Title", lines[0]!["text"]!.GetValue<string>());
        // Table-cell paragraphs are included with their full canonical paths.
        Assert.Contains(lines, l => l!["path"]!.GetValue<string>() == "/body/table[1]/tr[1]/tc[1]/p[1]");
    }

    [Fact]
    public void Text_view_range_pages_through_paragraphs()
    {
        var file = CreateSample();

        var data = Read(file, new JsonObject { ["view"] = "text", ["range"] = "3..4" });

        var lines = data["lines"]!.AsArray();
        Assert.Equal(2, lines.Count);
        Assert.Equal("/body/p[3]", lines[0]!["path"]!.GetValue<string>());
    }

    [Fact]
    public void Text_view_max_bytes_truncates_with_a_warning()
    {
        var file = CreateSample();

        var envelope = Handler.Read(Ctx(file, new JsonObject { ["view"] = "text", ["maxBytes"] = 30 }));

        var json = JsonNode.Parse(envelope.ToJson())!;
        Assert.True(json["data"]!["lines"]!.AsArray().Count < json["data"]!["totalParagraphs"]!.GetValue<int>());
        Assert.Contains(
            json["meta"]!["warnings"]!.AsArray(),
            w => w!["code"]!.GetValue<string>() == "result_truncated");
    }

    [Fact]
    public void Outline_nests_heading2_under_heading1()
    {
        var file = CreateSample();

        var headings = Read(file, new JsonObject { ["view"] = "outline" })["headings"]!.AsArray();

        var h1 = Assert.Single(headings)!;
        Assert.Equal("Doc Title", h1["text"]!.GetValue<string>());
        Assert.Equal(1, h1["level"]!.GetValue<int>());
        var h2 = Assert.Single(h1["children"]!.AsArray())!;
        Assert.Equal("Chapter One", h2["text"]!.GetValue<string>());
    }

    [Fact]
    public void Stats_counts_paragraphs_words_tables_headings()
    {
        var file = CreateSample();

        var stats = Read(file, new JsonObject { ["view"] = "stats" });

        Assert.Equal(6, stats["paragraphs"]!.GetValue<int>()); // 4 body + 2 cell paragraphs
        Assert.Equal(1, stats["tables"]!.GetValue<int>());
        Assert.Equal(2, stats["headings"]!.GetValue<int>());
        Assert.True(stats["words"]!.GetValue<int>() >= 9);
    }

    [Fact]
    public void Structure_view_is_depth_limited()
    {
        var file = CreateSample();

        var root = Read(file, new JsonObject { ["view"] = "structure", ["depth"] = 1 })["root"]!;

        Assert.Equal("body", root["type"]!.GetValue<string>());
        var table = root["children"]!.AsArray().First(c => c!["type"]!.GetValue<string>() == "table")!;
        Assert.Equal(1, table["childCount"]!.GetValue<int>());
        Assert.Null(table["children"]); // depth 1 stops below body children
    }

    [Fact]
    public void Unknown_view_is_invalid_args_with_candidates()
    {
        var file = CreateSample();

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Read(Ctx(file, new JsonObject { ["view"] = "summary" })));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("outline", ex.Candidates!);
    }

    [Fact]
    public void Render_html_emits_semantic_markup()
    {
        var file = CreateSample();

        var html = Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();

        Assert.Contains("""<h1 data-aio-path="/body/p[1]">Doc Title</h1>""", html, StringComparison.Ordinal);
        Assert.Contains("""<h2 data-aio-path="/body/p[3]">Chapter One</h2>""", html, StringComparison.Ordinal);
        Assert.Contains("<strong>Strong claim</strong>", html, StringComparison.Ordinal);
        Assert.Contains(
            """<td data-aio-path="/body/table[1]/tr[1]/tc[1]">K</td><td data-aio-path="/body/table[1]/tr[1]/tc[2]">V</td>""",
            html,
            StringComparison.Ordinal);
    }

    /// <summary>The data-aio-path render contract: every addressable block maps back to a path.</summary>
    [Fact]
    public void Render_html_tags_every_block_with_data_aio_path()
    {
        var file = CreateSample();

        var html = Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();

        Assert.Contains("""<p data-aio-path="/body/p[2]">Intro words here.</p>""", html, StringComparison.Ordinal);
        Assert.Contains("""<table data-aio-path="/body/table[1]">""", html, StringComparison.Ordinal);
        Assert.Contains("""<tr data-aio-path="/body/table[1]/tr[1]">""", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Render_scope_limits_output_to_one_node()
    {
        var file = CreateSample();

        var html = Data(Handler.Render(Ctx(file, new JsonObject
        {
            ["to"] = "html",
            ["scope"] = "/body/p[3]",
        })))["content"]!.GetValue<string>();

        Assert.Equal("""<h2 data-aio-path="/body/p[3]">Chapter One</h2>""", html);
    }

    [Fact]
    public void Render_text_returns_plain_paragraphs()
    {
        var file = CreateSample();

        var text = Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "text" })))["content"]!.GetValue<string>();

        Assert.Contains("Doc Title", text, StringComparison.Ordinal);
        Assert.Contains("K\tV", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Render_html_escapes_reserved_characters()
    {
        var file = CreateDoc();
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"a < b & c"}}]""");

        var html = Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();

        Assert.Contains("a &lt; b &amp; c", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Render_svg_is_unsupported_with_html_workaround()
    {
        var file = CreateSample();

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Render(Ctx(file, new JsonObject { ["to"] = "svg" })));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("html", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Render_out_writes_the_file_inside_the_workspace()
    {
        var file = CreateSample();

        Handler.Render(Ctx(file, new JsonObject { ["to"] = "html", ["output"] = "render.html" }));

        Assert.True(File.Exists(Path.Combine(Dir, "render.html")));
    }
}
