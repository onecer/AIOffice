using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class LinkTests : WordTestBase
{
    private string Html(string file) =>
        Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();

    [Fact]
    public void External_link_lands_as_relationship_with_hyperlink_style()
    {
        var file = CreateDoc(title: "Docs");

        var envelope = Edit(
            file,
            """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"AIOffice","url":"https://example.com/aioffice"}}]""");

        Assert.Equal("/body/p[1]/link[1]", Data(envelope)["ops"]!.AsArray()[0]!["path"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var main = doc.MainDocumentPart!;
            var hyperlink = Assert.Single(main.Document!.Body!.Descendants<Hyperlink>());
            var relationship = Assert.Single(main.HyperlinkRelationships);
            Assert.Equal(hyperlink.Id!.Value, relationship.Id);
            Assert.True(relationship.IsExternal);
            Assert.Equal("https://example.com/aioffice", relationship.Uri.OriginalString);

            // The run inside is Hyperlink-styled and the style is defined.
            var run = hyperlink.ChildElements.OfType<Run>().Single();
            Assert.Equal("Hyperlink", run.RunProperties!.RunStyle!.Val!.Value);
            Assert.Contains(
                main.StyleDefinitionsPart!.Styles!.Elements<Style>(),
                s => s.StyleId?.Value == "Hyperlink");
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_reports_kind_url_and_text()
    {
        var file = CreateDoc(title: "Get");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"home","url":"https://example.org"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]/link[1]" })));

        Assert.Equal("link", got["type"]!.GetValue<string>());
        Assert.Equal("link", got["properties"]!["kind"]!.GetValue<string>());
        Assert.Equal("https://example.org", got["properties"]!["url"]!.GetValue<string>());
        Assert.Equal("home", got["properties"]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void Query_matches_links_by_url()
    {
        var file = CreateDoc(title: "Query");
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"link","props":{"text":"a","url":"https://example.com/a"}},
              {"op":"add","path":"/body/p[2]","type":"link","props":{"text":"b","url":"https://other.net/b"}}
            ]
            """);

        var data = Data(Handler.Query(Ctx(file, new JsonObject { ["selector"] = "link[url*=example.com]" })));

        Assert.Equal(1, data["count"]!.GetValue<int>());
        Assert.Equal("/body/p[1]/link[1]", data["matches"]!.AsArray()[0]!["path"]!.GetValue<string>());
    }

    [Fact]
    public void Html_render_emits_anchor_tags_with_href()
    {
        var file = CreateDoc(title: "Render");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"click","url":"https://example.com/?a=1&b=2"}}]""");

        var html = Html(file);

        Assert.Contains(
            """<a data-aio-path="/body/p[1]/link[1]" href="https://example.com/?a=1&amp;b=2">""",
            html,
            StringComparison.Ordinal);
        Assert.Contains(">click</a>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Mailto_links_are_allowed()
    {
        var file = CreateDoc(title: "Mail");

        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"write us","url":"mailto:hi@example.com"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]/link[1]" })));
        Assert.Equal("mailto:hi@example.com", got["properties"]!["url"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Anchor_link_targets_a_bookmark_and_renders_as_fragment_href()
    {
        var file = CreateDoc(title: "Internal");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"The results section"}}]""");

        Edit(file, """
            [
              {"op":"add","path":"/body/p[3]","type":"bookmark","props":{"name":"Results"}},
              {"op":"add","path":"/body/p[1]","type":"link","props":{"text":"see results","anchor":"Results"}}
            ]
            """);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var hyperlink = Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<Hyperlink>());
            Assert.Equal("Results", hyperlink.Anchor!.Value);
            Assert.Null(hyperlink.Id); // internal links carry no relationship
        }

        Assert.Contains("href=\"#Results\"", Html(file), StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Anchor_link_to_missing_bookmark_is_invalid_args()
    {
        var file = CreateDoc(title: "Dangling");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"x","anchor":"Nowhere"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("bookmark", ex.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Non_http_url_is_invalid_args()
    {
        var file = CreateDoc(title: "Schemes");

        foreach (var url in (string[])["file:///etc/passwd", "javascript:alert(1)", "not a url"])
        {
            var ops = """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"x","url":"URL"}}]"""
                .Replace("URL", url, StringComparison.Ordinal);
            var ex = Assert.Throws<AiofficeException>(() => Edit(file, ops));
            Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        }
    }

    [Fact]
    public void Url_and_anchor_together_or_neither_is_invalid_args()
    {
        var file = CreateDoc(title: "Exclusive");

        var both = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"x","url":"https://a.com","anchor":"B"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, both.Code);

        var neither = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"x"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, neither.Code);
    }

    [Fact]
    public void Remove_link_drops_element_and_relationship()
    {
        var file = CreateDoc(title: "Removable");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"gone","url":"https://example.com"}}]""");

        Edit(file, """[{"op":"remove","path":"/body/p[1]/link[1]"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<Hyperlink>());
            Assert.Empty(doc.MainDocumentPart.HyperlinkRelationships);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void External_link_tooltip_round_trips_as_screentip()
    {
        var file = CreateDoc(title: "Tip");

        var envelope = Edit(
            file,
            """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"Docs","url":"https://example.com/docs","tooltip":"Open the docs"}}]""");

        // add-response echoes the tooltip.
        Assert.Equal("Open the docs", Data(envelope)["ops"]!.AsArray()[0]!["tooltip"]!.GetValue<string>());

        // Written as w:hyperlink/@w:tooltip and survives save+reload.
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var hyperlink = Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<Hyperlink>());
            Assert.Equal("Open the docs", hyperlink.Tooltip!.Value);
        }

        // get surfaces it alongside url + text.
        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]/link[1]" })));
        Assert.Equal("Open the docs", got["properties"]!["tooltip"]!.GetValue<string>());
        Assert.Equal("https://example.com/docs", got["properties"]!["url"]!.GetValue<string>());
        Assert.Equal("Docs", got["properties"]!["text"]!.GetValue<string>());

        AssertValidatesClean(file);
    }

    [Fact]
    public void Internal_anchor_link_tooltip_round_trips()
    {
        var file = CreateDoc(title: "TipAnchor");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"The results section"}}]""");

        Edit(file, """
            [
              {"op":"add","path":"/body/p[3]","type":"bookmark","props":{"name":"Results"}},
              {"op":"add","path":"/body/p[1]","type":"link","props":{"text":"see results","anchor":"Results","tooltip":"Jump to results"}}
            ]
            """);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var hyperlink = Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<Hyperlink>());
            Assert.Equal("Results", hyperlink.Anchor!.Value);
            Assert.Null(hyperlink.Id);
            Assert.Equal("Jump to results", hyperlink.Tooltip!.Value);
        }

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]/link[1]" })));
        Assert.Equal("Jump to results", got["properties"]!["tooltip"]!.GetValue<string>());
        Assert.Equal("Results", got["properties"]!["anchor"]!.GetValue<string>());

        AssertValidatesClean(file);
    }

    [Fact]
    public void Link_without_tooltip_writes_no_tooltip_attribute()
    {
        var file = CreateDoc(title: "NoTip");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"plain","url":"https://example.com"}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var hyperlink = Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<Hyperlink>());
        Assert.Null(hyperlink.Tooltip); // no w:tooltip attribute at all
    }

    [Fact]
    public void Get_on_tooltipless_link_omits_tooltip_key()
    {
        // The read-side trap: a null Dictionary value would emit tooltip:null; the key must be absent.
        var file = CreateDoc(title: "TrapGuard");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"link","props":{"text":"home","url":"https://example.org"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]/link[1]" })));
        var properties = got["properties"]!.AsObject();

        Assert.False(properties.ContainsKey("tooltip"));
    }

    [Fact]
    public void Empty_or_non_string_tooltip_is_invalid_args()
    {
        var file = CreateDoc(title: "BadTip");

        foreach (var badProps in (string[])
                 [
                     """{"text":"x","url":"https://example.com","tooltip":""}""",
                     """{"text":"x","url":"https://example.com","tooltip":"   "}""",
                     """{"text":"x","url":"https://example.com","tooltip":42}""",
                     """{"text":"x","url":"https://example.com","tooltip":true}""",
                 ])
        {
            var ops = $$"""[{"op":"add","path":"/body/p[1]","type":"link","props":{{badProps}}}]""";
            var ex = Assert.Throws<AiofficeException>(() => Edit(file, ops));
            Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        }

        // And no partial hyperlink was written by any of the rejected attempts.
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<Hyperlink>());
    }

    [Fact]
    public void Link_text_is_required()
    {
        var file = CreateDoc(title: "Empty");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"link","props":{"url":"https://example.com"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("text", ex.Message, StringComparison.Ordinal);
    }
}
