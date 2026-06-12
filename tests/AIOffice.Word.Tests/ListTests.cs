using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class ListTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    private string Html(string file) =>
        Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();

    // ------------------------------------------------------------ producing

    [Fact]
    public void Bullet_items_share_a_numbering_instance_and_get_reports_list_info()
    {
        var file = CreateDoc(title: "Groceries");

        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Apples","list":"bullet"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Pears","list":"bullet"}}
            ]
            """);

        var props = Get(file, "/body/p[3]")["properties"]!;
        Assert.Equal("bullet", props["list"]!.GetValue<string>());
        Assert.Equal(0, props["level"]!.GetValue<int>());
        Assert.Null(props["number"]); // bullets are unordered

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
            Assert.Single(numbering.Elements<AbstractNum>());
            Assert.Single(numbering.Elements<NumberingInstance>());

            var numIds = doc.MainDocumentPart.Document!.Body!.Elements<Paragraph>()
                .Select(p => p.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value)
                .OfType<int>()
                .ToList();
            Assert.Equal(2, numIds.Count);
            Assert.Equal(numIds[0], numIds[1]); // consecutive same-kind items share the instance
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Numbered_items_count_up_and_text_view_prefixes_markers()
    {
        var file = CreateDoc(title: "Steps");

        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Mix","list":"number"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Bake","list":"number"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Eat","list":"number"}}
            ]
            """);

        Assert.Equal(1, Get(file, "/body/p[3]")["properties"]!["number"]!.GetValue<int>());
        Assert.Equal(3, Get(file, "/body/p[5]")["properties"]!["number"]!.GetValue<int>());

        var texts = BodyTexts(file);
        Assert.Contains("1. Mix", texts);
        Assert.Contains("2. Bake", texts);
        Assert.Contains("3. Eat", texts);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Bullet_marker_prefixes_text_view()
    {
        var file = CreateDoc(title: "Dots");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Point","list":"bullet"}}]""");

        Assert.Contains("• Point", BodyTexts(file));
    }

    [Fact]
    public void Non_list_paragraph_breaks_the_sequence_and_numbering_restarts()
    {
        var file = CreateDoc(title: "Broken");

        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"One","list":"number"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Two","list":"number"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"An interlude"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"One again","list":"number"}}
            ]
            """);

        Assert.Equal(2, Get(file, "/body/p[4]")["properties"]!["number"]!.GetValue<int>());
        Assert.Equal(1, Get(file, "/body/p[6]")["properties"]!["number"]!.GetValue<int>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            // The fresh sequence is a separate instance with a startOverride, so Word restarts too.
            var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
            Assert.Equal(2, numbering.Elements<NumberingInstance>().Count());
            Assert.All(
                numbering.Elements<NumberingInstance>(),
                n => Assert.NotEmpty(n.Elements<LevelOverride>()));
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Different_list_kind_breaks_the_sequence()
    {
        var file = CreateDoc(title: "Mixed");

        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Bullet","list":"bullet"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Number","list":"number"}}
            ]
            """);

        Assert.Equal("bullet", Get(file, "/body/p[3]")["properties"]!["list"]!.GetValue<string>());
        Assert.Equal("number", Get(file, "/body/p[4]")["properties"]!["list"]!.GetValue<string>());
        Assert.Equal(1, Get(file, "/body/p[4]")["properties"]!["number"]!.GetValue<int>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
            Assert.Equal(2, numbering.Elements<AbstractNum>().Count()); // one per kind
            Assert.Equal(2, numbering.Elements<NumberingInstance>().Count());
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void ListRestart_forces_a_fresh_sequence()
    {
        var file = CreateDoc(title: "Restart");

        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"One","list":"number"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Two","list":"number"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"One again","list":"number","listRestart":true}}
            ]
            """);

        Assert.Equal(2, Get(file, "/body/p[4]")["properties"]!["number"]!.GetValue<int>());
        Assert.Equal(1, Get(file, "/body/p[5]")["properties"]!["number"]!.GetValue<int>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Nested_levels_reopen_after_a_shallower_item_resets_deeper_counters()
    {
        var file = CreateDoc(title: "Outline");

        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"First","list":"number"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"First sub","list":"number","level":1}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Second","list":"number"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Second sub","list":"number","level":1}}
            ]
            """);

        Assert.Equal(1, Get(file, "/body/p[4]")["properties"]!["number"]!.GetValue<int>());
        Assert.Equal(2, Get(file, "/body/p[5]")["properties"]!["number"]!.GetValue<int>());
        Assert.Equal(1, Get(file, "/body/p[6]")["properties"]!["number"]!.GetValue<int>()); // reset by the level-0 item
        AssertValidatesClean(file);
    }

    // -------------------------------------------------------------- render

    [Fact]
    public void Html_render_emits_ul_for_bullets_and_ol_for_numbers()
    {
        var file = CreateDoc(title: "Lists");

        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Dot","list":"bullet"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Plain"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Step","list":"number"}}
            ]
            """);

        var html = Html(file);
        Assert.Contains("<ul>\n<li data-aio-path=\"/body/p[3]\">Dot</li>\n</ul>", html, StringComparison.Ordinal);
        Assert.Contains("<ol>\n<li data-aio-path=\"/body/p[5]\">Step</li>\n</ol>", html, StringComparison.Ordinal);
        Assert.Contains("<p data-aio-path=\"/body/p[4]\">Plain</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_render_nests_sublists_inside_the_parent_li()
    {
        var file = CreateDoc(title: "Nested");

        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Parent","list":"bullet"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Child","list":"bullet","level":1}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Sibling","list":"bullet"}}
            ]
            """);

        var html = Html(file);
        Assert.Contains(
            "<li data-aio-path=\"/body/p[3]\">Parent<ul>\n<li data-aio-path=\"/body/p[4]\">Child</li>\n</ul>\n</li>",
            html,
            StringComparison.Ordinal);
        Assert.Contains("<li data-aio-path=\"/body/p[5]\">Sibling</li>", html, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------- honesty

    [Fact]
    public void Unknown_list_kind_is_invalid_args_with_candidates()
    {
        var file = CreateDoc(title: "Bad");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"x","list":"roman"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("bullet", ex.Candidates!);
        Assert.Contains("number", ex.Candidates!);
    }

    [Fact]
    public void Out_of_range_level_is_invalid_args()
    {
        var file = CreateDoc(title: "Deep");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"x","list":"bullet","level":9}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("0..8", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Level_without_list_is_invalid_args()
    {
        var file = CreateDoc(title: "Levels");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"x","level":1}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    // ------------------------------------------------------ reopen contract

    [Fact]
    public void List_numbering_xml_survives_reopen()
    {
        var file = CreateDoc(title: "Persist");
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Item","list":"number","level":1}}
            ]
            """);

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Last();
        var numPr = paragraph.ParagraphProperties!.NumberingProperties!;
        Assert.Equal(1, numPr.NumberingLevelReference!.Val!.Value);

        var numId = numPr.NumberingId!.Val!.Value;
        var numbering = doc.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
        var instance = numbering.Elements<NumberingInstance>().Single(n => n.NumberID!.Value == numId);
        var abstractId = instance.GetFirstChild<AbstractNumId>()!.Val!.Value;
        var abstractNum = numbering.Elements<AbstractNum>().Single(a => a.AbstractNumberId!.Value == abstractId);
        Assert.Equal(9, abstractNum.Elements<Level>().Count()); // levels 0..8 defined
    }
}
