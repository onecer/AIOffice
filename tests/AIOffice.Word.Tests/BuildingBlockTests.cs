using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>v1.5.0 building blocks / Quick Parts: glossary-part storage, insert-by-name, get/remove, structure.</summary>
public sealed class BuildingBlockTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    private JsonNode Structure(string file) =>
        Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));

    private string StoreDisclaimer(string file, string gallery = "quickParts")
    {
        var ops = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "add",
                ["path"] = "/buildingBlocks",
                ["type"] = "buildingBlock",
                ["props"] = new JsonObject
                {
                    ["name"] = "Disclaimer",
                    ["gallery"] = gallery,
                    ["content"] = "Confidential — do not distribute.",
                },
            },
        };
        return Edit(file, ops.ToJsonString()).ToJson();
    }

    // ----------------------------------------------------------------- store

    [Fact]
    public void Add_stores_a_building_block_in_the_glossary_part()
    {
        var file = CreateDoc(title: "Blocks");

        StoreDisclaimer(file);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var glossary = doc.MainDocumentPart!.GlossaryDocumentPart;
            Assert.NotNull(glossary);
            var docPart = glossary!.GlossaryDocument!.DocParts!.Elements<DocPart>().Single();
            Assert.Equal("Disclaimer", docPart.DocPartProperties!.DocPartName!.Val!.Value);
            Assert.Contains("Confidential", docPart.DocPartBody!.InnerText, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_reports_the_stored_block()
    {
        var file = CreateDoc(title: "Blocks");
        StoreDisclaimer(file, gallery: "autoText");

        var props = Get(file, "/buildingBlock[@name=Disclaimer]")["properties"]!;
        Assert.Equal("Disclaimer", props["name"]!.GetValue<string>());
        Assert.Equal("autoText", props["gallery"]!.GetValue<string>());
        Assert.Equal(1, props["paragraphs"]!.GetValue<int>());
        Assert.Contains("Confidential", props["content"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void Multiline_content_stores_one_paragraph_per_line()
    {
        var file = CreateDoc(title: "Blocks");
        Edit(file, """
        [{"op":"add","path":"/buildingBlocks","type":"buildingBlock",
          "props":{"name":"Address","gallery":"quickParts","content":"Acme Inc.\n123 Main St\nMetropolis"}}]
        """);

        var props = Get(file, "/buildingBlock[@name=Address]")["properties"]!;
        Assert.Equal(3, props["paragraphs"]!.GetValue<int>());
        Assert.Equal(3, props["contentLines"]!.AsArray().Count);
        Assert.Equal("123 Main St", props["contentLines"]!.AsArray()[1]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Array_content_stores_one_paragraph_per_entry()
    {
        var file = CreateDoc(title: "Blocks");
        Edit(file, """
        [{"op":"add","path":"/buildingBlocks","type":"buildingBlock",
          "props":{"name":"Bullets","gallery":"quickParts","content":["First","Second"]}}]
        """);

        var props = Get(file, "/buildingBlock[@name=Bullets]")["properties"]!;
        Assert.Equal(2, props["paragraphs"]!.GetValue<int>());
        AssertValidatesClean(file);
    }

    // ----------------------------------------------------------- insert by name

    [Fact]
    public void Insert_by_name_reproduces_the_stored_content_in_the_body()
    {
        var file = CreateDoc(title: "Blocks");
        StoreDisclaimer(file);

        var before = BodyTexts(file).Count;
        Edit(file, """[{"op":"add","path":"/body","type":"buildingBlockRef","props":{"name":"Disclaimer"}}]""");

        var texts = BodyTexts(file);
        Assert.Equal(before + 1, texts.Count);
        Assert.Contains(texts, t => t.Contains("Confidential", StringComparison.Ordinal));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Insert_before_a_paragraph_places_content_above_it()
    {
        var file = CreateDoc(title: "Top");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Anchor"}}]""");
        Edit(file, """
        [{"op":"add","path":"/buildingBlocks","type":"buildingBlock",
          "props":{"name":"Intro","gallery":"quickParts","content":"INSERTED"}}]
        """);

        // Anchor is the last body paragraph (after the title); insert before it.
        var anchorPath = "/body/p[" + BodyTexts(file).Count + "]";
        var refOps = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "add",
                ["path"] = anchorPath,
                ["type"] = "buildingBlockRef",
                ["position"] = "before",
                ["props"] = new JsonObject { ["name"] = "Intro" },
            },
        };
        Edit(file, refOps.ToJsonString());

        var texts = BodyTexts(file);
        var insertedIndex = texts.ToList().FindIndex(t => t.Contains("INSERTED", StringComparison.Ordinal));
        var anchorIndex = texts.ToList().FindIndex(t => t.Contains("Anchor", StringComparison.Ordinal));
        Assert.True(insertedIndex >= 0 && insertedIndex < anchorIndex);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Insert_of_an_unknown_block_is_invalid_args()
    {
        var file = CreateDoc(title: "Blocks");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"buildingBlockRef","props":{"name":"Missing"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    // -------------------------------------------------------------- replace

    [Fact]
    public void Re_adding_the_same_name_replaces_in_place()
    {
        var file = CreateDoc(title: "Blocks");
        StoreDisclaimer(file);
        Edit(file, """
        [{"op":"add","path":"/buildingBlocks","type":"buildingBlock",
          "props":{"name":"Disclaimer","gallery":"quickParts","content":"New text."}}]
        """);

        // Still exactly one block, now with the new content.
        var blocks = Structure(file)["buildingBlocks"]!.AsArray();
        Assert.Single(blocks);
        Assert.Contains("New text.", Get(file, "/buildingBlock[@name=Disclaimer]")["properties"]!["content"]!.GetValue<string>(), StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    // ---------------------------------------------------------------- remove

    [Fact]
    public void Remove_drops_the_block()
    {
        var file = CreateDoc(title: "Blocks");
        StoreDisclaimer(file);

        Edit(file, """[{"op":"remove","path":"/buildingBlock[@name=Disclaimer]"}]""");

        Assert.Null(Structure(file)["buildingBlocks"]);
        var ex = Assert.Throws<AiofficeException>(() => Get(file, "/buildingBlock[@name=Disclaimer]"));
        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        AssertValidatesClean(file);
    }

    // -------------------------------------------------------------- structure

    [Fact]
    public void Structure_lists_building_blocks()
    {
        var file = CreateDoc(title: "Blocks");
        StoreDisclaimer(file);
        Edit(file, """
        [{"op":"add","path":"/buildingBlocks","type":"buildingBlock",
          "props":{"name":"Sig","gallery":"autoText","content":"Best regards"}}]
        """);

        var blocks = Structure(file)["buildingBlocks"]!.AsArray();
        Assert.Equal(2, blocks.Count);
        Assert.Contains(blocks, b => b!["name"]!.GetValue<string>() == "Disclaimer");
        Assert.Contains(blocks, b => b!["name"]!.GetValue<string>() == "Sig" && b["gallery"]!.GetValue<string>() == "autoText");
    }

    [Fact]
    public void Unknown_gallery_is_invalid_args()
    {
        var file = CreateDoc(title: "Blocks");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """
            [{"op":"add","path":"/buildingBlocks","type":"buildingBlock",
              "props":{"name":"X","gallery":"cover","content":"y"}}]
            """));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }
}
