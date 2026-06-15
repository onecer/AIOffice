using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using B = DocumentFormat.OpenXml.Bibliography;

namespace AIOffice.Word.Tests;

/// <summary>v1.1.0 citations and bibliography: sources store, CITATION fields, BIBLIOGRAPHY block.</summary>
public sealed class CitationTests : WordTestBase
{
    private static List<string> WarningCodes(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"] is JsonArray warnings
            ? [.. warnings.Select(w => w!["code"]!.GetValue<string>())]
            : [];

    /// <summary>Adds one book source tagged Smith2020.</summary>
    private string CreateDocWithSource(string tag = "Smith2020")
    {
        var file = CreateDoc();
        Edit(file,
            "[{\"op\":\"add\",\"path\":\"/sources\",\"type\":\"source\",\"props\":{" +
            "\"tag\":\"" + tag + "\",\"kind\":\"book\",\"author\":\"Smith, John\"," +
            "\"title\":\"A Great Book\",\"year\":2020,\"publisher\":\"Acme Press\"}}]");
        return file;
    }

    // ------------------------------------------------------------------ sources

    [Fact]
    public void Add_source_writes_a_bibliography_sources_part_and_validates()
    {
        var file = CreateDocWithSource();

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "sources" })));
        Assert.Equal(1, data["count"]!.GetValue<int>());
        var src = data["sources"]!.AsArray()[0]!;
        Assert.Equal("/source[@tag=Smith2020]", src["path"]!.GetValue<string>());
        Assert.Equal("Smith2020", src["properties"]!["tag"]!.GetValue<string>());
        Assert.Equal("book", src["properties"]!["kind"]!.GetValue<string>());
        Assert.Equal("Smith, John", src["properties"]!["author"]!.GetValue<string>());

        // The store lives in a customXml part rooted at b:Sources.
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var part = doc.MainDocumentPart!.CustomXmlParts.Single();
        using var reader = new StreamReader(part.GetStream());
        var xml = reader.ReadToEnd();
        Assert.Contains("Sources", xml, StringComparison.Ordinal);
        Assert.Contains("Smith2020", xml, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_source_returns_its_props()
    {
        var file = CreateDocWithSource();

        var data = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/source[@tag=Smith2020]" })));
        Assert.Equal("source", data["type"]!.GetValue<string>());
        Assert.Equal("A Great Book", data["properties"]!["title"]!.GetValue<string>());
        Assert.Equal("2020", data["properties"]!["year"]!.GetValue<string>());
        Assert.Equal("Acme Press", data["properties"]!["publisher"]!.GetValue<string>());
    }

    [Fact]
    public void Re_adding_a_tag_replaces_in_place()
    {
        var file = CreateDocWithSource();
        Edit(file, """
            [{"op":"add","path":"/sources","type":"source","props":{
              "tag":"Smith2020","kind":"book","title":"Revised Title","year":2021}}]
            """);

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "sources" })));
        Assert.Equal(1, data["count"]!.GetValue<int>());
        Assert.Equal("Revised Title", data["sources"]!.AsArray()[0]!["properties"]!["title"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_source_drops_it()
    {
        var file = CreateDocWithSource();
        Edit(file, """[{"op":"remove","path":"/source[@tag=Smith2020]"}]""");

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "sources" })));
        Assert.Equal(0, data["count"]!.GetValue<int>());
    }

    [Fact]
    public void Unknown_source_kind_is_invalid_args()
    {
        var file = CreateDoc();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/sources","type":"source","props":{"tag":"X","kind":"podcast"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("book", ex.Candidates!);
    }

    // ---------------------------------------------------------------- citations

    [Fact]
    public void Add_citation_emits_a_field_with_a_cached_label_and_validates()
    {
        var file = CreateDocWithSource();
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"As shown previously."}}]""");

        var envelope = Edit(file, """[{"op":"add","path":"/body/p[1]","type":"citation","props":{"source":"Smith2020"}}]""");
        var op = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("(Smith, 2020)", op["cached"]!.GetValue<string>());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().First();
        Assert.Contains("CITATION Smith2020", field.Instruction!.Value, StringComparison.Ordinal);
        Assert.Equal("(Smith, 2020)", field.InnerText);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Citation_with_pages_and_suppressAuthor_reflects_in_the_label_and_instruction()
    {
        var file = CreateDocWithSource();

        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"citation","props":{"source":"Smith2020","pages":"42","suppressAuthor":true}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().First();
        Assert.Contains("\\p \"42\"", field.Instruction!.Value, StringComparison.Ordinal);
        Assert.Contains("\\n", field.Instruction!.Value, StringComparison.Ordinal);
        // suppressAuthor drops the author; pages append.
        Assert.Equal("(2020, p. 42)", field.InnerText);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Citation_to_unknown_tag_is_invalid_args_with_candidate_tags()
    {
        var file = CreateDocWithSource();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"citation","props":{"source":"Nope2099"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("Smith2020", ex.Candidates!);
    }

    // ------------------------------------------------------------- bibliography

    [Fact]
    public void Add_bibliography_emits_field_plus_static_entries_and_a_cached_warning()
    {
        var file = CreateDocWithSource();
        // A cited source so the bibliography includes it.
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"citation","props":{"source":"Smith2020"}}]""");

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"bibliography","props":{"style":"APA"}}]""");
        var op = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("/bibliography[1]", op["path"]!.GetValue<string>());
        Assert.Equal(1, op["entries"]!.GetValue<int>());
        Assert.Contains("bibliography_cached", WarningCodes(envelope));

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var sdt = body.Elements<SdtBlock>().Single(s =>
            s.SdtProperties!.GetFirstChild<SdtContentDocPartObject>()!.GetFirstChild<DocPartGallery>()!.Val!.Value == "Bibliographies");

        // Complex BIBLIOGRAPHY field.
        var code = sdt.Descendants<FieldCode>().Single();
        Assert.Contains("BIBLIOGRAPHY", code.Text, StringComparison.Ordinal);
        var fieldTypes = sdt.Descendants<FieldChar>().Select(f => f.FieldCharType!.Value).ToList();
        Assert.Equal(new[] { FieldCharValues.Begin, FieldCharValues.Separate, FieldCharValues.End }, fieldTypes);

        // The pre-rendered entry mentions the source so it reads without recompute.
        Assert.Contains("A Great Book", sdt.InnerText, StringComparison.Ordinal);
        Assert.Contains("Smith", sdt.InnerText, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Bibliography_entries_match_the_sources()
    {
        var file = CreateDoc();
        Edit(file, """
            [
              {"op":"add","path":"/sources","type":"source","props":{"tag":"A","kind":"book","author":"Adams, Ann","title":"Alpha","year":2001}},
              {"op":"add","path":"/sources","type":"source","props":{"tag":"B","kind":"journalArticle","author":"Brown, Bob","title":"Beta","year":2002,"journal":"J. Things"}}
            ]
            """);

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"bibliography"}]""");
        // No citations yet -> every stored source is listed.
        Assert.Equal(2, Data(envelope)["ops"]!.AsArray()[0]!["entries"]!.GetValue<int>());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var text = doc.MainDocumentPart!.Document!.Body!.Elements<SdtBlock>()
            .Single(s => s.SdtProperties!.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value == "Bibliographies")
            .InnerText;
        Assert.Contains("Alpha", text, StringComparison.Ordinal);
        Assert.Contains("Beta", text, StringComparison.Ordinal);
        Assert.Contains("J. Things", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Re_adding_bibliography_replaces_it()
    {
        var file = CreateDocWithSource();
        Edit(file, """[{"op":"add","path":"/body","type":"bibliography","props":{"style":"APA"}}]""");
        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"bibliography","props":{"style":"MLA"}}]""");

        Assert.True(Data(envelope)["ops"]!.AsArray()[0]!["replaced"]!.GetValue<bool>());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var blocks = doc.MainDocumentPart!.Document!.Body!.Elements<SdtBlock>()
            .Count(s => s.SdtProperties!.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value == "Bibliographies");
        Assert.Equal(1, blocks);
    }

    [Fact]
    public void Remove_bibliography_drops_the_block_but_keeps_sources()
    {
        var file = CreateDocWithSource();
        Edit(file, """[{"op":"add","path":"/body","type":"bibliography"}]""");
        Edit(file, """[{"op":"remove","path":"/bibliography[1]"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.DoesNotContain(doc.MainDocumentPart!.Document!.Body!.Elements<SdtBlock>(),
                s => s.SdtProperties!.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value == "Bibliographies");
        }

        // The sources store survives a bibliography removal.
        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "sources" })));
        Assert.Equal(1, data["count"]!.GetValue<int>());
    }

    [Fact]
    public void Unknown_bibliography_style_is_invalid_args()
    {
        var file = CreateDocWithSource();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"bibliography","props":{"style":"Vancouver"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    // -------------------------------------------------------------- round-trip

    [Fact]
    public void Sources_citation_and_bibliography_survive_a_no_op_reopen()
    {
        var file = CreateDocWithSource();
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"citation","props":{"source":"Smith2020"}}]""");
        Edit(file, """[{"op":"add","path":"/body","type":"bibliography","props":{"style":"APA"}}]""");

        var before = File.ReadAllBytes(file);
        using (var doc = WordprocessingDocument.Open(file, isEditable: true))
        {
            doc.Save();
        }

        var after = File.ReadAllBytes(file);
        Assert.Equal(before.Length, after.Length);

        // Everything still reads back.
        Assert.Equal(1, Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "sources" })))["count"]!.GetValue<int>());
        AssertValidatesClean(file);
    }
}
