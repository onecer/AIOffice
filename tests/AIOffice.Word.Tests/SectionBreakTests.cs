using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class SectionBreakTests : WordTestBase
{
    /// <summary>A doc whose body is exactly p[1]=A, p[2]=B, p[3]=C.</summary>
    private string CreateThreeParagraphDoc()
    {
        var file = CreateDoc();
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"A"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"B"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"C"}}
            ]
            """);
        return file;
    }

    private JsonNode GetSection(string file, int index) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = $"/section[{index}]" })))["properties"]!;

    [Fact]
    public void Section_breaks_create_independent_sections_with_mixed_page_setup()
    {
        var file = CreateThreeParagraphDoc();

        // p[1] ends section 1 (next page), p[2] ends section 2 (continuous), the body sectPr is section 3.
        var envelope = Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"sectionBreak","props":{"kind":"nextPage"}},
              {"op":"add","path":"/body/p[2]","type":"sectionBreak","props":{"kind":"continuous"}}
            ]
            """);
        var ops = JsonNode.Parse(envelope.ToJson())!["data"]!["ops"]!.AsArray();
        Assert.Equal("/section[1]", ops[0]!["path"]!.GetValue<string>());
        Assert.Equal("/section[2]", ops[1]!["path"]!.GetValue<string>());
        Assert.Equal(3, ops[1]!["sections"]!.GetValue<int>());

        Edit(file, """
            [
              {"op":"set","path":"/section[1]","props":{"pageSize":"A4","orientation":"landscape"}},
              {"op":"set","path":"/section[2]","props":{"pageSize":"A3"}},
              {"op":"set","path":"/section[3]","props":{"pageSize":"Letter","orientation":"landscape"}}
            ]
            """);

        // Reopen and verify each sectPr independently at the XML level.
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var paragraphs = body.Elements<Paragraph>().ToList();

            var first = paragraphs[0].ParagraphProperties!.SectionProperties!;
            Assert.Null(first.GetFirstChild<SectionType>()); // nextPage is the default, no w:type written
            var firstSize = first.GetFirstChild<PageSize>()!;
            Assert.Equal(16838u, firstSize.Width!.Value); // A4 landscape: swapped
            Assert.Equal(11906u, firstSize.Height!.Value);

            var second = paragraphs[1].ParagraphProperties!.SectionProperties!;
            Assert.Equal(SectionMarkValues.Continuous, second.GetFirstChild<SectionType>()!.Val!.Value);
            var secondSize = second.GetFirstChild<PageSize>()!;
            Assert.Equal(16838u, secondSize.Width!.Value); // A3 portrait
            Assert.Equal(23811u, secondSize.Height!.Value);

            var bodySectPr = body.Elements<SectionProperties>().Single();
            Assert.Same(bodySectPr, body.LastChild); // body-level sectPr stays last
            var thirdSize = bodySectPr.GetFirstChild<PageSize>()!;
            Assert.Equal(15840u, thirdSize.Width!.Value); // Letter landscape: swapped
            Assert.Equal(12240u, thirdSize.Height!.Value);
        }

        // And through the addressing surface.
        Assert.Equal("landscape", GetSection(file, 1)["orientation"]!.GetValue<string>());
        Assert.Equal("A4", GetSection(file, 1)["pageSize"]!.GetValue<string>());
        Assert.Equal("nextPage", GetSection(file, 1)["kind"]!.GetValue<string>());
        Assert.Equal("portrait", GetSection(file, 2)["orientation"]!.GetValue<string>());
        Assert.Equal("A3", GetSection(file, 2)["pageSize"]!.GetValue<string>());
        Assert.Equal("continuous", GetSection(file, 2)["kind"]!.GetValue<string>());
        Assert.Equal("landscape", GetSection(file, 3)["orientation"]!.GetValue<string>());
        Assert.Equal("Letter", GetSection(file, 3)["pageSize"]!.GetValue<string>());
        Assert.Equal(3, GetSection(file, 3)["sections"]!.GetValue<int>());

        AssertValidatesClean(file);
    }

    [Fact]
    public void New_break_clones_the_governing_sections_setup()
    {
        var file = CreateThreeParagraphDoc();
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"A3","orientation":"landscape"}}]""");

        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"sectionBreak","props":{"kind":"nextPage"}}]""");

        // Both halves start out identical, like Word's own section break insert.
        Assert.Equal("A3", GetSection(file, 1)["pageSize"]!.GetValue<string>());
        Assert.Equal("landscape", GetSection(file, 1)["orientation"]!.GetValue<string>());
        Assert.Equal("A3", GetSection(file, 2)["pageSize"]!.GetValue<string>());
        Assert.Equal("landscape", GetSection(file, 2)["orientation"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Structure_view_lists_sections_with_their_ranges()
    {
        var file = CreateThreeParagraphDoc();
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"sectionBreak","props":{"kind":"nextPage"}},
              {"op":"add","path":"/body/p[2]","type":"sectionBreak","props":{"kind":"continuous"}}
            ]
            """);

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));
        var sections = data["sections"]!.AsArray();

        Assert.Equal(3, sections.Count);
        Assert.Equal("/section[1]", sections[0]!["path"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", sections[0]!["start"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", sections[0]!["end"]!.GetValue<string>());
        Assert.Equal("nextPage", sections[0]!["kind"]!.GetValue<string>());
        Assert.Equal("/body/p[2]", sections[1]!["start"]!.GetValue<string>());
        Assert.Equal("continuous", sections[1]!["kind"]!.GetValue<string>());
        Assert.Equal("/body/p[3]", sections[2]!["start"]!.GetValue<string>());
        Assert.Equal("/body/p[3]", sections[2]!["end"]!.GetValue<string>());
        Assert.Equal(1, sections[2]!["blocks"]!.GetValue<int>());
    }

    [Fact]
    public void Single_section_documents_still_report_one_full_range()
    {
        var file = CreateThreeParagraphDoc();

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));
        var section = Assert.Single(data["sections"]!.AsArray())!;

        Assert.Equal("/section[1]", section["path"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", section["start"]!.GetValue<string>());
        Assert.Equal("/body/p[3]", section["end"]!.GetValue<string>());
        Assert.Equal(3, section["blocks"]!.GetValue<int>());
    }

    [Fact]
    public void Removing_a_break_merges_forward_into_the_following_section()
    {
        var file = CreateThreeParagraphDoc();
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"sectionBreak","props":{"kind":"nextPage"}}]""");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"orientation":"landscape"}}]""");

        var envelope = Edit(file, """[{"op":"remove","path":"/section[1]"}]""");

        var op = JsonNode.Parse(envelope.ToJson())!["data"]!["ops"]!.AsArray()[0]!;
        Assert.Equal("forward", op["merge"]!.GetValue<string>());
        Assert.Equal(1, op["sections"]!.GetValue<int>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Null(body.Elements<Paragraph>().First().ParagraphProperties?.SectionProperties);
        }

        // The merged content adopts the FOLLOWING section's setup (which was untouched portrait).
        Assert.Equal("portrait", GetSection(file, 1)["orientation"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Removing_the_final_section_merges_backward()
    {
        var file = CreateThreeParagraphDoc();
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"sectionBreak","props":{"kind":"nextPage"}}]""");
        Edit(file, """
            [
              {"op":"set","path":"/section[1]","props":{"orientation":"landscape"}},
              {"op":"set","path":"/section[2]","props":{"orientation":"portrait"}}
            ]
            """);

        var envelope = Edit(file, """[{"op":"remove","path":"/section[2]"}]""");

        var op = JsonNode.Parse(envelope.ToJson())!["data"]!["ops"]!.AsArray()[0]!;
        Assert.Equal("backward", op["merge"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Null(body.Elements<Paragraph>().First().ParagraphProperties?.SectionProperties);
            var bodySectPr = Assert.IsType<SectionProperties>(body.LastChild); // previous break's setup moved to the body
            Assert.Equal(PageOrientationValues.Landscape, bodySectPr.GetFirstChild<PageSize>()!.Orient!.Value);
        }

        Assert.Equal("landscape", GetSection(file, 1)["orientation"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Removing_the_only_section_is_invalid_args()
    {
        var file = CreateThreeParagraphDoc();
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"A4"}}]"""); // materialize it

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"remove","path":"/section[1]"}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Break_targets_must_be_top_level_body_paragraphs()
    {
        var file = CreateThreeParagraphDoc();
        Edit(file, """
            [
              {"op":"add","path":"/header[1]","type":"header","props":{"text":"H"}},
              {"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":1}}
            ]
            """);

        var header = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/header[1]/p[1]","type":"sectionBreak"}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, header.Code);

        var cell = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/table[1]/tr[1]/tc[1]/p[1]","type":"sectionBreak"}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, cell.Code);
    }

    [Fact]
    public void Duplicate_break_on_the_same_paragraph_is_invalid_args()
    {
        var file = CreateThreeParagraphDoc();
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"sectionBreak"}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"sectionBreak"}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("already", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_kind_is_invalid_args_with_candidates()
    {
        var file = CreateThreeParagraphDoc();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"sectionBreak","props":{"kind":"evenPage"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("continuous", ex.Candidates!);
    }

    [Fact]
    public void Tracked_section_break_is_unsupported()
    {
        var file = CreateThreeParagraphDoc();

        var ex = Assert.Throws<AiofficeException>(() => Edit(
            file,
            """[{"op":"add","path":"/body/p[1]","type":"sectionBreak"}]""",
            new JsonObject { ["track"] = true }));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }
}
