using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>v1.21 section verticalAlign: w:sectPr/w:vAlign — surface 'justify' is OOXML 'both'.</summary>
public sealed class SectionVerticalAlignTests : WordTestBase
{
    private JsonNode GetSection(string file, string path = "/section[1]") =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    [Fact]
    public void Set_center_writes_vAlign_at_its_schema_rank()
    {
        var file = CreateDoc(title: "Cover");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"H"}}]""");

        Edit(file, """
        [{"op":"set","path":"/section[1]","props":{"pageSize":"A4","columns":2,"verticalAlign":"center"}}]
        """);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
            var vAlign = sectPr.GetFirstChild<VerticalTextAlignmentOnPage>()!;
            Assert.Equal(VerticalJustificationValues.Center, vAlign.Val!.Value);
            Assert.Equal("center", vAlign.Val.InnerText); // raw w:val

            // SectPrOrder rank: after headerRef/pgSz/cols, before nothing that outranks it here.
            var order = sectPr.ChildElements.Select(c => c.GetType().Name).ToList();
            Assert.True(
                order.IndexOf("HeaderReference") < order.IndexOf("VerticalTextAlignmentOnPage"),
                string.Join(",", order));
            Assert.True(
                order.IndexOf("PageSize") < order.IndexOf("VerticalTextAlignmentOnPage"),
                string.Join(",", order));
            Assert.True(
                order.IndexOf("Columns") < order.IndexOf("VerticalTextAlignmentOnPage"),
                string.Join(",", order));
        }

        Assert.Equal("center", GetSection(file)["properties"]!["verticalAlign"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Theory]
    [InlineData("top", "top")]
    [InlineData("center", "center")]
    [InlineData("justify", "both")] // the trap: surface 'justify' <-> OOXML 'both'
    [InlineData("bottom", "bottom")]
    public void All_tokens_round_trip_and_both_never_surfaces(string token, string xmlVal)
    {
        var file = CreateDoc(title: "Tokens");

        Edit(file, """[{"op":"set","path":"/section[1]","props":{"verticalAlign":"TOKEN"}}]""".Replace("TOKEN", token));

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var vAlign = doc.MainDocumentPart!.Document!.Body!
                .Elements<SectionProperties>().Single()
                .GetFirstChild<VerticalTextAlignmentOnPage>()!;
            Assert.Equal(xmlVal, vAlign.Val!.InnerText); // 'justify' writes w:val="both", never "justify"
        }

        var readBack = GetSection(file)["properties"]!["verticalAlign"]!.GetValue<string>();
        Assert.Equal(token, readBack);
        Assert.NotEqual("both", readBack); // the XML name never leaks onto the surface
        AssertValidatesClean(file);
    }

    [Fact]
    public void Legacy_section_gains_no_key_and_omitting_the_prop_writes_nothing()
    {
        var file = CreateDoc(title: "Legacy");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"A4"}}]""");

        // No w:vAlign was written by an unrelated set…
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Null(doc.MainDocumentPart!.Document!.Body!
                .Elements<SectionProperties>().Single()
                .GetFirstChild<VerticalTextAlignmentOnPage>());
        }

        // …and the wire shape carries no verticalAlign key at all (byte-stable projection).
        var properties = GetSection(file)["properties"]!.AsObject();
        Assert.False(properties.ContainsKey("verticalAlign"));

        // Same for the implicit (never materialized) section.
        var implicitFile = CreateDoc(name: "implicit.docx", title: "Implicit");
        Assert.False(GetSection(implicitFile)["properties"]!.AsObject().ContainsKey("verticalAlign"));
    }

    [Fact]
    public void Multi_prop_set_stays_ordered_and_sections_are_independent()
    {
        var file = CreateDoc(title: "Multi");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"A4"}}]""");

        // Simulate a Word-authored section break: a paragraph-level sectPr ahead of the body one.
        using (var doc = WordprocessingDocument.Open(file, isEditable: true))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
            var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();
            pPr.SectionProperties = new SectionProperties();
        }

        Edit(file, """[{"op":"set","path":"/section[2]","props":{"pageSize":"A4","verticalAlign":"bottom"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sections = doc.MainDocumentPart!.Document!.Body!.Descendants<SectionProperties>().ToList();
            Assert.Null(sections[0].GetFirstChild<VerticalTextAlignmentOnPage>()); // /section[1] untouched

            var sectPr = sections[1];
            Assert.Equal(
                VerticalJustificationValues.Bottom,
                sectPr.GetFirstChild<VerticalTextAlignmentOnPage>()!.Val!.Value);
            var order = sectPr.ChildElements.Select(c => c.GetType().Name).ToList();
            Assert.True(
                order.IndexOf("PageSize") < order.IndexOf("VerticalTextAlignmentOnPage"),
                string.Join(",", order));
        }

        Assert.False(GetSection(file, "/section[1]")["properties"]!.AsObject().ContainsKey("verticalAlign"));
        Assert.Equal("bottom", GetSection(file, "/section[2]")["properties"]!["verticalAlign"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Setting_twice_replaces_the_single_vAlign()
    {
        var file = CreateDoc(title: "Twice");

        Edit(file, """[{"op":"set","path":"/section[1]","props":{"verticalAlign":"top"}}]""");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"verticalAlign":"justify"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var vAligns = doc.MainDocumentPart!.Document!.Body!
                .Elements<SectionProperties>().Single()
                .Elements<VerticalTextAlignmentOnPage>().ToList();
            Assert.Single(vAligns); // replace, not stack
            Assert.Equal("both", vAligns[0].Val!.InnerText);
        }

        Assert.Equal("justify", GetSection(file)["properties"]!["verticalAlign"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Unknown_token_is_invalid_args_with_candidates()
    {
        var file = CreateDoc(title: "Bad");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/section[1]","props":{"verticalAlign":"middle"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Equal(["top", "center", "justify", "bottom"], ex.Candidates!);
    }
}
