using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class SectionTests : WordTestBase
{
    private JsonNode GetSection(string file, string path = "/section[1]") =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    [Fact]
    public void Set_page_size_and_orientation_writes_swapped_pgSz()
    {
        var file = CreateDoc(title: "Wide");

        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"A4","orientation":"landscape"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
            var pgSz = sectPr.GetFirstChild<PageSize>()!;
            Assert.Equal(16838u, pgSz.Width!.Value); // A4 landscape: dimensions swapped
            Assert.Equal(11906u, pgSz.Height!.Value);
            Assert.Equal(PageOrientationValues.Landscape, pgSz.Orient!.Value);
        }

        var properties = GetSection(file)["properties"]!;
        Assert.Equal("landscape", properties["orientation"]!.GetValue<string>());
        Assert.Equal("A4", properties["pageSize"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_margins_in_cm_writes_complete_pgMar()
    {
        var file = CreateDoc(title: "Margins");

        Edit(file, """[{"op":"set","path":"/section[1]","props":{"marginTop":"2cm","marginLeft":"3cm"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var pgMar = doc.MainDocumentPart!.Document!.Body!
                .Elements<SectionProperties>().Single()
                .GetFirstChild<PageMargin>()!;
            Assert.Equal(1134, pgMar.Top!.Value); // 2cm = 1134 twips
            Assert.Equal(1701u, pgMar.Left!.Value); // 3cm
            Assert.Equal(1440, pgMar.Bottom!.Value); // untouched margins keep Word defaults
            Assert.Equal(1440u, pgMar.Right!.Value);
            Assert.NotNull(pgMar.Header); // pgMar stays schema-complete
            Assert.NotNull(pgMar.Footer);
            Assert.NotNull(pgMar.Gutter);
        }

        var properties = GetSection(file)["properties"]!;
        Assert.Equal(2.0, properties["marginTopCm"]!.GetValue<double>());
        Assert.Equal(3.0, properties["marginLeftCm"]!.GetValue<double>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Orientation_only_swaps_existing_dimensions_back_and_forth()
    {
        var file = CreateDoc(title: "Flip");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"Letter"}}]""");

        Edit(file, """[{"op":"set","path":"/section[1]","props":{"orientation":"landscape"}}]""");
        var landscape = GetSection(file)["properties"]!;
        Assert.Equal("landscape", landscape["orientation"]!.GetValue<string>());
        Assert.Equal("Letter", landscape["pageSize"]!.GetValue<string>());
        Assert.True(landscape["widthCm"]!.GetValue<double>() > landscape["heightCm"]!.GetValue<double>());

        Edit(file, """[{"op":"set","path":"/section[1]","props":{"orientation":"portrait"}}]""");
        var portrait = GetSection(file)["properties"]!;
        Assert.Equal("portrait", portrait["orientation"]!.GetValue<string>());
        Assert.True(portrait["widthCm"]!.GetValue<double>() < portrait["heightCm"]!.GetValue<double>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_on_document_without_sectPr_reports_portrait_defaults()
    {
        var file = CreateDoc(title: "Implicit");

        var got = GetSection(file);

        Assert.Equal("section", got["type"]!.GetValue<string>());
        Assert.Equal("portrait", got["properties"]!["orientation"]!.GetValue<string>());
        Assert.Null(got["properties"]!["pageSize"]);

        // Reading must not have materialized anything.
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>());
    }

    [Fact]
    public void Set_materializes_the_implicit_section_and_keeps_sectPr_last()
    {
        var file = CreateDoc(title: "Materialize");

        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"A3"}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.IsType<SectionProperties>(body.LastChild);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Header_reference_stays_ahead_of_page_setup_in_sectPr()
    {
        var file = CreateDoc(title: "Ordered");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"H"}}]""");

        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"A4","marginTop":"2cm"}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
        var order = sectPr.ChildElements.Select(c => c.GetType().Name).ToList();
        Assert.True(
            order.IndexOf("HeaderReference") < order.IndexOf("PageSize"),
            string.Join(",", order));
        Assert.True(
            order.IndexOf("PageSize") < order.IndexOf("PageMargin"),
            string.Join(",", order));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Multi_section_documents_address_each_section_independently()
    {
        var file = CreateDoc(title: "Two sections");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"A4"}}]""");

        // Simulate a Word-authored section break: a paragraph-level sectPr ahead of the body one.
        using (var doc = WordprocessingDocument.Open(file, isEditable: true))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
            var pPr = paragraph.ParagraphProperties ??= new ParagraphProperties();
            pPr.SectionProperties = new SectionProperties();
        }

        Edit(file, """[{"op":"set","path":"/section[2]","props":{"pageSize":"A3"}}]""");

        Assert.Equal(2, GetSection(file, "/section[1]")["properties"]!["sections"]!.GetValue<int>());
        Assert.Null(GetSection(file, "/section[1]")["properties"]!["pageSize"]); // the break section is untouched
        Assert.Equal("A3", GetSection(file, "/section[2]")["properties"]!["pageSize"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Unknown_section_prop_is_unsupported_with_candidates()
    {
        var file = CreateDoc(title: "Props");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/section[1]","props":{"columns":2}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("orientation", ex.Candidates!);
    }

    [Fact]
    public void Unknown_page_size_is_invalid_args_with_candidates()
    {
        var file = CreateDoc(title: "Sizes");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/section[1]","props":{"pageSize":"A7"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("A4", ex.Candidates!);
    }

    [Fact]
    public void Missing_section_index_is_invalid_path()
    {
        var file = CreateDoc(title: "One section");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/section[3]","props":{"pageSize":"A4"}}]"""));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/section[1]", ex.Candidates!);
    }
}
