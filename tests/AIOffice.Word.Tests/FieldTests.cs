using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>M5: w:fldSimple fields in body, headers and footers.</summary>
public sealed class FieldTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    private string CreateWithFooter()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"add","path":"/footer[1]","type":"footer","props":{"text":""}}]""");
        return file;
    }

    [Fact]
    public void Page_number_field_lands_as_fld_simple_with_cached_text()
    {
        var file = CreateWithFooter();

        var envelope = Edit(file, """
            [{"op":"add","path":"/footer[1]/p[1]","type":"field","props":{"kind":"pageNumber"}}]
            """);

        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("pageNumber", summary["kind"]!.GetValue<string>());
        Assert.Equal("1", summary["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.FooterParts.Single().Footer!
                .Descendants<SimpleField>().Single();
            Assert.Contains("PAGE", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.Equal("1", field.InnerText);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Page_x_of_y_composite_builds_from_two_fields_in_one_paragraph()
    {
        var file = CreateWithFooter();

        Edit(file, """
            [
              {"op":"set","path":"/footer[1]/p[1]","props":{"text":"Page ","alignment":"center"}},
              {"op":"add","path":"/footer[1]/p[1]","type":"field","props":{"kind":"pageNumber"}},
              {"op":"add","path":"/footer[1]/p[1]","type":"field","props":{"kind":"numPages","leadingText":" of "}}
            ]
            """);

        // The cached composite reads exactly like the documented pattern promises.
        var props = Get(file, "/footer[1]/p[1]")["properties"]!;
        Assert.Equal("Page 1 of 1", props["text"]!.GetValue<string>());
        Assert.Equal(["pageNumber", "numPages"], props["fields"]!.AsArray().Select(f => f!.GetValue<string>()));

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var instructions = doc.MainDocumentPart!.FooterParts.Single().Footer!
                .Descendants<SimpleField>()
                .Select(f => f.Instruction!.Value!.Trim().Split(' ')[0])
                .ToList();
            Assert.Equal(["PAGE", "NUMPAGES"], instructions);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Date_field_honors_the_format_picture()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"date","format":"yyyy"}}]
            """);

        var cached = Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>();
        Assert.Equal(DateTime.Now.Year.ToString(System.Globalization.CultureInfo.InvariantCulture), cached);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("DATE \\@ \"yyyy\"", field.Instruction!.Value!, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Doc_title_field_caches_a_placeholder_when_no_core_title_exists()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"docTitle"}}]
            """);

        var cached = Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>();
        Assert.False(string.IsNullOrWhiteSpace(cached));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Field_kinds_surface_in_get_and_render()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"pageNumber","leadingText":"p. "}}]
            """);

        var props = Get(file, "/body/p[1]")["properties"]!;
        Assert.Equal(["pageNumber"], props["fields"]!.AsArray().Select(f => f!.GetValue<string>()));

        var html = Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();
        Assert.Contains("p. 1", html, StringComparison.Ordinal); // cached result renders

        var text = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "text" })));
        Assert.Contains("p. 1", text["lines"]!.AsArray()[0]!["text"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_field_kind_is_invalid_args_with_candidates()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"pageCount"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("numPages", ex.Candidates!);
    }

    [Fact]
    public void Field_on_a_non_paragraph_is_invalid_args_pointing_inside()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":1}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/table[1]/tr[1]/tc[1]","type":"field","props":{"kind":"date"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("/p[1]", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Tracked_field_insertion_is_unsupported()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() => Edit(
            file,
            """[{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"pageNumber"}}]""",
            new JsonObject { ["track"] = true }));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }
}
