using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class MergeFieldTests : WordTestBase
{
    private static JsonNode FirstOp(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["data"]!["ops"]!.AsArray()[0]!;

    private static List<string> WarningCodes(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"] is JsonArray warnings
            ? [.. warnings.Select(w => w!["code"]!.GetValue<string>())]
            : [];

    [Fact]
    public void Add_merge_field_emits_a_mergefield_with_chevron_cache()
    {
        var file = CreateDoc(title: "Letter");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Dear "}}]""");

        var envelope = Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FirstName"}}]""");

        var op = FirstOp(envelope);
        Assert.Equal("mergeField", op["type"]!.GetValue<string>());
        Assert.Equal("FirstName", op["name"]!.GetValue<string>());
        Assert.Equal("«FirstName»", op["cached"]!.GetValue<string>());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var p = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var instr = Assert.Single(p.Descendants<FieldCode>(), c => c.Text.Contains("MERGEFIELD"));
        Assert.Contains("MERGEFIELD FirstName", instr.Text, StringComparison.Ordinal);
        Assert.Contains("«FirstName»", p.InnerText, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_merge_field_reports_its_name_and_value()
    {
        var file = CreateDoc(title: "Letter");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"City"}}]""");

        var data = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/mergeField[@name=City]" })));

        Assert.Equal("mergeField", data["type"]!.GetValue<string>());
        Assert.Equal("City", data["properties"]!["name"]!.GetValue<string>());
        Assert.Equal("«City»", data["properties"]!["value"]!.GetValue<string>());
    }

    [Fact]
    public void Fields_view_lists_merge_fields_with_current_value()
    {
        var file = CreateDoc(title: "Letter");
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FirstName"}},
              {"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"LastName"}}
            ]
            """);

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "fields" })));
        var merge = data["fields"]!.AsArray()
            .Where(f => f!["kind"]!.GetValue<string>() == "mergeField")
            .ToList();

        Assert.Equal(2, merge.Count);
        Assert.Equal("FirstName", merge[0]!["name"]!.GetValue<string>());
        Assert.Equal("«FirstName»", merge[0]!["value"]!.GetValue<string>());
        Assert.Equal("LastName", merge[1]!["name"]!.GetValue<string>());
    }

    [Fact]
    public void Template_fills_both_braces_and_mergefields_from_one_data_file()
    {
        // A doc carrying a {{greeting}} marker and a MERGEFIELD FirstName.
        var file = CreateDoc(title: "Letter");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"{{greeting}}, "}}]""");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FirstName"}}]""");

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["greeting"] = "Hello", ["FirstName"] = "Ada" },
        }));

        var data = Data(envelope);
        // Both the {{greeting}} marker and the MERGEFIELD were filled.
        Assert.Equal(2, data["replaced"]!.GetValue<int>());

        var text = BodyTexts(file)[0];
        Assert.Contains("Hello", text, StringComparison.Ordinal);
        Assert.Contains("Ada", text, StringComparison.Ordinal);
        Assert.DoesNotContain("«FirstName»", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{greeting}}", text, StringComparison.Ordinal);

        // The MERGEFIELD instruction survives — only its cached result changed.
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Single(doc.MainDocumentPart!.Document!.Descendants<FieldCode>(), c => c.Text.Contains("MERGEFIELD FirstName"));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Template_reports_unresolved_merge_field_names()
    {
        var file = CreateDoc(title: "Letter");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Dear "}}]""");
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FirstName"}},
              {"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"LastName"}}
            ]
            """);

        // Only FirstName is supplied; LastName must be reported unresolved.
        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["FirstName"] = "Ada" },
        }));

        var data = Data(envelope);
        Assert.Equal(1, data["replaced"]!.GetValue<int>());
        Assert.Contains("LastName", data["unresolved"]!.AsArray().Select(n => n!.GetValue<string>()));
        Assert.Contains("unresolved_keys", WarningCodes(envelope));

        // The unfilled merge field keeps its chevron placeholder.
        Assert.Contains("«LastName»", BodyTexts(file)[0], StringComparison.Ordinal);
    }

    [Fact]
    public void Find_places_the_merge_field_before_matched_text()
    {
        var file = CreateDoc(title: "Letter");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Name: here"}}]""");

        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FullName","find":"here"}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var p = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var children = p.ChildElements.ToList();
        var beginIndex = children.FindIndex(c =>
            c is Run r && r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Begin);
        var hereIndex = children.FindIndex(c => c.InnerText == "here");
        Assert.True(beginIndex >= 0 && beginIndex < hereIndex);
    }

    [Fact]
    public void Invalid_merge_field_name_is_invalid_args()
    {
        var file = CreateDoc(title: "Letter");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"First Name"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }
}
