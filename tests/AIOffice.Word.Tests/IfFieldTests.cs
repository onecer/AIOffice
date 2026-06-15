using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class IfFieldTests : WordTestBase
{
    private static JsonNode FirstOp(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["data"]!["ops"]!.AsArray()[0]!;

    private string CreateDocWithIfField(
        string field = "Country", string op = "=", string value = "US",
        string trueText = "Domestic", string falseText = "International")
    {
        var file = CreateDoc(title: "Letter");
        var ops = """[{"op":"add","path":"/body/p[1]","type":"ifField","props":{"field":"FIELD","operator":"OP","value":"VALUE","trueText":"TRUE","falseText":"FALSE"}}]"""
            .Replace("FIELD", field, StringComparison.Ordinal)
            .Replace("OP", op, StringComparison.Ordinal)
            .Replace("VALUE", value, StringComparison.Ordinal)
            .Replace("TRUE", trueText, StringComparison.Ordinal)
            .Replace("FALSE", falseText, StringComparison.Ordinal);
        Edit(file, ops);
        return file;
    }

    [Fact]
    public void Add_if_field_emits_a_complex_if_field_with_false_branch_cached()
    {
        var file = CreateDoc(title: "Letter");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"ifField","props":{"field":"Country","operator":"=","value":"US","trueText":"Domestic","falseText":"International"}}]
            """);

        var op = FirstOp(envelope);
        Assert.Equal("ifField", op["type"]!.GetValue<string>());
        Assert.Equal("Country", op["field"]!.GetValue<string>());
        Assert.Equal("International", op["cached"]!.GetValue<string>());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var p = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var instr = Assert.Single(p.Descendants<FieldCode>(), c => c.Text.Contains("IF "));
        Assert.Contains("IF Country = \"US\"", instr.Text, StringComparison.Ordinal);
        // Empty-merge state shows the false branch.
        Assert.Contains("International", p.InnerText, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_if_field_reports_its_condition()
    {
        var file = CreateDocWithIfField();

        var data = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/ifField[@field=Country]" })));

        Assert.Equal("ifField", data["type"]!.GetValue<string>());
        var props = data["properties"]!;
        Assert.Equal("Country", props["field"]!.GetValue<string>());
        Assert.Equal("=", props["operator"]!.GetValue<string>());
        Assert.Equal("US", props["value"]!.GetValue<string>());
        Assert.Equal("Domestic", props["trueText"]!.GetValue<string>());
        Assert.Equal("International", props["falseText"]!.GetValue<string>());
        Assert.Equal("International", props["result"]!.GetValue<string>());
    }

    [Fact]
    public void Fields_view_lists_the_if_field()
    {
        var file = CreateDocWithIfField();

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "fields" })));
        var ifField = Assert.Single(data["fields"]!.AsArray(), f => f!["kind"]!.GetValue<string>() == "ifField");

        Assert.Equal("Country", ifField!["field"]!.GetValue<string>());
        Assert.Equal("International", ifField!["value"]!.GetValue<string>());
    }

    [Fact]
    public void Template_resolves_the_true_branch_when_condition_holds()
    {
        var file = CreateDocWithIfField();

        Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["Country"] = "US" },
        }));

        var text = BodyTexts(file)[0];
        Assert.Contains("Domestic", text, StringComparison.Ordinal);
        Assert.DoesNotContain("International", text, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Template_resolves_the_false_branch_when_condition_fails()
    {
        var file = CreateDocWithIfField();

        Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["Country"] = "FR" },
        }));

        var text = BodyTexts(file)[0];
        Assert.Contains("International", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Domestic", text, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Numeric_comparison_operators_resolve()
    {
        var file = CreateDocWithIfField(field: "Total", op: ">", value: "100", trueText: "Big", falseText: "Small");

        Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["Total"] = "250" },
        }));

        Assert.Contains("Big", BodyTexts(file)[0], StringComparison.Ordinal);
    }

    [Fact]
    public void Not_equals_operator_resolves()
    {
        var file = CreateDocWithIfField(field: "Status", op: "!=", value: "active", trueText: "Inactive", falseText: "Active");

        Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["Status"] = "frozen" },
        }));

        Assert.Contains("Inactive", BodyTexts(file)[0], StringComparison.Ordinal);
    }

    [Fact]
    public void Missing_field_is_invalid_args()
    {
        var file = CreateDoc(title: "Letter");

        var ex = Assert.Throws<AiofficeException>(() => Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"ifField","props":{"value":"US","trueText":"A","falseText":"B"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Unknown_operator_is_invalid_args()
    {
        var file = CreateDoc(title: "Letter");

        var ex = Assert.Throws<AiofficeException>(() => Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"ifField","props":{"field":"X","operator":"~","value":"1","trueText":"A","falseText":"B"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }
}
