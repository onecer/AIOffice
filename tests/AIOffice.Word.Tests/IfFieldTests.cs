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

    // ------------------------------------------------------------ tracked add

    private const string AddIfField =
        """[{"op":"add","path":"/body/p[1]","type":"ifField","props":{"field":"Country","operator":"=","value":"US","trueText":"Domestic","falseText":"International"}}]""";

    private static JsonObject Track(string? author = null)
    {
        var args = new JsonObject { ["track"] = true };
        if (author is not null)
        {
            args["author"] = author;
        }

        return args;
    }

    private JsonArray Revisions(string file) =>
        Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "revisions" })))["revisions"]!.AsArray();

    /// <summary>The raw bytes of the main document part (word/document.xml).</summary>
    private static byte[] DocumentXml(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        using var stream = doc.MainDocumentPart!.GetStream();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    /// <summary>A doc whose p[1] holds the author's "Ship to: " run, the IF-field anchor.</summary>
    private string IfAnchorDoc(string name = "doc.docx")
    {
        var file = CreateDoc(name, title: "Letter");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Ship to: "}}]""");
        return file;
    }

    [Fact]
    public void Tracked_if_field_add_wraps_only_the_new_runs_in_a_single_ins()
    {
        var file = IfAnchorDoc();

        Edit(file, AddIfField, Track("Reviewer"));

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var anchor = body.Elements<Paragraph>().First(); // p[1]

        // Exactly one w:ins, carrying the five IF complex-field runs (begin / instr /
        // separate / cached / end); three of them carry a FieldChar.
        var ins = Assert.Single(anchor.Elements<InsertedRun>());
        Assert.Equal("Reviewer", ins.Author!.Value);
        Assert.NotNull(ins.Id);
        Assert.NotNull(ins.Date);
        Assert.Equal(5, ins.Elements<Run>().Count());
        Assert.Equal(3, ins.Elements<Run>().Count(r => r.GetFirstChild<FieldChar>() is not null));
        Assert.Contains("IF Country = \"US\"", ins.Descendants<FieldCode>().Single().Text, StringComparison.Ordinal);
        Assert.Contains("International", ins.InnerText, StringComparison.Ordinal); // empty-merge false branch

        // The anchor's pre-existing "Ship to: " run is NOT wrapped, and the
        // paragraph mark is untouched (no inserted run-mark properties).
        Assert.Single(anchor.Elements<Run>()); // only the pre-existing author run is a direct child
        Assert.Equal("Ship to: ", anchor.Elements<Run>().Single().InnerText);
        Assert.Null(anchor.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Inserted>());

        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_if_field_add_reports_exactly_one_insert_revision()
    {
        var file = IfAnchorDoc();

        Edit(file, AddIfField, Track());

        var revision = Assert.Single(Revisions(file));
        Assert.Equal("insert", revision!["kind"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", revision["at"]!.GetValue<string>());
        Assert.Null(revision["mark"]); // not a paragraph-mark insertion
        AssertValidatesClean(file);
    }

    [Fact]
    public void Untracked_if_field_add_stays_on_the_byte_stable_legacy_path()
    {
        // The untracked branch is first-in-code and must not change: no w:ins wraps
        // the field. Two independent untracked runs are byte-identical.
        var a = IfAnchorDoc("a.docx");
        var b = IfAnchorDoc("b.docx");
        Edit(a, AddIfField);
        Edit(b, AddIfField);

        Assert.True(DocumentXml(a).SequenceEqual(DocumentXml(b)));

        using var doc = WordprocessingDocument.Open(a, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Empty(body.Descendants<InsertedRun>());

        var anchor = body.Elements<Paragraph>().First();
        Assert.Equal(3, anchor.Elements<Run>().Count(r => r.GetFirstChild<FieldChar>() is not null));
        Assert.Empty(anchor.Descendants<InsertedRun>());
        AssertValidatesClean(a);
    }

    [Fact]
    public void Tracked_if_field_accepts_to_kept_field_and_rejects_to_gone()
    {
        var fileAccept = IfAnchorDoc("accept.docx");
        Edit(fileAccept, AddIfField, Track());
        Edit(fileAccept, """[{"op":"accept","path":"/body"}]""");

        Assert.Empty(Revisions(fileAccept));
        using (var doc = WordprocessingDocument.Open(fileAccept, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            // Accept unwraps the w:ins keeping the IF field intact.
            Assert.Empty(body.Descendants<InsertedRun>());
            Assert.Contains(body.Descendants<FieldCode>().Select(c => c.Text),
                t => t.Contains("IF Country", StringComparison.Ordinal));
        }
        AssertValidatesClean(fileAccept);

        var fileReject = IfAnchorDoc("reject.docx");
        Edit(fileReject, AddIfField, Track());
        Edit(fileReject, """[{"op":"reject","path":"/body"}]""");

        Assert.Empty(Revisions(fileReject));
        using (var doc = WordprocessingDocument.Open(fileReject, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            // Reject drops the w:ins and the field runs it wrapped.
            Assert.Empty(body.Descendants<InsertedRun>());
            Assert.DoesNotContain(body.Descendants<FieldCode>().Select(c => c.Text),
                t => t.Contains("IF ", StringComparison.Ordinal));
        }
        AssertValidatesClean(fileReject);
    }

    [Fact]
    public void Tracked_if_field_with_find_stays_unsupported()
    {
        // The mid-paragraph (props.find) path splices runs into an existing
        // paragraph — too large a change for CT_Ins, so it stays refused.
        var file = CreateDoc(title: "Letter");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Status: here"}}]""");

        var ex = Assert.Throws<AiofficeException>(() => Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"ifField","props":{"field":"Status","value":"x","trueText":"A","falseText":"B","find":"here"}}]""",
            Track()));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }
}
