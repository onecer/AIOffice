using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class MailMergeTests : WordTestBase
{
    /// <summary>A letter carrying a {{greeting}} marker and a MERGEFIELD Name.</summary>
    private string CreateLetter(string name = "letter.docx")
    {
        var file = CreateDoc(name, title: "Letter");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"{{greeting}}, "}}]""");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"Name"}}]""");
        return file;
    }

    private static List<string> Produced(Envelope envelope) =>
        [.. Data(envelope)["produced"]!.AsArray().Select(p => p!.GetValue<string>())];

    private static List<string> WarningCodes(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"] is JsonArray warnings
            ? [.. warnings.Select(w => w!["code"]!.GetValue<string>())]
            : [];

    [Fact]
    public void Array_with_output_pattern_produces_one_doc_per_record()
    {
        var file = CreateLetter();
        var before = File.ReadAllBytes(file);

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonArray
            {
                new JsonObject { ["greeting"] = "Hello", ["Name"] = "Ada" },
                new JsonObject { ["greeting"] = "Hi", ["Name"] = "Grace" },
                new JsonObject { ["greeting"] = "Dear", ["Name"] = "Linus" },
            },
            ["output"] = "out-{n}.docx",
        }));

        var data = Data(envelope);
        Assert.Equal(3, data["records"]!.GetValue<int>());
        var produced = Produced(envelope);
        Assert.Equal(3, produced.Count);

        // The source letter is left untouched (output went to new files).
        Assert.Equal(before, File.ReadAllBytes(file));

        // Each output carries its own merged values.
        Assert.Contains("Hello", BodyTexts(produced[0])[0], StringComparison.Ordinal);
        Assert.Contains("Ada", BodyTexts(produced[0])[0], StringComparison.Ordinal);
        Assert.Contains("Grace", BodyTexts(produced[1])[0], StringComparison.Ordinal);
        Assert.Contains("Linus", BodyTexts(produced[2])[0], StringComparison.Ordinal);

        // No chevron placeholders survive a fully-supplied merge.
        Assert.DoesNotContain("«Name»", BodyTexts(produced[0])[0], StringComparison.Ordinal);

        foreach (var path in produced)
        {
            Assert.EndsWith(".docx", path, StringComparison.Ordinal);
            AssertValidatesClean(path);
        }
    }

    [Fact]
    public void Output_pattern_can_substitute_a_record_field()
    {
        var file = CreateLetter();

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonArray
            {
                new JsonObject { ["greeting"] = "Hello", ["Name"] = "Ada" },
                new JsonObject { ["greeting"] = "Hi", ["Name"] = "Grace" },
            },
            ["output"] = "letters/{Name}.docx",
        }));

        var produced = Produced(envelope);
        Assert.Equal(2, produced.Count);
        Assert.EndsWith(Path.Combine("letters", "Ada.docx"), produced[0], StringComparison.Ordinal);
        Assert.EndsWith(Path.Combine("letters", "Grace.docx"), produced[1], StringComparison.Ordinal);
        Assert.True(File.Exists(produced[0]));
        Assert.True(File.Exists(produced[1]));
    }

    [Fact]
    public void Escaping_output_pattern_is_sandbox_denied()
    {
        var file = CreateLetter();

        var ex = Assert.Throws<AiofficeException>(() => Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonArray
            {
                new JsonObject { ["Name"] = "Ada" },
            },
            ["output"] = "../escape-{n}.docx",
        })));

        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
    }

    [Fact]
    public void Array_without_output_makes_one_combined_doc_with_a_section_per_record()
    {
        var file = CreateLetter();

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonArray
            {
                new JsonObject { ["greeting"] = "Hello", ["Name"] = "Ada" },
                new JsonObject { ["greeting"] = "Hi", ["Name"] = "Grace" },
                new JsonObject { ["greeting"] = "Dear", ["Name"] = "Linus" },
            },
        }));

        var data = Data(envelope);
        Assert.Equal(3, data["records"]!.GetValue<int>());
        var produced = Produced(envelope);
        Assert.Single(produced);
        Assert.Equal(file, produced[0]); // combined doc written back to the source

        // The combined document holds every record's merged text.
        var text = string.Join("\n", BodyTexts(file));
        Assert.Contains("Ada", text, StringComparison.Ordinal);
        Assert.Contains("Grace", text, StringComparison.Ordinal);
        Assert.Contains("Linus", text, StringComparison.Ordinal);

        // Three records -> three sections (two next-page breaks + the final body sectPr).
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sections = doc.MainDocumentPart!.Document!.Body!.Descendants<SectionProperties>().Count();
            Assert.Equal(3, sections);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Combined_doc_snapshots_the_pre_image()
    {
        var file = CreateLetter();
        var preRev = Rev.OfFile(file);

        Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonArray
            {
                new JsonObject { ["greeting"] = "Hello", ["Name"] = "Ada" },
                new JsonObject { ["greeting"] = "Hi", ["Name"] = "Grace" },
            },
        }));

        Assert.Contains(Snapshots.List(file), e => e.Rev == preRev);
    }

    [Fact]
    public void Unresolved_fields_raise_a_template_unresolved_warning()
    {
        var file = CreateLetter();

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonArray
            {
                new JsonObject { ["greeting"] = "Hello" }, // Name missing -> unresolved
            },
            ["output"] = "out-{n}.docx",
        }));

        Assert.Contains("template_unresolved", WarningCodes(envelope));
        Assert.Contains("Name", Data(envelope)["unresolved"]!.AsArray().Select(n => n!.GetValue<string>()));
    }

    [Fact]
    public void Single_object_template_still_works_unchanged()
    {
        // The original single-object fill path must behave exactly as before:
        // returns {replaced, keys, unresolved, written}, not the mail-merge shape.
        var file = CreateLetter();

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonObject { ["greeting"] = "Hello", ["Name"] = "Ada" },
        }));

        var data = Data(envelope);
        Assert.Equal(2, data["replaced"]!.GetValue<int>());
        Assert.NotNull(data["written"]);
        Assert.True(data["records"] is null); // not a mail-merge envelope
        Assert.Contains("Ada", BodyTexts(file)[0], StringComparison.Ordinal);
    }

    [Fact]
    public void Empty_records_array_is_invalid_args()
    {
        var file = CreateLetter();

        var ex = Assert.Throws<AiofficeException>(() => Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonArray(),
            ["output"] = "out-{n}.docx",
        })));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void If_field_resolves_per_record_in_a_mail_merge()
    {
        var file = CreateDoc(title: "Letter");
        Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"ifField","props":{"field":"Country","operator":"=","value":"US","trueText":"Domestic","falseText":"International"}}]
            """);

        var envelope = Handler.Template(Ctx(file, new JsonObject
        {
            ["data"] = new JsonArray
            {
                new JsonObject { ["Country"] = "US" },
                new JsonObject { ["Country"] = "FR" },
            },
            ["output"] = "out-{n}.docx",
        }));

        var produced = Produced(envelope);
        Assert.Contains("Domestic", BodyTexts(produced[0])[0], StringComparison.Ordinal);
        Assert.Contains("International", BodyTexts(produced[1])[0], StringComparison.Ordinal);
    }
}
