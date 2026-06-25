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

    // ------------------------------------------------------------ tracked add

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

    /// <summary>A doc whose p[1] holds the author's "Dear " run, the merge-field anchor.</summary>
    private string MergeAnchorDoc(string name = "doc.docx")
    {
        var file = CreateDoc(name, title: "Letter");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Dear "}}]""");
        return file;
    }

    [Fact]
    public void Tracked_merge_field_add_wraps_only_the_new_runs_in_a_single_ins()
    {
        var file = MergeAnchorDoc();

        Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FirstName"}}]""",
            Track("Reviewer"));

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var anchor = body.Elements<Paragraph>().First(); // p[1]

        // Exactly one w:ins, carrying the five complex-field runs (begin / instr /
        // separate / «result» / end); three of them carry a FieldChar.
        var ins = Assert.Single(anchor.Elements<InsertedRun>());
        Assert.Equal("Reviewer", ins.Author!.Value);
        Assert.NotNull(ins.Id);
        Assert.NotNull(ins.Date);
        Assert.Equal(5, ins.Elements<Run>().Count());
        Assert.Equal(3, ins.Elements<Run>().Count(r => r.GetFirstChild<FieldChar>() is not null));
        Assert.Contains("MERGEFIELD FirstName", ins.Descendants<FieldCode>().Single().Text, StringComparison.Ordinal);
        Assert.Contains("«FirstName»", ins.InnerText, StringComparison.Ordinal);

        // The anchor's pre-existing "Dear " run is NOT wrapped, and the paragraph
        // mark is untouched (no inserted run-mark properties).
        Assert.Single(anchor.Elements<Run>()); // only the pre-existing author run is a direct child
        Assert.Equal("Dear ", anchor.Elements<Run>().Single().InnerText);
        Assert.Null(anchor.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Inserted>());

        AssertValidatesClean(file);
    }

    [Fact]
    public void Tracked_merge_field_add_uses_a_unique_revision_id_per_field()
    {
        var file = MergeAnchorDoc();

        Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FirstName"}}]""",
            Track());
        Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"LastName"}}]""",
            Track());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var ids = doc.MainDocumentPart!.Document!.Body!.Descendants<InsertedRun>()
            .Select(i => i.Id!.Value).ToList();
        Assert.Equal(2, ids.Count);
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void Tracked_merge_field_add_reports_exactly_one_insert_revision()
    {
        var file = MergeAnchorDoc();

        Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FirstName"}}]""",
            Track());

        var revision = Assert.Single(Revisions(file));
        Assert.Equal("insert", revision!["kind"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", revision["at"]!.GetValue<string>());
        Assert.Null(revision["mark"]); // not a paragraph-mark insertion
        AssertValidatesClean(file);
    }

    [Fact]
    public void Untracked_merge_field_add_stays_on_the_byte_stable_legacy_path()
    {
        // The untracked branch is first-in-code and must not change: no w:ins wraps
        // the field, and the five complex-field runs are direct children of the
        // anchor paragraph. Two independent untracked runs are byte-identical.
        var a = MergeAnchorDoc("a.docx");
        var b = MergeAnchorDoc("b.docx");
        const string add =
            """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FirstName"}}]""";
        Edit(a, add);
        Edit(b, add);

        Assert.True(DocumentXml(a).SequenceEqual(DocumentXml(b)));

        using var doc = WordprocessingDocument.Open(a, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Empty(body.Descendants<InsertedRun>());

        // The anchor paragraph keeps its author run plus the five field runs as
        // direct children (no w:ins); three of them carry a FieldChar.
        var anchor = body.Elements<Paragraph>().First();
        Assert.Equal(3, anchor.Elements<Run>().Count(r => r.GetFirstChild<FieldChar>() is not null));
        Assert.Empty(anchor.Descendants<InsertedRun>());
        AssertValidatesClean(a);
    }

    [Fact]
    public void Tracked_merge_field_accepts_to_kept_field_and_rejects_to_gone()
    {
        var fileAccept = MergeAnchorDoc("accept.docx");
        Edit(fileAccept,
            """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FirstName"}}]""",
            Track());
        Edit(fileAccept, """[{"op":"accept","path":"/body"}]""");

        Assert.Empty(Revisions(fileAccept));
        using (var doc = WordprocessingDocument.Open(fileAccept, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            // Accept unwraps the w:ins keeping the MERGEFIELD intact.
            Assert.Empty(body.Descendants<InsertedRun>());
            Assert.Contains(body.Descendants<FieldCode>().Select(c => c.Text),
                t => t.Contains("MERGEFIELD FirstName", StringComparison.Ordinal));
        }
        AssertValidatesClean(fileAccept);

        var fileReject = MergeAnchorDoc("reject.docx");
        Edit(fileReject,
            """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FirstName"}}]""",
            Track());
        Edit(fileReject, """[{"op":"reject","path":"/body"}]""");

        Assert.Empty(Revisions(fileReject));
        using (var doc = WordprocessingDocument.Open(fileReject, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            // Reject drops the w:ins and the field runs it wrapped.
            Assert.Empty(body.Descendants<InsertedRun>());
            Assert.DoesNotContain(body.Descendants<FieldCode>().Select(c => c.Text),
                t => t.Contains("MERGEFIELD", StringComparison.Ordinal));
        }
        AssertValidatesClean(fileReject);
    }

    [Fact]
    public void Tracked_merge_field_with_find_stays_unsupported()
    {
        // The mid-paragraph (props.find) path splices runs into an existing
        // paragraph — too large a change for CT_Ins, so it stays refused.
        var file = CreateDoc(title: "Letter");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"Name: here"}}]""");

        var ex = Assert.Throws<AiofficeException>(() => Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"mergeField","props":{"name":"FullName","find":"here"}}]""",
            Track()));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }
}
