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

    // -------------------------------------------- v1.7 STYLEREF/SYMBOL/QUOTE

    [Fact]
    public void Styleref_field_references_a_style_and_reopens()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":""}}]""");

        var envelope = Edit(file, """
            [{"op":"add","path":"/header[1]/p[1]","type":"field","props":{"kind":"styleRef","styleRef":"Heading 1"}}]
            """);
        Assert.Equal("styleRef", Data(envelope)["ops"]!.AsArray()[0]!["kind"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.HeaderParts.Single().Header!.Descendants<SimpleField>().Single();
            Assert.Contains("STYLEREF \"Heading 1\"", field.Instruction!.Value!, StringComparison.Ordinal);
        }

        // get reports the kind via the field-kind reverse lookup.
        var props = Get(file, "/header[1]/p[1]")["properties"]!;
        Assert.Contains("styleRef", props["fields"]!.AsArray().Select(f => f!.GetValue<string>()));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Symbol_field_inserts_a_glyph_by_char_code_and_font()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"symbol","charCode":169,"symbolFont":"Symbol"}}]
            """);
        var cached = Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>();
        Assert.Equal("©", cached); // U+00A9

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("SYMBOL 169", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.Contains("\\f \"Symbol\"", field.Instruction!.Value!, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Symbol_field_accepts_hex_char_codes()
    {
        var file = CreateDoc(title: "Doc");
        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"symbol","charCode":"0xA9"}}]
            """);
        Assert.Equal("©", Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("SYMBOL 169", field.Instruction!.Value!, StringComparison.Ordinal); // hex decoded to decimal
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Quote_field_carries_literal_text_as_its_result()
    {
        var file = CreateDoc(title: "Doc");
        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"quote","quoteText":"Confidential"}}]
            """);
        Assert.Equal("Confidential", Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("QUOTE \"Confidential\"", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.Equal("Confidential", field.InnerText);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Styleref_without_a_style_name_is_invalid_args()
    {
        var file = CreateDoc(title: "Doc");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"styleRef"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("styleRef", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Symbol_with_a_bad_char_code_is_invalid_args()
    {
        var file = CreateDoc(title: "Doc");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"symbol","charCode":"banana"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void New_field_kinds_are_listed_in_the_unknown_kind_candidates()
    {
        var file = CreateDoc(title: "Doc");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"bogus"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("styleRef", ex.Candidates!);
        Assert.Contains("symbol", ex.Candidates!);
        Assert.Contains("quote", ex.Candidates!);
        // The pre-1.7 kinds are still offered (additive, nothing removed).
        Assert.Contains("pageNumber", ex.Candidates!);
        Assert.Contains("date", ex.Candidates!);
        // 1.12 document-info & reference kinds are offered too.
        Assert.Contains("fileName", ex.Candidates!);
        Assert.Contains("numWords", ex.Candidates!);
        Assert.Contains("author", ex.Candidates!);
        Assert.Contains("ref", ex.Candidates!);
        Assert.Contains("hyperlink", ex.Candidates!);
        Assert.Contains("fillIn", ex.Candidates!);
    }

    // --------------------------------------- v1.12 document-info & reference fields

    [Fact]
    public void FileName_field_caches_the_document_file_name()
    {
        var file = CreateDoc("report.docx", title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"fileName"}}]
            """);
        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("fileName", summary["kind"]!.GetValue<string>());
        Assert.Equal("report.docx", summary["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("FILENAME", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.DoesNotContain("\\p", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.Equal("report.docx", field.InnerText);
        }

        var props = Get(file, "/body/p[1]")["properties"]!;
        Assert.Contains("fileName", props["fields"]!.AsArray().Select(f => f!.GetValue<string>()));
        AssertValidatesClean(file);
    }

    [Fact]
    public void FileName_field_with_include_path_adds_the_p_switch_and_full_path()
    {
        var file = CreateDoc("report.docx", title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"fileName","includePath":true}}]
            """);
        var cached = Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>();
        Assert.Equal(Path.GetFullPath(file), cached);
        Assert.EndsWith("report.docx", cached, StringComparison.Ordinal);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("FILENAME \\p", field.Instruction!.Value!, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void NumWords_field_caches_the_body_word_count()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"one two three four five"}}]""");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"numWords"}}]
            """);
        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("numWords", summary["kind"]!.GetValue<string>());
        Assert.Equal("5", summary["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("NUMWORDS", field.Instruction!.Value!, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void NumChars_field_caches_the_body_character_count()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"abcde"}}]""");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"numChars"}}]
            """);
        var cached = Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>();
        Assert.Equal("5", cached);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("NUMCHARS", field.Instruction!.Value!, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Author_field_caches_the_core_property_creator()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"set","path":"/properties","props":{"author":"Ada Lovelace"}}]""");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"author"}}]
            """);
        var cached = Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>();
        Assert.Equal("Ada Lovelace", cached);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("AUTHOR", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.Equal("Ada Lovelace", field.InnerText);
        }

        // props.author is the author override here, not stripped batch-author noise.
        var props = Get(file, "/body/p[1]")["properties"]!;
        Assert.Contains("author", props["fields"]!.AsArray().Select(f => f!.GetValue<string>()));
        AssertValidatesClean(file);
    }

    [Fact]
    public void CreateDate_field_caches_the_core_created_date_with_the_format_picture()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"set","path":"/properties","props":{"created":"2026-06-13T10:00:00Z"}}]""");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"createDate","format":"yyyy"}}]
            """);
        Assert.Equal("2026", Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("CREATEDATE \\@ \"yyyy\"", field.Instruction!.Value!, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void SaveDate_field_caches_the_core_modified_date()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"set","path":"/properties","props":{"modified":"2026-06-14T08:00:00Z"}}]""");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"saveDate","format":"yyyy-MM-dd"}}]
            """);
        Assert.Equal("2026-06-14", Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("SAVEDATE", field.Instruction!.Value!, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void PrintDate_field_writes_the_field_and_caches_now()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"printDate","format":"yyyy"}}]
            """);
        Assert.Equal(
            DateTime.Now.Year.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("PRINTDATE \\@ \"yyyy\"", field.Instruction!.Value!, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Ref_field_caches_the_bookmark_text()
    {
        var file = CreateDoc(title: "Doc");
        // Append a text paragraph and bookmark it (the bookmark wraps the paragraph
        // whose text a mode-text REF resolves to); the REF field lives elsewhere.
        Edit(file, """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"The quarter beat plan."}}
            ]
            """);
        var summaryParaIndex = BodyTexts(file)
            .Select((t, i) => (t, i))
            .First(x => x.t == "The quarter beat plan.").i + 1;
        Edit(file, "[{\"op\":\"add\",\"path\":\"/body/p[" + summaryParaIndex +
            "]\",\"type\":\"bookmark\",\"props\":{\"name\":\"Summary\"}}]");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"ref","bookmark":"Summary"}}]
            """);
        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("ref", summary["kind"]!.GetValue<string>());
        Assert.Equal("The quarter beat plan.", summary["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("REF Summary", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.Equal("The quarter beat plan.", field.InnerText);
        }

        var props = Get(file, "/body/p[1]")["properties"]!;
        Assert.Contains("ref", props["fields"]!.AsArray().Select(f => f!.GetValue<string>()));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Ref_field_page_mode_caches_a_page_placeholder()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Target"}}]""");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"ref","bookmark":"Target","mode":"page"}}]
            """);
        Assert.Equal("1", Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("REF Target \\p", field.Instruction!.Value!, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Ref_without_a_bookmark_is_invalid_args()
    {
        var file = CreateDoc(title: "Doc");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"ref"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("bookmark", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Hyperlink_field_caches_the_link_text()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"hyperlink","url":"https://example.com","linkText":"Example"}}]
            """);
        Assert.Equal("Example", Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("HYPERLINK \"https://example.com\"", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.Equal("Example", field.InnerText);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Hyperlink_field_without_link_text_caches_the_url()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"hyperlink","url":"https://example.com"}}]
            """);
        Assert.Equal("https://example.com", Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Hyperlink_without_a_url_is_invalid_args()
    {
        var file = CreateDoc(title: "Doc");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"hyperlink"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("url", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FillIn_field_caches_the_default_text()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"fillIn","prompt":"Enter your name","default":"Name here"}}]
            """);
        Assert.Equal("Name here", Data(envelope)["ops"]!.AsArray()[0]!["cached"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var field = doc.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single();
            Assert.Contains("FILLIN \"Enter your name\"", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.Contains("\\d \"Name here\"", field.Instruction!.Value!, StringComparison.Ordinal);
            Assert.Equal("Name here", field.InnerText);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void FillIn_without_a_prompt_is_invalid_args()
    {
        var file = CreateDoc(title: "Doc");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"field","props":{"kind":"fillIn","default":"x"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("prompt", ex.Message, StringComparison.Ordinal);
    }
}
