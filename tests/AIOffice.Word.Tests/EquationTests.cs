using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace AIOffice.Word.Tests;

/// <summary>
/// End-to-end equation tests through the edit surface: add inline/display
/// equations, reopen and verify the OMML lands in the document, the validator
/// stays clean, the original LaTeX reads back faithfully, and get/remove/read
/// all behave. These complement the parser-level <see cref="EquationParserTests"/>.
/// </summary>
public sealed class EquationTests : WordTestBase
{
    private JsonNode AddEquation(string file, string path, string latex, bool display = false, JsonObject? args = null)
    {
        var props = new JsonObject { ["latex"] = latex };
        if (display)
        {
            props["display"] = true;
        }

        var ops = new JsonArray
        {
            new JsonObject { ["op"] = "add", ["path"] = path, ["type"] = "equation", ["props"] = props },
        };
        // The per-op summary (with path/display/latex) is the first entry of data.ops.
        return Data(Edit(file, ops.ToJsonString(), args))["ops"]![0]!;
    }

    private static int CountEquations(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<M.OfficeMath>().Count();
    }

    [Fact]
    public void Inline_equation_appends_omml_to_the_paragraph_and_validates()
    {
        var file = CreateDoc(title: "Eq");
        var result = AddEquation(file, "/body/p[2]", "E = mc^2");

        Assert.Equal("/body/p[2]/omath[1]", result["path"]!.GetValue<string>());
        Assert.False(result["display"]!.GetValue<bool>());
        Assert.Equal(1, CountEquations(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Display_equation_creates_a_centered_math_paragraph()
    {
        var file = CreateDoc(title: "Eq");
        var result = AddEquation(file, "/body", "\\sum_{i=1}^{n} i^2", display: true);

        Assert.True(result["display"]!.GetValue<bool>());
        Assert.EndsWith("/omath[1]", result["path"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            // A display equation lives inside an m:oMathPara with centered justification.
            var mathPara = doc.MainDocumentPart!.Document!.Body!.Descendants<M.Paragraph>().Single();
            Assert.Equal(M.JustificationValues.Center, mathPara.ParagraphProperties?.Justification?.Val?.Value);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Stored_latex_reads_back_verbatim_via_get()
    {
        var file = CreateDoc(title: "Eq");
        const string latex = "\\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}";
        AddEquation(file, "/body/p[2]", latex);

        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[2]/omath[1]" })))["properties"]!;
        Assert.Equal(latex, properties["latex"]!.GetValue<string>());
        Assert.False(properties["display"]!.GetValue<bool>());
    }

    [Fact]
    public void Latex_survives_a_reopen_byte_for_byte_in_meaning()
    {
        var file = CreateDoc(title: "Eq");
        const string latex = "\\int_0^\\infty e^{-x} \\, dx = 1";
        AddEquation(file, "/body/p[2]", latex);

        // Reopen from disk (fresh package) and read the stored LaTeX again.
        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[2]/omath[1]" })))["properties"]!;
        Assert.Equal(latex, properties["latex"]!.GetValue<string>());
    }

    [Fact]
    public void Unknown_tokens_raise_an_equation_partial_warning_but_still_succeed()
    {
        var file = CreateDoc(title: "Eq");
        var envelope = Edit(
            file,
            """[{"op":"add","path":"/body/p[2]","type":"equation","props":{"latex":"\\foobar{x} + y"}}]""");

        var json = JsonNode.Parse(envelope.ToJson())!;
        Assert.True(json["ok"]!.GetValue<bool>());

        var warnings = json["meta"]!["warnings"]!.AsArray();
        var partial = warnings.Single(w => w!["code"]!.GetValue<string>() == "equation_partial");
        Assert.Contains("\\foobar", partial!["message"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Inline_equation_renders_as_dollar_latex_in_text_view()
    {
        var file = CreateDoc(title: "Eq");
        AddEquation(file, "/body/p[2]", "a^2 + b^2");

        var texts = BodyTexts(file);
        Assert.Contains("$a^2 + b^2$", texts);
    }

    [Fact]
    public void Display_equation_renders_as_double_dollar_in_text_view()
    {
        var file = CreateDoc(title: "Eq");
        AddEquation(file, "/body", "\\sqrt{2}", display: true);

        var texts = BodyTexts(file);
        Assert.Contains(texts, t => t.Contains("$$\\sqrt{2}$$"));
    }

    [Fact]
    public void Markdown_view_emits_dollar_math()
    {
        var file = CreateDoc(title: "Eq");
        AddEquation(file, "/body/p[2]", "x_i");
        AddEquation(file, "/body", "\\frac{1}{2}", display: true);

        var markdown = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "markdown" })))["markdown"]!.GetValue<string>();
        Assert.Contains("$x_i$", markdown);
        Assert.Contains("$$\\frac{1}{2}$$", markdown);
    }

    [Fact]
    public void Html_render_emits_mathml()
    {
        var file = CreateDoc(title: "Eq");
        AddEquation(file, "/body/p[2]", "x^2");

        var html = Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();
        Assert.Contains("<math", html);
        Assert.Contains("<msup>", html);
        Assert.Contains("class=\"equation\"", html);
        // The source LaTeX is preserved as a title for accessibility/fallback.
        Assert.Contains("title=\"x^2\"", html);
    }

    [Fact]
    public void Second_inline_equation_is_addressed_as_omath_2()
    {
        var file = CreateDoc(title: "Eq");
        AddEquation(file, "/body/p[2]", "a");
        var second = AddEquation(file, "/body/p[2]", "b");
        Assert.Equal("/body/p[2]/omath[2]", second["path"]!.GetValue<string>());
        Assert.Equal(2, CountEquations(file));
    }

    [Fact]
    public void Remove_drops_an_inline_equation()
    {
        var file = CreateDoc(title: "Eq");
        AddEquation(file, "/body/p[2]", "a");
        AddEquation(file, "/body/p[2]", "b");

        Edit(file, """[{"op":"remove","path":"/body/p[2]/omath[1]"}]""");
        Assert.Equal(1, CountEquations(file));

        // The survivor (originally omath[2]) is now omath[1] and still reads back.
        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[2]/omath[1]" })))["properties"]!;
        Assert.Equal("b", properties["latex"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_drops_a_display_equation_paragraph()
    {
        var file = CreateDoc(title: "Eq");
        AddEquation(file, "/body", "\\pi", display: true);
        Assert.Equal(1, CountEquations(file));

        Edit(file, """[{"op":"remove","path":"/body/p[3]/omath[1]"}]""");
        Assert.Equal(0, CountEquations(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Missing_latex_is_a_clear_invalid_args_error()
    {
        var file = CreateDoc(title: "Eq");
        var ex = Assert.Throws<AIOffice.Core.AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[2]","type":"equation","props":{}}]"""));

        Assert.Equal(AIOffice.Core.ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotEmpty(ex.Suggestion);
    }

    [Fact]
    public void Get_on_a_missing_equation_index_lists_candidates()
    {
        var file = CreateDoc(title: "Eq");
        AddEquation(file, "/body/p[2]", "a");

        var ex = Assert.Throws<AIOffice.Core.AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[2]/omath[5]" })));

        Assert.Equal(AIOffice.Core.ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/body/p[2]/omath[1]", ex.Candidates!);
    }

    [Fact]
    public void Tracked_equation_add_is_unsupported_with_a_workaround()
    {
        var file = CreateDoc(title: "Eq");
        var ex = Assert.Throws<AIOffice.Core.AiofficeException>(() => Edit(
            file,
            """[{"op":"add","path":"/body/p[2]","type":"equation","props":{"latex":"x"}}]""",
            new JsonObject { ["track"] = true }));

        Assert.Equal(AIOffice.Core.ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.NotEmpty(ex.Suggestion);
    }

    // ------------------------------------------------ v1.7 numbered equations

    private JsonNode AddNumberedEquation(string file, string latex, JsonNode number, string path = "/body")
    {
        var props = new JsonObject { ["latex"] = latex, ["display"] = true, ["number"] = number };
        var ops = new JsonArray
        {
            new JsonObject { ["op"] = "add", ["path"] = path, ["type"] = "equation", ["props"] = props },
        };
        return Data(Edit(file, ops.ToJsonString()))["ops"]![0]!;
    }

    [Fact]
    public void Numbered_equation_auto_increments_and_reopens_with_its_label()
    {
        var file = CreateDoc(title: "Eq");
        var first = AddNumberedEquation(file, "a^2 + b^2 = c^2", true);
        var second = AddNumberedEquation(file, "E = mc^2", true);

        Assert.Equal("(1)", first["number"]!.GetValue<string>());
        Assert.Equal("(2)", second["number"]!.GetValue<string>());

        // Reopen and confirm the number reads back via the equation's omath path.
        var firstPath = first["path"]!.GetValue<string>();
        var props = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = firstPath })))["properties"]!;
        Assert.Equal("(1)", props["number"]!.GetValue<string>());
        Assert.True(props["display"]!.GetValue<bool>());
        Assert.Equal("a^2 + b^2 = c^2", props["latex"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Explicit_number_label_is_stored_verbatim()
    {
        var file = CreateDoc(title: "Eq");
        var result = AddNumberedEquation(file, "x = 1", "(1.1)");
        Assert.Equal("(1.1)", result["number"]!.GetValue<string>());

        var props = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = result["path"]!.GetValue<string>() })))["properties"]!;
        Assert.Equal("(1.1)", props["number"]!.GetValue<string>());
    }

    [Fact]
    public void Auto_number_continues_past_an_explicit_label()
    {
        var file = CreateDoc(title: "Eq");
        AddNumberedEquation(file, "x = 1", "(3)");      // explicit
        var next = AddNumberedEquation(file, "y = 2", true); // auto continues from 3
        Assert.Equal("(4)", next["number"]!.GetValue<string>());
    }

    [Fact]
    public void Numbered_equation_paragraph_holds_the_equation_and_a_visible_number()
    {
        var file = CreateDoc(title: "Eq");
        AddNumberedEquation(file, "F = ma", "(2.7)");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        // The equation and its number share one tab-aligned paragraph (no oMathPara).
        var paragraph = body.Descendants<Paragraph>().Single(p => p.Descendants<M.OfficeMath>().Any());
        Assert.Empty(paragraph.Descendants<M.Paragraph>()); // not wrapped in an m:oMathPara
        Assert.Contains("(2.7)", paragraph.InnerText, StringComparison.Ordinal);
        Assert.NotNull(paragraph.ParagraphProperties?.GetFirstChild<Tabs>()); // tab-stop layout
    }

    [Fact]
    public void Numbered_equation_is_addressable_by_its_number()
    {
        var file = CreateDoc(title: "Eq");
        AddNumberedEquation(file, "a = b", "(1.1)");
        AddNumberedEquation(file, "c = d", "(1.2)");

        // Address by the verbatim label "(1.2)" and by the bare number "1.2"; both resolve.
        var byLabel = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/equation[@num=(1.2)]" })))["properties"]!;
        Assert.Equal("(1.2)", byLabel["number"]!.GetValue<string>());
        Assert.Equal("c = d", byLabel["latex"]!.GetValue<string>());
        Assert.EndsWith("/omath[1]", byLabel["path"]!.GetValue<string>());

        var byBareNumber = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/equation[@num=1.2]" })))["properties"]!;
        Assert.Equal("(1.2)", byBareNumber["number"]!.GetValue<string>());
        Assert.Equal("c = d", byBareNumber["latex"]!.GetValue<string>());
    }

    [Fact]
    public void Missing_equation_number_lists_existing_numbers_as_candidates()
    {
        var file = CreateDoc(title: "Eq");
        AddNumberedEquation(file, "a = b", "(1.1)");

        var ex = Assert.Throws<AIOffice.Core.AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/equation[@num=9.9]" })));
        Assert.Equal(AIOffice.Core.ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/equation[@num=(1.1)]", ex.Candidates!);
    }

    [Fact]
    public void Numbered_equation_raises_the_cached_warning()
    {
        var file = CreateDoc(title: "Eq");
        var envelope = Edit(
            file,
            """[{"op":"add","path":"/body","type":"equation","props":{"latex":"x","display":true,"number":true}}]""");

        var warnings = JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"]!.AsArray();
        Assert.Contains(warnings, w => w!["code"]!.GetValue<string>() == "equation_numbers_cached");
    }

    [Fact]
    public void Number_on_an_inline_equation_is_invalid_args()
    {
        var file = CreateDoc(title: "Eq");
        var ex = Assert.Throws<AIOffice.Core.AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[2]","type":"equation","props":{"latex":"x","number":true}}]"""));

        Assert.Equal(AIOffice.Core.ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("display:true", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Removing_a_numbered_equation_drops_its_whole_paragraph()
    {
        var file = CreateDoc(title: "Eq");
        var result = AddNumberedEquation(file, "x = 1", "(1)");
        Assert.Equal(1, CountEquations(file));

        Edit(file, $$"""[{"op":"remove","path":"{{result["path"]!.GetValue<string>()}}"}]""");
        Assert.Equal(0, CountEquations(file));
        // The number text leaves with the paragraph — no orphan "(1)" run remains.
        Assert.DoesNotContain(BodyTexts(file), t => t.Contains("(1)", StringComparison.Ordinal));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Numbered_equation_renders_as_block_math_with_its_label_in_markdown()
    {
        var file = CreateDoc(title: "Eq");
        AddNumberedEquation(file, "\\frac{1}{2}", "(1.1)");

        var markdown = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "markdown" })))["markdown"]!.GetValue<string>();
        Assert.Contains("$$\\frac{1}{2}$$ (1.1)", markdown, StringComparison.Ordinal);
        // The number must appear exactly once (no duplicate orphan run rendering).
        Assert.Equal(1, CountOccurrences(markdown, "(1.1)"));
    }

    [Fact]
    public void A_deepened_construct_round_trips_in_a_numbered_equation()
    {
        var file = CreateDoc(title: "Eq");
        const string latex = "\\begin{cases} 1 & x>0 \\\\ 0 & x \\leq 0 \\end{cases}";
        var result = AddNumberedEquation(file, latex, true);

        var props = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = result["path"]!.GetValue<string>() })))["properties"]!;
        Assert.Equal(latex, props["latex"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        var count = 0;
        var index = 0;
        while ((index = haystack.IndexOf(needle, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += needle.Length;
        }

        return count;
    }
}
