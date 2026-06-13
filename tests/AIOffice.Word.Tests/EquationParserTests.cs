// M10: the LaTeX parser (LatexParser, MathNode, LatexParseResult) moved to the
// shared AIOffice.Core.Equations so both Word and Pptx consume one converter; the
// DocumentFormat.OpenXml emitter (OmmlEmitter) stays in AIOffice.Word.Equations.
using AIOffice.Core.Equations;
using AIOffice.Word.Equations;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using M = DocumentFormat.OpenXml.Math;

namespace AIOffice.Word.Tests;

/// <summary>
/// Unit tests for the self-contained LaTeX→OMML engine: every supported
/// construct parses to the expected tree shape and emits validator-clean OMML.
/// These exercise the parser/emitter directly (the internal subsystem), in
/// addition to the end-to-end equation tests that go through the edit surface.
/// </summary>
public sealed class EquationParserTests
{
    /// <summary>The validator oracle for a bare equation: wrap it in a minimal docx and assert 0 errors.</summary>
    private static void AssertOmmlValidates(string latex)
    {
        var parsed = LatexParser.Parse(latex);
        var oMath = OmmlEmitter.ToOfficeMath(parsed.Root);

        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(oMath.CloneNode(true))));
        }

        ms.Position = 0;
        using var reopened = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(reopened).ToList();
        Assert.True(
            errors.Count == 0,
            $"OMML for '{latex}' had validator errors:\n" + string.Join('\n', errors.Select(e => $"{e.Id}: {e.Description}")));
    }

    private static M.OfficeMath Emit(string latex) => OmmlEmitter.ToOfficeMath(LatexParser.Parse(latex).Root);

    [Theory]
    [InlineData("E = mc^2")]
    [InlineData("x_i")]
    [InlineData("x_i^2")]
    [InlineData("a^{n+1}")]
    [InlineData("\\frac{a+b}{c-d}")]
    [InlineData("\\frac{1}{\\frac{2}{3}}")]
    [InlineData("\\sqrt{2}")]
    [InlineData("\\sqrt[3]{x+1}")]
    [InlineData("\\sqrt[n]{x}")]
    [InlineData("\\sum_{i=1}^{n} i")]
    [InlineData("\\prod_{k=1}^{n} k")]
    [InlineData("\\int_0^1 x \\, dx")]
    [InlineData("\\oint_C f")]
    [InlineData("\\left( \\frac{a}{b} \\right)")]
    [InlineData("\\left[ x \\right]")]
    [InlineData("\\left\\{ x \\right\\}")]
    [InlineData("\\begin{matrix} a & b \\\\ c & d \\end{matrix}")]
    [InlineData("\\begin{pmatrix} 1 & 2 \\\\ 3 & 4 \\end{pmatrix}")]
    [InlineData("\\begin{bmatrix} 1 & 0 \\\\ 0 & 1 \\end{bmatrix}")]
    [InlineData("\\hat{x}")]
    [InlineData("\\bar{x}")]
    [InlineData("\\vec{v} + \\dot{y}")]
    [InlineData("\\alpha + \\beta = \\gamma")]
    [InlineData("\\Gamma(\\Omega)")]
    [InlineData("a \\times b \\cdot c \\pm d")]
    [InlineData("x \\leq y \\geq z \\neq w \\approx v")]
    [InlineData("f \\to \\infty")]
    [InlineData("\\partial \\nabla \\sum")]
    [InlineData("\\sin(x) + \\cos(y)")]
    [InlineData("\\lim_{x \\to 0} \\frac{\\sin x}{x}")]
    [InlineData("\\text{if } x > 0 \\text{ then}")]
    [InlineData("")]
    public void Every_supported_construct_emits_validator_clean_omml(string latex) =>
        AssertOmmlValidates(latex);

    [Fact]
    public void Superscript_emits_msup_object()
    {
        var math = Emit("x^2");
        Assert.Single(math.Elements<M.Superscript>());
    }

    [Fact]
    public void Subscript_emits_msub_object()
    {
        var math = Emit("x_i");
        Assert.Single(math.Elements<M.Subscript>());
    }

    [Fact]
    public void Combined_sub_and_sup_emit_one_subsuperscript()
    {
        var math = Emit("x_i^2");
        Assert.Single(math.Elements<M.SubSuperscript>());
        Assert.Empty(math.Elements<M.Subscript>());
        Assert.Empty(math.Elements<M.Superscript>());
    }

    [Fact]
    public void Fraction_emits_numerator_and_denominator()
    {
        var fraction = Assert.Single(Emit("\\frac{a}{b}").Elements<M.Fraction>());
        Assert.Equal("a", fraction.Numerator!.InnerText);
        Assert.Equal("b", fraction.Denominator!.InnerText);
    }

    [Fact]
    public void Square_root_hides_its_degree()
    {
        var radical = Assert.Single(Emit("\\sqrt{x}").Elements<M.Radical>());
        Assert.Equal(M.BooleanValues.True, radical.RadicalProperties?.GetFirstChild<M.HideDegree>()?.Val?.Value);
    }

    [Fact]
    public void Nth_root_keeps_the_degree()
    {
        var radical = Assert.Single(Emit("\\sqrt[3]{x}").Elements<M.Radical>());
        Assert.Equal("3", radical.Degree!.InnerText);
        Assert.Null(radical.RadicalProperties?.GetFirstChild<M.HideDegree>());
    }

    [Fact]
    public void Nary_sum_carries_operator_and_both_limits()
    {
        var nary = Assert.Single(Emit("\\sum_{i=1}^{n} i").Elements<M.Nary>());
        Assert.Equal("∑", nary.NaryProperties?.GetFirstChild<M.AccentChar>()?.Val?.Value);
        Assert.Equal(M.BooleanValues.False, nary.NaryProperties?.GetFirstChild<M.HideSubArgument>()?.Val?.Value);
        Assert.Equal(M.BooleanValues.False, nary.NaryProperties?.GetFirstChild<M.HideSuperArgument>()?.Val?.Value);
        Assert.Equal("i=1", nary.SubArgument!.InnerText);
        Assert.Equal("n", nary.SuperArgument!.InnerText);
    }

    [Fact]
    public void Nary_integral_without_limits_hides_both()
    {
        var nary = Assert.Single(Emit("\\int x").Elements<M.Nary>());
        Assert.Equal("∫", nary.NaryProperties?.GetFirstChild<M.AccentChar>()?.Val?.Value);
        Assert.Equal(M.BooleanValues.True, nary.NaryProperties?.GetFirstChild<M.HideSubArgument>()?.Val?.Value);
        Assert.Equal(M.BooleanValues.True, nary.NaryProperties?.GetFirstChild<M.HideSuperArgument>()?.Val?.Value);
    }

    [Fact]
    public void Pmatrix_wraps_a_two_by_two_matrix_in_parentheses()
    {
        var math = Emit("\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}");
        var delimiter = Assert.Single(math.Elements<M.Delimiter>());
        Assert.Equal("(", delimiter.DelimiterProperties?.GetFirstChild<M.BeginChar>()?.Val?.Value);
        Assert.Equal(")", delimiter.DelimiterProperties?.GetFirstChild<M.EndChar>()?.Val?.Value);

        var matrix = Assert.Single(delimiter.Descendants<M.Matrix>());
        Assert.Equal(2, matrix.Elements<M.MatrixRow>().Count());
        Assert.All(matrix.Elements<M.MatrixRow>(), row => Assert.Equal(2, row.Elements<M.Base>().Count()));
    }

    [Fact]
    public void Plain_matrix_has_no_surrounding_delimiter()
    {
        var math = Emit("\\begin{matrix} a & b \\\\ c & d \\end{matrix}");
        Assert.Empty(math.Elements<M.Delimiter>());
        Assert.Single(math.Elements<M.Matrix>());
    }

    [Fact]
    public void Accent_emits_an_accent_object_with_its_char()
    {
        var accent = Assert.Single(Emit("\\hat{x}").Elements<M.Accent>());
        Assert.Equal("̂", accent.AccentProperties?.GetFirstChild<M.AccentChar>()?.Val?.Value);
    }

    [Fact]
    public void Bar_emits_a_bar_object()
    {
        Assert.Single(Emit("\\bar{x}").Elements<M.Bar>());
        Assert.Single(Emit("\\overline{x+y}").Elements<M.Bar>());
    }

    [Fact]
    public void Greek_letters_map_to_unicode()
    {
        Assert.Contains("α", Emit("\\alpha").InnerText);
        Assert.Contains("Ω", Emit("\\Omega").InnerText);
        Assert.Contains("π", Emit("\\pi").InnerText);
    }

    [Fact]
    public void Operators_map_to_unicode()
    {
        Assert.Contains("×", Emit("a \\times b").InnerText);
        Assert.Contains("≤", Emit("a \\leq b").InnerText);
        Assert.Contains("→", Emit("a \\to b").InnerText);
        Assert.Contains("∞", Emit("\\infty").InnerText);
    }

    [Fact]
    public void Text_command_emits_an_upright_run()
    {
        var math = Emit("\\text{hello}");
        var run = Assert.Single(math.Elements<M.Run>());
        Assert.NotNull(run.MathRunProperties?.GetFirstChild<M.NormalText>());
        Assert.Equal("hello", run.InnerText);
    }

    [Fact]
    public void Function_names_render_upright()
    {
        var math = Emit("\\sin");
        var run = Assert.Single(math.Elements<M.Run>());
        Assert.NotNull(run.MathRunProperties?.GetFirstChild<M.NormalText>());
        Assert.Equal("sin", run.InnerText);
    }

    [Fact]
    public void Unknown_command_is_reported_and_emitted_literally()
    {
        var parsed = LatexParser.Parse("\\foobar{x} + y");
        Assert.Contains("\\foobar", parsed.UnknownTokens);
        // The literal token survives in the output text so nothing is silently dropped.
        Assert.Contains("\\foobar", OmmlEmitter.ToOfficeMath(parsed.Root).InnerText);
    }

    [Fact]
    public void Multiple_unknown_commands_are_all_reported()
    {
        var parsed = LatexParser.Parse("\\foo + \\bar{x} + \\baz");
        Assert.Contains("\\foo", parsed.UnknownTokens);
        Assert.Contains("\\baz", parsed.UnknownTokens);
        // \bar IS recognized (overbar), so it must not appear as unknown.
        Assert.DoesNotContain("\\bar", parsed.UnknownTokens);
    }

    [Fact]
    public void Known_constructs_report_no_unknown_tokens()
    {
        var parsed = LatexParser.Parse("\\frac{\\sqrt{a}}{\\sum_{i=1}^{n} b_i}");
        Assert.Empty(parsed.UnknownTokens);
    }

    [Fact]
    public void Unbraced_single_token_argument_is_accepted()
    {
        // \sqrt 3 and x^2 take the next single atom as the argument (TeX rule).
        Assert.Single(Emit("\\sqrt 3").Elements<M.Radical>());
        Assert.Single(Emit("x^2").Elements<M.Superscript>());
    }

    [Fact]
    public void Nested_fraction_and_root_stay_validator_clean()
    {
        AssertOmmlValidates("\\frac{\\sqrt[3]{x^2 + 1}}{\\sum_{k=0}^{\\infty} \\frac{1}{k!}}");
    }
}
