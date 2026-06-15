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

    // ------------------------------------------------------- v1.7 deepened set

    [Theory]
    [InlineData("\\begin{cases} a & x > 0 \\\\ b & x \\leq 0 \\end{cases}")]
    [InlineData("f(x) = \\begin{cases} 1 \\\\ 0 \\end{cases}")]
    [InlineData("\\begin{aligned} a &= b + c \\\\ &= d \\end{aligned}")]
    [InlineData("\\begin{align} x &= 1 \\\\ y &= 2 \\end{align}")]
    [InlineData("\\begin{alignedat}{2} a &= b & c &= d \\end{alignedat}")]
    [InlineData("\\binom{n}{k}")]
    [InlineData("\\binom{n}{k} + \\binom{n}{k-1}")]
    [InlineData("\\overbrace{a+b+c}^{n}")]
    [InlineData("\\underbrace{x_1 + x_2}_{\\text{sum}}")]
    [InlineData("\\overline{AB} \\perp \\underline{CD}")]
    [InlineData("\\lim_{x \\to \\infty} f(x)")]
    [InlineData("\\max_{i} a_i + \\min_{j} b_j")]
    [InlineData("\\iint_D f \\, dA + \\iiint_V g \\, dV + \\oint_C h")]
    [InlineData("\\forall x \\in S \\; \\exists y \\notin S")]
    [InlineData("\\langle u, v \\rangle \\perp \\parallel")]
    [InlineData("\\Re z + \\Im z = \\hbar \\ell")]
    [InlineData("A \\Rightarrow B \\Leftarrow C \\Leftrightarrow D \\mapsto E")]
    [InlineData("\\lceil x \\rceil + \\lfloor y \\rfloor")]
    [InlineData("\\varepsilon \\vartheta \\varphi \\varpi \\varrho \\varsigma")]
    public void Deepened_constructs_emit_validator_clean_omml(string latex) =>
        AssertOmmlValidates(latex);

    [Fact]
    public void Cases_emits_an_equation_array_in_a_left_brace()
    {
        var math = Emit("\\begin{cases} a & x>0 \\\\ b & x \\leq 0 \\end{cases}");
        var delimiter = Assert.Single(math.Elements<M.Delimiter>());
        Assert.Equal("{", delimiter.DelimiterProperties?.GetFirstChild<M.BeginChar>()?.Val?.Value);
        Assert.Equal(".", delimiter.DelimiterProperties?.GetFirstChild<M.EndChar>()?.Val?.Value); // open-only brace
        var array = Assert.Single(delimiter.Descendants<M.EquationArray>());
        Assert.Equal(2, array.Elements<M.Base>().Count()); // two rows
    }

    [Fact]
    public void Aligned_emits_a_bare_equation_array_with_a_row_per_line()
    {
        var math = Emit("\\begin{aligned} a &= b \\\\ &= c \\\\ &= d \\end{aligned}");
        Assert.Empty(math.Elements<M.Delimiter>()); // aligned has no surrounding brace
        var array = Assert.Single(math.Elements<M.EquationArray>());
        Assert.Equal(3, array.Elements<M.Base>().Count());
    }

    [Fact]
    public void Binomial_emits_a_no_bar_fraction_in_parentheses()
    {
        var math = Emit("\\binom{n}{k}");
        var delimiter = Assert.Single(math.Elements<M.Delimiter>());
        Assert.Equal("(", delimiter.DelimiterProperties?.GetFirstChild<M.BeginChar>()?.Val?.Value);
        Assert.Equal(")", delimiter.DelimiterProperties?.GetFirstChild<M.EndChar>()?.Val?.Value);
        var fraction = Assert.Single(delimiter.Descendants<M.Fraction>());
        Assert.Equal(M.FractionTypeValues.NoBar, fraction.FractionProperties?.GetFirstChild<M.FractionType>()?.Val?.Value);
        Assert.Equal("n", fraction.Numerator!.InnerText);
        Assert.Equal("k", fraction.Denominator!.InnerText);
    }

    [Fact]
    public void Overbrace_emits_a_top_group_char_with_its_label_as_an_upper_limit()
    {
        var math = Emit("\\overbrace{a+b}^{n}");
        var limit = Assert.Single(math.Elements<M.LimitUpper>());
        var groupChar = Assert.Single(limit.Descendants<M.GroupChar>());
        Assert.Equal("⏞", groupChar.GroupCharProperties?.GetFirstChild<M.AccentChar>()?.Val?.Value);
        Assert.Equal(M.VerticalJustificationValues.Top, groupChar.GroupCharProperties?.GetFirstChild<M.Position>()?.Val?.Value);
        Assert.Equal("n", limit.Limit!.InnerText);
    }

    [Fact]
    public void Underbrace_emits_a_bottom_group_char_with_a_lower_limit_label()
    {
        var math = Emit("\\underbrace{x}_{k}");
        var limit = Assert.Single(math.Elements<M.LimitLower>());
        var groupChar = Assert.Single(limit.Descendants<M.GroupChar>());
        Assert.Equal("⏟", groupChar.GroupCharProperties?.GetFirstChild<M.AccentChar>()?.Val?.Value);
        Assert.Equal(M.VerticalJustificationValues.Bottom, groupChar.GroupCharProperties?.GetFirstChild<M.Position>()?.Val?.Value);
        Assert.Equal("k", limit.Limit!.InnerText);
    }

    [Fact]
    public void Bare_brace_without_label_emits_a_plain_group_char()
    {
        var math = Emit("\\overbrace{a+b}");
        Assert.Empty(math.Elements<M.LimitUpper>());
        Assert.Single(math.Elements<M.GroupChar>());
    }

    [Fact]
    public void Underline_emits_a_bottom_bar()
    {
        var bar = Assert.Single(Emit("\\underline{x+y}").Elements<M.Bar>());
        Assert.Equal(M.VerticalJustificationValues.Bottom, bar.BarProperties?.GetFirstChild<M.Position>()?.Val?.Value);
    }

    [Fact]
    public void Overline_stays_a_top_bar()
    {
        var bar = Assert.Single(Emit("\\overline{x}").Elements<M.Bar>());
        Assert.Equal(M.VerticalJustificationValues.Top, bar.BarProperties?.GetFirstChild<M.Position>()?.Val?.Value);
    }

    [Fact]
    public void Lim_with_subscript_emits_a_lower_limit_not_a_subscript()
    {
        var math = Emit("\\lim_{x \\to 0} f");
        var lim = Assert.Single(math.Elements<M.LimitLower>());
        Assert.Empty(math.Elements<M.Subscript>()); // the bound is a lower limit, not a subscript
        Assert.Equal("lim", lim.Base!.InnerText);
        Assert.Contains("x", lim.Limit!.InnerText, StringComparison.Ordinal);
    }

    [Fact]
    public void Lim_function_name_renders_upright()
    {
        var lim = Assert.Single(Emit("\\lim_{n} a").Elements<M.LimitLower>());
        var run = Assert.Single(lim.Base!.Descendants<M.Run>());
        Assert.NotNull(run.MathRunProperties?.GetFirstChild<M.NormalText>());
    }

    [Fact]
    public void Multi_integrals_carry_their_own_operator_glyphs()
    {
        Assert.Equal("∬", Assert.Single(Emit("\\iint_D f").Elements<M.Nary>())
            .NaryProperties?.GetFirstChild<M.AccentChar>()?.Val?.Value);
        Assert.Equal("∭", Assert.Single(Emit("\\iiint_V g").Elements<M.Nary>())
            .NaryProperties?.GetFirstChild<M.AccentChar>()?.Val?.Value);
        Assert.Equal("∮", Assert.Single(Emit("\\oint_C h").Elements<M.Nary>())
            .NaryProperties?.GetFirstChild<M.AccentChar>()?.Val?.Value);
    }

    [Fact]
    public void New_symbols_map_to_their_unicode_glyphs()
    {
        Assert.Contains("⊥", Emit("a \\perp b").InnerText);
        Assert.Contains("∥", Emit("a \\parallel b").InnerText);
        Assert.Contains("∀", Emit("\\forall x").InnerText);
        Assert.Contains("∃", Emit("\\exists y").InnerText);
        Assert.Contains("∈", Emit("x \\in S").InnerText);
        Assert.Contains("∉", Emit("x \\notin S").InnerText);
        Assert.Contains("⟨", Emit("\\langle u \\rangle").InnerText);
        Assert.Contains("⌈", Emit("\\lceil x \\rceil").InnerText);
        Assert.Contains("⌊", Emit("\\lfloor x \\rfloor").InnerText);
        Assert.Contains("ℜ", Emit("\\Re z").InnerText);
        Assert.Contains("ℏ", Emit("\\hbar").InnerText);
        Assert.Contains("ℓ", Emit("\\ell").InnerText);
        Assert.Contains("ε", Emit("\\varepsilon").InnerText);
        Assert.Contains("ϑ", Emit("\\vartheta").InnerText);
        Assert.Contains("⟹", Emit("a \\Longrightarrow b").InnerText);
    }

    [Fact]
    public void Cases_shorthand_brace_form_parses_like_the_environment()
    {
        var math = Emit("\\cases{ a & b \\\\ c & d }");
        var delimiter = Assert.Single(math.Elements<M.Delimiter>());
        Assert.Equal("{", delimiter.DelimiterProperties?.GetFirstChild<M.BeginChar>()?.Val?.Value);
        Assert.Equal(2, Assert.Single(delimiter.Descendants<M.EquationArray>()).Elements<M.Base>().Count());
    }

    [Fact]
    public void Deepened_constructs_report_no_unknown_tokens()
    {
        var parsed = LatexParser.Parse(
            "\\begin{cases} \\binom{n}{k} \\\\ \\lim_{x\\to0} \\overbrace{a}^{b} \\end{cases}");
        Assert.Empty(parsed.UnknownTokens);
    }

    [Fact]
    public void Unknown_command_still_degrades_to_literal_and_warns()
    {
        // The honest degrade path must survive the deepening: a still-unknown command
        // appears literally and is reported.
        var parsed = LatexParser.Parse("\\frobnicate{x} + \\binom{n}{k}");
        Assert.Contains("\\frobnicate", parsed.UnknownTokens);
        Assert.DoesNotContain("\\binom", parsed.UnknownTokens);
        Assert.Contains("\\frobnicate", OmmlEmitter.ToOfficeMath(parsed.Root).InnerText);
    }
}
