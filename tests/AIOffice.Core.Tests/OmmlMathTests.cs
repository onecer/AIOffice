using System.Xml.Linq;
using AIOffice.Core.Equations;
using Xunit;

namespace AIOffice.Core.Tests;

/// <summary>
/// Structure tests for the shared pure-XML OMML producer (<see cref="OmmlMath"/>),
/// the exact converter the Pptx handler consumes (the Word handler uses the
/// DocumentFormat.OpenXml emitter, which mirrors this one). Core carries no OOXML
/// dependency, so these assert the emitted <c>m:oMath</c> element tree directly;
/// the validator-clean oracle for the same constructs lives in the Word and Pptx
/// test suites where the SDK is referenced.
/// </summary>
public sealed class OmmlMathTests
{
    private static readonly XNamespace M = OmmlMath.M;

    private static XElement Emit(string latex) => OmmlMath.FromLatex(latex).OMath;

    private static IReadOnlyList<string> Unknown(string latex) => OmmlMath.FromLatex(latex).UnknownTokens;

    /// <summary>All descendants (and self) with the given local name in the math namespace.</summary>
    private static IEnumerable<XElement> All(XElement root, string local) =>
        root.DescendantsAndSelf(M + local);

    [Fact]
    public void Every_emitted_tree_is_rooted_in_o_math()
    {
        Assert.Equal(M + "oMath", Emit("x").Name);
    }

    // ----------------------------------------------------------- environments

    [Fact]
    public void Cases_wraps_an_equation_array_in_an_open_only_left_brace()
    {
        var math = Emit("\\begin{cases} a & x>0 \\\\ b & x \\leq 0 \\end{cases}");
        var delimiter = Assert.Single(All(math, "d"));
        var props = delimiter.Element(M + "dPr")!;
        Assert.Equal("{", props.Element(M + "begChr")!.Attribute(M + "val")!.Value);
        Assert.Equal(".", props.Element(M + "endChr")!.Attribute(M + "val")!.Value);

        var array = Assert.Single(All(math, "eqArr"));
        Assert.Equal(2, array.Elements(M + "e").Count()); // one m:e per row
    }

    [Fact]
    public void Aligned_emits_a_bare_equation_array_with_a_row_per_line()
    {
        var math = Emit("\\begin{aligned} a &= b \\\\ &= c \\\\ &= d \\end{aligned}");
        Assert.Empty(All(math, "d")); // no surrounding delimiter
        var array = Assert.Single(All(math, "eqArr"));
        Assert.Equal(3, array.Elements(M + "e").Count());
        Assert.Equal("center", array.Element(M + "eqArrPr")!.Element(M + "baseJc")!.Attribute(M + "val")!.Value);
    }

    [Fact]
    public void Align_and_alignedat_are_recognized_environments()
    {
        Assert.Single(All(Emit("\\begin{align} x &= 1 \\\\ y &= 2 \\end{align}"), "eqArr"));
        // alignedat's mandatory {n} column-pair count is consumed, not emitted as text.
        var math = Emit("\\begin{alignedat}{2} a &= b & c &= d \\end{alignedat}");
        Assert.Single(All(math, "eqArr"));
        Assert.DoesNotContain("2", Assert.Single(All(math, "eqArr")).Value);
    }

    [Fact]
    public void Cases_shorthand_brace_form_parses_like_the_environment()
    {
        var math = Emit("\\cases{ a & b \\\\ c & d }");
        var delimiter = Assert.Single(All(math, "d"));
        Assert.Equal("{", delimiter.Element(M + "dPr")!.Element(M + "begChr")!.Attribute(M + "val")!.Value);
        Assert.Equal(2, Assert.Single(All(math, "eqArr")).Elements(M + "e").Count());
    }

    // -------------------------------------------------------------- binomial

    [Fact]
    public void Binom_emits_a_no_bar_fraction_inside_parentheses()
    {
        var math = Emit("\\binom{n}{k}");
        var delimiter = Assert.Single(All(math, "d"));
        var props = delimiter.Element(M + "dPr")!;
        Assert.Equal("(", props.Element(M + "begChr")!.Attribute(M + "val")!.Value);
        Assert.Equal(")", props.Element(M + "endChr")!.Attribute(M + "val")!.Value);

        var fraction = Assert.Single(All(math, "f"));
        Assert.Equal("noBar", fraction.Element(M + "fPr")!.Element(M + "type")!.Attribute(M + "val")!.Value);
        Assert.Equal("n", fraction.Element(M + "num")!.Value);
        Assert.Equal("k", fraction.Element(M + "den")!.Value);
    }

    [Theory]
    [InlineData("dbinom")]
    [InlineData("tbinom")]
    public void Binom_aliases_emit_the_same_no_bar_fraction(string command)
    {
        var math = Emit($"\\{command}{{a}}{{b}}");
        Assert.Equal("noBar", Assert.Single(All(math, "f")).Element(M + "fPr")!.Element(M + "type")!.Attribute(M + "val")!.Value);
    }

    // ----------------------------------------------------------------- braces

    [Fact]
    public void Overbrace_emits_a_top_group_char_with_an_upper_limit_label()
    {
        var math = Emit("\\overbrace{a+b}^{n}");
        var limit = Assert.Single(All(math, "limUpp"));
        var group = Assert.Single(All(math, "groupChr"));
        var groupProps = group.Element(M + "groupChrPr")!;
        Assert.Equal("⏞", groupProps.Element(M + "chr")!.Attribute(M + "val")!.Value);
        Assert.Equal("top", groupProps.Element(M + "pos")!.Attribute(M + "val")!.Value);
        Assert.Equal("n", limit.Element(M + "lim")!.Value);
    }

    [Fact]
    public void Underbrace_emits_a_bottom_group_char_with_a_lower_limit_label()
    {
        var math = Emit("\\underbrace{x_1+x_2}_{\\text{sum}}");
        var limit = Assert.Single(All(math, "limLow"));
        var group = Assert.Single(All(math, "groupChr"));
        Assert.Equal("⏟", group.Element(M + "groupChrPr")!.Element(M + "chr")!.Attribute(M + "val")!.Value);
        Assert.Equal("bot", group.Element(M + "groupChrPr")!.Element(M + "pos")!.Attribute(M + "val")!.Value);
        Assert.Equal("sum", limit.Element(M + "lim")!.Value);
    }

    [Fact]
    public void Brace_without_a_label_is_a_plain_group_char()
    {
        var math = Emit("\\overbrace{a}");
        Assert.Empty(All(math, "limUpp"));
        Assert.Single(All(math, "groupChr"));
    }

    // ------------------------------------------------------------ bars/limits

    [Fact]
    public void Overline_is_a_top_bar_and_underline_is_a_bottom_bar()
    {
        Assert.Equal("top", Assert.Single(All(Emit("\\overline{x}"), "bar"))
            .Element(M + "barPr")!.Element(M + "pos")!.Attribute(M + "val")!.Value);
        Assert.Equal("bot", Assert.Single(All(Emit("\\underline{x}"), "bar"))
            .Element(M + "barPr")!.Element(M + "pos")!.Attribute(M + "val")!.Value);
    }

    [Fact]
    public void Lim_with_subscript_becomes_a_lower_limit_not_a_subscript()
    {
        var math = Emit("\\lim_{x \\to 0} f(x)");
        var lim = Assert.Single(All(math, "limLow"));
        Assert.Empty(All(math, "sSub")); // the bound is a lower limit, not a subscript
        Assert.Equal("lim", lim.Element(M + "e")!.Value);
        Assert.Contains("x", lim.Element(M + "lim")!.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void Lim_operator_name_is_an_upright_run()
    {
        var lim = Assert.Single(All(Emit("\\max_{i} a_i"), "limLow"));
        var run = Assert.Single(lim.Element(M + "e")!.Elements(M + "r"));
        Assert.NotNull(run.Element(M + "rPr")!.Element(M + "nor"));
    }

    // ---------------------------------------------------------- multi-integral

    [Theory]
    [InlineData("iint", "∬")]
    [InlineData("iiint", "∭")]
    [InlineData("oint", "∮")]
    public void Multi_integrals_emit_their_own_nary_glyph(string command, string glyph)
    {
        var nary = Assert.Single(All(Emit($"\\{command}_D f"), "nary"));
        Assert.Equal(glyph, nary.Element(M + "naryPr")!.Element(M + "chr")!.Attribute(M + "val")!.Value);
    }

    // -------------------------------------------------------------- symbols

    [Theory]
    [InlineData("\\perp", "⊥")]
    [InlineData("\\parallel", "∥")]
    [InlineData("\\forall", "∀")]
    [InlineData("\\exists", "∃")]
    [InlineData("\\nexists", "∄")]
    [InlineData("\\in", "∈")]
    [InlineData("\\notin", "∉")]
    [InlineData("\\ni", "∋")]
    [InlineData("\\subseteq", "⊆")]
    [InlineData("\\cup", "∪")]
    [InlineData("\\cap", "∩")]
    [InlineData("\\emptyset", "∅")]
    [InlineData("\\otimes", "⊗")]
    [InlineData("\\oplus", "⊕")]
    [InlineData("\\equiv", "≡")]
    [InlineData("\\simeq", "≃")]
    [InlineData("\\cong", "≅")]
    [InlineData("\\propto", "∝")]
    [InlineData("\\Re", "ℜ")]
    [InlineData("\\Im", "ℑ")]
    [InlineData("\\hbar", "ℏ")]
    [InlineData("\\ell", "ℓ")]
    [InlineData("\\Rightarrow", "⇒")]
    [InlineData("\\Leftrightarrow", "⇔")]
    [InlineData("\\mapsto", "↦")]
    [InlineData("\\langle", "⟨")]
    [InlineData("\\rangle", "⟩")]
    [InlineData("\\lceil", "⌈")]
    [InlineData("\\rfloor", "⌋")]
    [InlineData("\\varepsilon", "ε")]
    [InlineData("\\vartheta", "ϑ")]
    [InlineData("\\varphi", "φ")]
    [InlineData("\\varpi", "ϖ")]
    [InlineData("\\varrho", "ϱ")]
    [InlineData("\\varsigma", "ς")]
    [InlineData("\\Longrightarrow", "⟹")]
    [InlineData("\\iff", "⟺")]
    public void Symbol_maps_to_its_unicode_glyph(string latex, string glyph)
    {
        Assert.Contains(glyph, Emit(latex).Value);
    }

    // ----------------------------------------------------------- honest degrade

    [Fact]
    public void Deepened_constructs_report_no_unknown_tokens()
    {
        Assert.Empty(Unknown("\\begin{cases} \\binom{n}{k} \\\\ \\lim_{x\\to0} \\overbrace{a}^{b} \\end{cases}"));
        Assert.Empty(Unknown("\\underline{x} + \\overline{y} + \\iint_D f"));
    }

    [Fact]
    public void Still_unknown_command_degrades_to_a_literal_run_and_is_reported()
    {
        var result = OmmlMath.FromLatex("\\frobnicate{x} + \\binom{n}{k}");
        Assert.Contains("\\frobnicate", result.UnknownTokens);
        Assert.DoesNotContain("\\binom", result.UnknownTokens);
        Assert.Contains("\\frobnicate", result.OMath.Value);
    }
}
