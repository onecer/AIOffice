using AIOffice.Core;
using Xunit;

namespace AIOffice.Core.Tests;

public class SelectorTests
{
    [Fact]
    public void Parses_element_with_style_attribute()
    {
        var selector = Selector.Parse("p[style=Heading1]");

        Assert.Equal("p", selector.Element);
        var predicate = Assert.IsType<AttributePredicate>(Assert.Single(selector.Predicates));
        Assert.Equal("style", predicate.Attribute);
        Assert.Equal(SelectorOperator.Equals, predicate.Op);
        Assert.Equal("Heading1", predicate.Value);
    }

    [Fact]
    public void Parses_numeric_greater_than()
    {
        var selector = Selector.Parse("cell[value>100]");

        var predicate = Assert.IsType<AttributePredicate>(Assert.Single(selector.Predicates));
        Assert.Equal(SelectorOperator.GreaterThan, predicate.Op);
        Assert.Equal("100", predicate.Value);
        Assert.Equal(100d, predicate.NumericValue);
    }

    [Fact]
    public void Parses_contains_pseudo_class()
    {
        var selector = Selector.Parse("shape:contains('Q3')");

        Assert.Equal("shape", selector.Element);
        var predicate = Assert.IsType<ContainsPredicate>(Assert.Single(selector.Predicates));
        Assert.Equal("Q3", predicate.Text);
    }

    [Theory]
    [InlineData("cell[value>=100]", SelectorOperator.GreaterOrEqual)]
    [InlineData("cell[value<=100]", SelectorOperator.LessOrEqual)]
    [InlineData("cell[value<100]", SelectorOperator.LessThan)]
    [InlineData("cell[value!=100]", SelectorOperator.NotEquals)]
    [InlineData("cell[value*=100]", SelectorOperator.ContainsText)]
    public void Parses_all_comparison_operators(string text, SelectorOperator expected)
    {
        var selector = Selector.Parse(text);

        var predicate = Assert.IsType<AttributePredicate>(Assert.Single(selector.Predicates));
        Assert.Equal(expected, predicate.Op);
    }

    [Fact]
    public void Parses_compound_predicates_in_order()
    {
        var selector = Selector.Parse("p[style=Heading1][align=center]:contains('Intro')");

        Assert.Equal(3, selector.Predicates.Count);
        Assert.IsType<AttributePredicate>(selector.Predicates[0]);
        Assert.IsType<AttributePredicate>(selector.Predicates[1]);
        var contains = Assert.IsType<ContainsPredicate>(selector.Predicates[2]);
        Assert.Equal("Intro", contains.Text);
    }

    [Fact]
    public void Parses_universal_element()
    {
        var selector = Selector.Parse("*:contains('total')");

        Assert.Equal("*", selector.Element);
    }

    [Fact]
    public void Quoted_attribute_values_support_spaces()
    {
        var selector = Selector.Parse("p[style='Heading 1']");

        var predicate = Assert.IsType<AttributePredicate>(Assert.Single(selector.Predicates));
        Assert.Equal("Heading 1", predicate.Value);
    }

    [Fact]
    public void Contains_supports_escaped_quote()
    {
        var selector = Selector.Parse("p:contains('it''s')");

        var predicate = Assert.IsType<ContainsPredicate>(Assert.Single(selector.Predicates));
        Assert.Equal("it's", predicate.Text);
    }

    [Fact]
    public void Canonical_string_round_trips()
    {
        var text = "cell[value>100]:contains('Q3')";
        var selector = Selector.Parse(text);
        var reparsed = Selector.Parse(selector.ToCanonicalString());

        Assert.Equal(text, selector.ToCanonicalString());
        Assert.Equal(selector.Element, reparsed.Element);
        Assert.Equal(selector.Predicates, reparsed.Predicates); // sequence equality of record predicates
    }

    [Theory]
    [InlineData("")]                       // empty
    [InlineData("[style=x]")]              // missing element
    [InlineData("p[style=]")]              // empty value
    [InlineData("p[style~Heading1]")]      // unknown operator
    [InlineData("p[style=Heading1")]       // unterminated bracket
    [InlineData("p:hover")]                // unsupported pseudo-class
    [InlineData("p:contains('open)")]      // unterminated quote
    [InlineData("p[style=x] p[style=y]")]  // descendant combinators not in M0
    public void Invalid_selectors_are_rejected(string text)
    {
        Assert.False(Selector.TryParse(text, out _, out var error));
        Assert.False(string.IsNullOrWhiteSpace(error));
    }

    [Fact]
    public void Parse_throws_invalid_args_with_suggestion()
    {
        var ex = Assert.Throws<AiofficeException>(() => Selector.Parse("p["));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("selector", ex.Suggestion, StringComparison.OrdinalIgnoreCase);
    }
}
