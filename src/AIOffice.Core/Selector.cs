using System.Globalization;
using System.Text;

namespace AIOffice.Core;

/// <summary>Comparison operators usable inside <c>[attr OP value]</c> predicates.</summary>
public enum SelectorOperator
{
    Equals,         // =
    NotEquals,      // !=
    GreaterThan,    // >
    GreaterOrEqual, // >=
    LessThan,       // <
    LessOrEqual,    // <=
    ContainsText,   // *=
}

/// <summary>Base of the small selector predicate AST.</summary>
public abstract record SelectorPredicate;

/// <summary>An attribute comparison: <c>[style=Heading1]</c>, <c>[value&gt;100]</c>.</summary>
public sealed record AttributePredicate(string Attribute, SelectorOperator Op, string Value) : SelectorPredicate
{
    /// <summary>The value parsed as an invariant double when it is numeric (for &gt;/&lt; comparisons).</summary>
    public double? NumericValue =>
        double.TryParse(Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : null;
}

/// <summary>A text containment pseudo-class: <c>:contains('Q3')</c>.</summary>
public sealed record ContainsPredicate(string Text) : SelectorPredicate;

/// <summary>
/// A parsed CSS-like selector: an element name (or <c>*</c>) plus zero or more
/// predicates. Examples: <c>p[style=Heading1]</c>, <c>cell[value&gt;100]</c>,
/// <c>shape:contains('Q3')</c>. Evaluation is performed per format handler;
/// this type owns only the grammar.
/// </summary>
public sealed record Selector
{
    private const string GrammarHint =
        "Selectors look like p[style=Heading1], cell[value>100] or shape:contains('Q3'); " +
        "operators are = != > >= < <= *=. Run 'aioffice help selectors' for details.";

    /// <summary>Element name to match, or <c>*</c> for any element.</summary>
    public required string Element { get; init; }

    public required IReadOnlyList<SelectorPredicate> Predicates { get; init; }

    public string ToCanonicalString()
    {
        var sb = new StringBuilder(Element);
        foreach (var predicate in Predicates)
        {
            switch (predicate)
            {
                case AttributePredicate a:
                    sb.Append('[').Append(a.Attribute).Append(OperatorToken(a.Op)).Append(a.Value).Append(']');
                    break;
                case ContainsPredicate c:
                    sb.Append(":contains('").Append(c.Text.Replace("'", "''", StringComparison.Ordinal)).Append("')");
                    break;
                default:
                    throw new InvalidOperationException($"Unknown predicate: {predicate.GetType().Name}");
            }
        }

        return sb.ToString();
    }

    public override string ToString() => ToCanonicalString();

    private static string OperatorToken(SelectorOperator op) => op switch
    {
        SelectorOperator.Equals => "=",
        SelectorOperator.NotEquals => "!=",
        SelectorOperator.GreaterThan => ">",
        SelectorOperator.GreaterOrEqual => ">=",
        SelectorOperator.LessThan => "<",
        SelectorOperator.LessOrEqual => "<=",
        SelectorOperator.ContainsText => "*=",
        _ => throw new InvalidOperationException($"Unknown operator: {op}"),
    };

    /// <summary>Parses a selector or throws <c>invalid_args</c> with a grammar hint.</summary>
    public static Selector Parse(string text)
    {
        if (!TryParse(text, out var selector, out var error))
        {
            throw new AiofficeException(ErrorCodes.InvalidArgs, error!, GrammarHint);
        }

        return selector!;
    }

    public static bool TryParse(string? text, out Selector? selector, out string? error)
    {
        selector = null;
        error = null;

        if (string.IsNullOrWhiteSpace(text))
        {
            error = "Selector is empty.";
            return false;
        }

        text = text.Trim();
        var i = 0;

        var element = ReadElementName(text, ref i);
        if (element.Length == 0)
        {
            error = $"Selector must start with an element name or '*': {text}";
            return false;
        }

        var predicates = new List<SelectorPredicate>();
        while (i < text.Length)
        {
            switch (text[i])
            {
                case '[':
                {
                    var predicate = ParseAttribute(text, ref i, ref error);
                    if (predicate is null)
                    {
                        return false;
                    }

                    predicates.Add(predicate);
                    break;
                }

                case ':':
                {
                    var predicate = ParseContains(text, ref i, ref error);
                    if (predicate is null)
                    {
                        return false;
                    }

                    predicates.Add(predicate);
                    break;
                }

                default:
                    error = $"Unexpected character '{text[i]}' at position {i} in selector: {text}";
                    return false;
            }
        }

        selector = new Selector { Element = element, Predicates = predicates };
        return true;
    }

    private static string ReadElementName(string text, ref int i)
    {
        if (text[i] == '*')
        {
            i++;
            return "*";
        }

        var start = i;
        while (i < text.Length && (char.IsAsciiLetterOrDigit(text[i]) || text[i] is '_' or '-'))
        {
            i++;
        }

        return text[start..i];
    }

    private static AttributePredicate? ParseAttribute(string text, ref int i, ref string? error)
    {
        i++; // consume '['
        var nameStart = i;
        while (i < text.Length && (char.IsAsciiLetterOrDigit(text[i]) || text[i] is '_' or '-' or '.'))
        {
            i++;
        }

        var attribute = text[nameStart..i];
        if (attribute.Length == 0)
        {
            error = $"Attribute name missing at position {i} in selector: {text}";
            return null;
        }

        var op = ReadOperator(text, ref i);
        if (op is null)
        {
            error = $"Expected an operator (= != > >= < <= *=) after '{attribute}' in selector: {text}";
            return null;
        }

        string? value;
        if (i < text.Length && text[i] is '\'' or '"')
        {
            value = ReadQuoted(text, ref i, ref error);
            if (value is null)
            {
                return null;
            }
        }
        else
        {
            var valueStart = i;
            while (i < text.Length && text[i] != ']')
            {
                i++;
            }

            value = text[valueStart..i].Trim();
        }

        if (i >= text.Length || text[i] != ']')
        {
            error = $"Unterminated attribute predicate (missing ']') in selector: {text}";
            return null;
        }

        i++; // consume ']'
        if (value.Length == 0)
        {
            error = $"Attribute '{attribute}' has an empty comparison value in selector: {text}";
            return null;
        }

        return new AttributePredicate(attribute, op.Value, value);
    }

    private static SelectorOperator? ReadOperator(string text, ref int i)
    {
        if (i >= text.Length)
        {
            return null;
        }

        switch (text[i])
        {
            case '=':
                i++;
                return SelectorOperator.Equals;
            case '!' when Peek(text, i + 1) == '=':
                i += 2;
                return SelectorOperator.NotEquals;
            case '>' when Peek(text, i + 1) == '=':
                i += 2;
                return SelectorOperator.GreaterOrEqual;
            case '>':
                i++;
                return SelectorOperator.GreaterThan;
            case '<' when Peek(text, i + 1) == '=':
                i += 2;
                return SelectorOperator.LessOrEqual;
            case '<':
                i++;
                return SelectorOperator.LessThan;
            case '*' when Peek(text, i + 1) == '=':
                i += 2;
                return SelectorOperator.ContainsText;
            default:
                return null;
        }
    }

    private static ContainsPredicate? ParseContains(string text, ref int i, ref string? error)
    {
        const string prefix = ":contains(";
        if (!text.AsSpan(i).StartsWith(prefix, StringComparison.Ordinal))
        {
            error = $"Unknown pseudo-class at position {i}; only :contains('text') is supported in: {text}";
            return null;
        }

        i += prefix.Length;
        if (i >= text.Length || text[i] is not ('\'' or '"'))
        {
            error = $":contains() requires a quoted string in selector: {text}";
            return null;
        }

        var value = ReadQuoted(text, ref i, ref error);
        if (value is null)
        {
            return null;
        }

        if (i >= text.Length || text[i] != ')')
        {
            error = $"Unterminated :contains( in selector: {text}";
            return null;
        }

        i++; // consume ')'
        return new ContainsPredicate(value);
    }

    /// <summary>Reads a quoted string ('…' or "…"); a doubled quote escapes itself.</summary>
    private static string? ReadQuoted(string text, ref int i, ref string? error)
    {
        var quote = text[i];
        i++; // consume opening quote
        var sb = new StringBuilder();
        while (i < text.Length)
        {
            if (text[i] == quote)
            {
                if (Peek(text, i + 1) == quote)
                {
                    sb.Append(quote);
                    i += 2;
                    continue;
                }

                i++; // consume closing quote
                return sb.ToString();
            }

            sb.Append(text[i]);
            i++;
        }

        error = $"Unterminated quoted string in selector: {text}";
        return null;
    }

    private static char Peek(string text, int index) => index < text.Length ? text[index] : '\0';
}
