using System.Text;

namespace AIOffice.Word.Equations;

/// <summary>The kinds of token the LaTeX subset lexer produces.</summary>
internal enum LatexTokenKind
{
    /// <summary>A run of ordinary characters (letters, digits, punctuation) — math text.</summary>
    Text,

    /// <summary>A control sequence: a backslash command like <c>\frac</c>, <c>\alpha</c>, <c>\left</c>.</summary>
    Command,

    /// <summary><c>{</c> — opens a group.</summary>
    OpenBrace,

    /// <summary><c>}</c> — closes a group.</summary>
    CloseBrace,

    /// <summary><c>[</c> — opens an optional argument (e.g. the root degree).</summary>
    OpenBracket,

    /// <summary><c>]</c> — closes an optional argument.</summary>
    CloseBracket,

    /// <summary><c>^</c> — superscript marker.</summary>
    Caret,

    /// <summary><c>_</c> — subscript marker.</summary>
    Underscore,

    /// <summary><c>&amp;</c> — matrix column separator.</summary>
    Ampersand,

    /// <summary><c>\\</c> — matrix/array row separator.</summary>
    RowBreak,

    /// <summary>End of input.</summary>
    End,
}

/// <summary>One lexical token with its source text (for round-trippable error reporting).</summary>
internal readonly record struct LatexToken(LatexTokenKind Kind, string Value)
{
    public override string ToString() => Kind switch
    {
        LatexTokenKind.Command => "\\" + Value,
        LatexTokenKind.End => "<end>",
        _ => Value,
    };
}

/// <summary>
/// A hand-rolled lexer for the supported LaTeX-math subset. It segments the
/// source into commands (<c>\name</c> or a single non-letter escape like
/// <c>\{</c>), grouping/optional braces, sub/superscript markers, matrix
/// separators, and runs of ordinary text. Whitespace is insignificant in math
/// mode (it never produces a token) except that it terminates a command name,
/// exactly as TeX treats it.
/// </summary>
internal sealed class LatexLexer
{
    private readonly string _src;
    private int _pos;

    public LatexLexer(string source) => _src = source;

    public LatexToken Next()
    {
        SkipInsignificantWhitespace();
        if (_pos >= _src.Length)
        {
            return new LatexToken(LatexTokenKind.End, string.Empty);
        }

        var c = _src[_pos];
        switch (c)
        {
            case '\\':
                return LexCommand();

            case '{':
                _pos++;
                return new LatexToken(LatexTokenKind.OpenBrace, "{");

            case '}':
                _pos++;
                return new LatexToken(LatexTokenKind.CloseBrace, "}");

            case '[':
                _pos++;
                return new LatexToken(LatexTokenKind.OpenBracket, "[");

            case ']':
                _pos++;
                return new LatexToken(LatexTokenKind.CloseBracket, "]");

            case '^':
                _pos++;
                return new LatexToken(LatexTokenKind.Caret, "^");

            case '_':
                _pos++;
                return new LatexToken(LatexTokenKind.Underscore, "_");

            case '&':
                _pos++;
                return new LatexToken(LatexTokenKind.Ampersand, "&");

            default:
                return LexText();
        }
    }

    /// <summary>
    /// A control sequence. <c>\\</c> is the row-break token. A backslash before a
    /// letter consumes the maximal run of ASCII letters as the command name; a
    /// backslash before any other character is a single-symbol escape command
    /// (e.g. <c>\{</c>, <c>\,</c>, <c>\%</c>).
    /// </summary>
    private LatexToken LexCommand()
    {
        _pos++; // consume backslash
        if (_pos >= _src.Length)
        {
            return new LatexToken(LatexTokenKind.Command, string.Empty);
        }

        var first = _src[_pos];
        if (first == '\\')
        {
            _pos++;
            return new LatexToken(LatexTokenKind.RowBreak, "\\\\");
        }

        if (!char.IsAsciiLetter(first))
        {
            _pos++;
            return new LatexToken(LatexTokenKind.Command, first.ToString());
        }

        var start = _pos;
        while (_pos < _src.Length && char.IsAsciiLetter(_src[_pos]))
        {
            _pos++;
        }

        return new LatexToken(LatexTokenKind.Command, _src[start.._pos]);
    }

    /// <summary>
    /// A run of ordinary characters up to the next special character. Each text
    /// token is a single source character so the parser can attach sub/sup to the
    /// exact preceding atom (TeX scripts bind to one character), while still
    /// letting the parser coalesce adjacent text into one math run on emit.
    /// </summary>
    private LatexToken LexText()
    {
        var c = _src[_pos];
        _pos++;
        return new LatexToken(LatexTokenKind.Text, c.ToString());
    }

    private void SkipInsignificantWhitespace()
    {
        while (_pos < _src.Length && _src[_pos] is ' ' or '\t' or '\r' or '\n')
        {
            _pos++;
        }
    }

    /// <summary>Tokenizes the whole source eagerly (the parser walks the list with lookahead).</summary>
    public static List<LatexToken> Tokenize(string source)
    {
        var lexer = new LatexLexer(source);
        var tokens = new List<LatexToken>();
        LatexToken token;
        do
        {
            token = lexer.Next();
            tokens.Add(token);
        }
        while (token.Kind != LatexTokenKind.End);

        return tokens;
    }
}
