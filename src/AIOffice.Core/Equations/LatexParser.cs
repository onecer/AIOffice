namespace AIOffice.Core.Equations;

/// <summary>The parse outcome: the math tree plus any unrecognized command tokens (for the equation_partial warning).</summary>
public sealed record LatexParseResult(MathNode Root, IReadOnlyList<string> UnknownTokens);

/// <summary>
/// A recursive-descent parser for the supported LaTeX-math subset. It never
/// throws on unknown input: an unrecognized command becomes a literal math run
/// of its own token text and is recorded in <see cref="LatexParseResult.UnknownTokens"/>
/// so the caller can attach an <c>equation_partial</c> meta warning. Structural
/// errors that would otherwise corrupt the tree (an unmatched <c>}</c>, a
/// <c>\frac</c> missing an argument) degrade to empty arguments rather than
/// failing the whole equation.
/// </summary>
public sealed class LatexParser
{
    private readonly List<LatexToken> _tokens;
    private int _pos;
    private readonly List<string> _unknown = [];

    private LatexParser(List<LatexToken> tokens) => _tokens = tokens;

    public static LatexParseResult Parse(string latex)
    {
        var parser = new LatexParser(LatexLexer.Tokenize(latex));
        var root = parser.ParseSequence(stopAtRowOrAmp: false);
        return new LatexParseResult(root, parser._unknown);
    }

    private LatexToken Peek => _tokens[_pos];

    private LatexToken Advance() => _tokens[_pos++];

    private bool AtEnd => Peek.Kind == LatexTokenKind.End;

    /// <summary>
    /// Parses a run of atoms until a closing brace/bracket or end. <c>^</c> and
    /// <c>_</c> are post-fix operators that bind to the atom just emitted, so the
    /// sequence loop folds them onto the preceding node.
    /// </summary>
    private MathList ParseSequence(bool stopAtRowOrAmp)
    {
        var items = new List<MathNode>();
        while (!AtEnd)
        {
            var kind = Peek.Kind;
            if (kind is LatexTokenKind.CloseBrace or LatexTokenKind.CloseBracket)
            {
                break;
            }

            if (stopAtRowOrAmp && kind is LatexTokenKind.Ampersand or LatexTokenKind.RowBreak)
            {
                break;
            }

            if (kind is LatexTokenKind.Caret or LatexTokenKind.Underscore)
            {
                var baseNode = items.Count > 0 ? items[^1] : new MathText(string.Empty);
                if (items.Count > 0)
                {
                    items.RemoveAt(items.Count - 1);
                }

                items.Add(ParseScripts(baseNode));
                continue;
            }

            items.Add(ParseAtom());
        }

        return new MathList(items);
    }

    /// <summary>Folds any chain of <c>_</c>/<c>^</c> after a base into one script node.</summary>
    private MathNode ParseScripts(MathNode baseNode)
    {
        MathNode? sub = null;
        MathNode? sup = null;

        while (Peek.Kind is LatexTokenKind.Caret or LatexTokenKind.Underscore)
        {
            var isSup = Advance().Kind == LatexTokenKind.Caret;
            var arg = ParseArgument();
            if (isSup)
            {
                sup = sup is null ? arg : new MathList([sup, arg]);
            }
            else
            {
                sub = sub is null ? arg : new MathList([sub, arg]);
            }
        }

        return new MathScript(baseNode, sub, sup);
    }

    /// <summary>One atom: a command construct, a brace group, or a single text character.</summary>
    private MathNode ParseAtom()
    {
        var token = Peek;
        switch (token.Kind)
        {
            case LatexTokenKind.OpenBrace:
                return ParseBraceGroup();

            case LatexTokenKind.Command:
                return ParseCommand();

            case LatexTokenKind.Text:
                Advance();
                return new MathText(token.Value);

            case LatexTokenKind.OpenBracket:
                // A bare '[' outside an optional-argument context is ordinary text.
                Advance();
                return new MathText("[");

            case LatexTokenKind.CloseBracket:
                Advance();
                return new MathText("]");

            case LatexTokenKind.Ampersand:
            case LatexTokenKind.RowBreak:
                // Stray separators outside a matrix render as their literal glyph.
                Advance();
                return new MathText(token.Value);

            default:
                Advance();
                return new MathText(string.Empty);
        }
    }

    /// <summary>The body inside <c>{ … }</c>, returned as a list (possibly empty).</summary>
    private MathNode ParseBraceGroup()
    {
        Advance(); // {
        var body = ParseSequence(stopAtRowOrAmp: false);
        Expect(LatexTokenKind.CloseBrace);
        return body;
    }

    /// <summary>
    /// One required argument: a brace group, or—TeX style—the single next atom
    /// when the argument is unbraced (e.g. <c>x^2</c>, <c>\sqrt 3</c>).
    /// </summary>
    private MathNode ParseArgument()
    {
        if (Peek.Kind == LatexTokenKind.OpenBrace)
        {
            return ParseBraceGroup();
        }

        if (Peek.Kind is LatexTokenKind.Command or LatexTokenKind.Text)
        {
            return ParseAtom();
        }

        // No argument present (e.g. trailing ^ at end): empty.
        return new MathList([]);
    }

    /// <summary>An optional <c>[ … ]</c> argument (root degree), or null when absent.</summary>
    private MathNode? ParseOptionalArgument()
    {
        if (Peek.Kind != LatexTokenKind.OpenBracket)
        {
            return null;
        }

        Advance(); // [
        var body = ParseSequence(stopAtRowOrAmp: false);
        Expect(LatexTokenKind.CloseBracket);
        return body;
    }

    private MathNode ParseCommand()
    {
        var name = Advance().Value;

        switch (name)
        {
            case "frac" or "dfrac" or "tfrac" or "cfrac":
            {
                var num = ParseArgument();
                var den = ParseArgument();
                return new MathFraction(num, den);
            }

            case "sqrt":
            {
                var degree = ParseOptionalArgument();
                var radicand = ParseArgument();
                return new MathRadical(degree, radicand);
            }

            case "text" or "mathrm" or "mathbf" or "mathit" or "operatorname":
                return new MathLiteral(ReadRawGroupText());

            case "left":
                return ParseLeftRight();

            case "begin":
                return ParseEnvironment();

            case "bar" or "overline":
                return new MathBar(ParseArgument());

            case "underline":
                return new MathBar(ParseArgument(), Below: true);

            case "binom" or "dbinom" or "tbinom" or "choose":
            {
                var top = ParseArgument();
                var bottom = ParseArgument();
                return new MathBinomial(top, bottom);
            }

            case "overbrace" or "underbrace":
                return ParseBrace(below: name == "underbrace");

            case "cases":
                // \cases{...} shorthand: braces hold the rows the same way \begin{cases} does.
                return ParseCasesShorthand();

            case "frac{}{}": // never produced; defensive
                return new MathList([]);
        }

        // \lim (and limsup/liminf/max/min/…) with a following _subscript become a
        // lower-limit object so the bound sits under the operator, the way Word renders it.
        if (LatexSymbols.LowerLimitFunctions.Contains(name) && Peek.Kind == LatexTokenKind.Underscore)
        {
            Advance(); // _
            var limit = ParseArgument();
            return new MathLowerLimit(new MathFunctionName(name), limit);
        }

        if (LatexSymbols.NaryOperators.TryGetValue(name, out var naryGlyph))
        {
            return ParseNary(naryGlyph);
        }

        if (LatexSymbols.Accents.TryGetValue(name, out var accentChar))
        {
            return new MathAccent(accentChar, ParseArgument());
        }

        if (LatexSymbols.Functions.Contains(name))
        {
            return new MathFunctionName(name);
        }

        if (LatexSymbols.Map.TryGetValue(name, out var glyph))
        {
            return new MathText(glyph);
        }

        // Spacing commands are recognized (consumed) but emit nothing visible.
        if (name is "," or ";" or "!" or ":" or " " or "quad" or "qquad" or "thinspace" or "medspace" or "thickspace")
        {
            return new MathText(name is "quad" or "qquad" ? "  " : " ");
        }

        // An escaped literal delimiter/brace: \{ \} \% \& \# \$ \_ render that char.
        if (name.Length == 1 && !char.IsAsciiLetter(name[0]))
        {
            return new MathText(name);
        }

        // Unknown command: emit the literal token text and record it.
        _unknown.Add("\\" + name);
        return new MathText("\\" + name);
    }

    /// <summary>An n-ary operator with optional <c>_</c> lower / <c>^</c> upper limits, then a body atom.</summary>
    private MathNode ParseNary(string glyph)
    {
        MathNode? sub = null;
        MathNode? sup = null;
        while (Peek.Kind is LatexTokenKind.Caret or LatexTokenKind.Underscore)
        {
            var isSup = Advance().Kind == LatexTokenKind.Caret;
            var arg = ParseArgument();
            if (isSup)
            {
                sup = arg;
            }
            else
            {
                sub = arg;
            }
        }

        var body = ParseNaryBody();
        return new MathNary(glyph, sub, sup, body);
    }

    /// <summary>
    /// The integrand/summand following an n-ary operator: the rest of the current
    /// sequence up to the next separator or closer. Keeping it greedy matches how
    /// readers expect <c>\sum_i a_i</c> to scope its body.
    /// </summary>
    private MathNode ParseNaryBody()
    {
        var items = new List<MathNode>();
        while (!AtEnd)
        {
            var kind = Peek.Kind;
            if (kind is LatexTokenKind.CloseBrace or LatexTokenKind.CloseBracket
                or LatexTokenKind.Ampersand or LatexTokenKind.RowBreak)
            {
                break;
            }

            if (kind is LatexTokenKind.Caret or LatexTokenKind.Underscore)
            {
                var baseNode = items.Count > 0 ? items[^1] : new MathText(string.Empty);
                if (items.Count > 0)
                {
                    items.RemoveAt(items.Count - 1);
                }

                items.Add(ParseScripts(baseNode));
                continue;
            }

            // Another n-ary operator starts its own object; stop so we don't nest greedily.
            if (kind == LatexTokenKind.Command && LatexSymbols.NaryOperators.ContainsKey(Peek.Value) && items.Count > 0)
            {
                break;
            }

            items.Add(ParseAtom());
        }

        return new MathList(items);
    }

    /// <summary>
    /// <c>\left&lt;d&gt; … \right&lt;d&gt;</c>. The delimiter glyph follows the command as a
    /// single token (a symbol command like <c>\{</c> or a bracket/paren char).
    /// </summary>
    private MathNode ParseLeftRight()
    {
        var open = ReadDelimiter();
        var body = ParseDelimiterBody();
        var close = "."; // default if \right is missing
        if (Peek.Kind == LatexTokenKind.Command && Peek.Value == "right")
        {
            Advance();
            close = ReadDelimiter();
        }

        return new MathDelimiter(DelimiterGlyph(open, opening: true), DelimiterGlyph(close, opening: false), body);
    }

    /// <summary>The content between <c>\left</c> and its matching <c>\right</c>.</summary>
    private MathNode ParseDelimiterBody()
    {
        var items = new List<MathNode>();
        while (!AtEnd)
        {
            if (Peek.Kind == LatexTokenKind.Command && Peek.Value == "right")
            {
                break;
            }

            if (Peek.Kind is LatexTokenKind.CloseBrace or LatexTokenKind.CloseBracket)
            {
                break;
            }

            if (Peek.Kind is LatexTokenKind.Caret or LatexTokenKind.Underscore)
            {
                var baseNode = items.Count > 0 ? items[^1] : new MathText(string.Empty);
                if (items.Count > 0)
                {
                    items.RemoveAt(items.Count - 1);
                }

                items.Add(ParseScripts(baseNode));
                continue;
            }

            items.Add(ParseAtom());
        }

        return new MathList(items);
    }

    /// <summary>Reads the delimiter token after <c>\left</c>/<c>\right</c> and returns its key.</summary>
    private string ReadDelimiter()
    {
        var token = Peek;
        switch (token.Kind)
        {
            case LatexTokenKind.Command:
                Advance();

                // \{ and \} arrive as single-char escape commands; keep the escaped form as the key.
                return token.Value is "{" or "}" ? "\\" + token.Value : token.Value;

            case LatexTokenKind.Text:
                Advance();
                return token.Value;

            case LatexTokenKind.OpenBracket:
                Advance();
                return "[";

            case LatexTokenKind.CloseBracket:
                Advance();
                return "]";

            default:
                return ".";
        }
    }

    /// <summary>Resolves a delimiter key to the literal glyph stored in OMML (empty = invisible).</summary>
    private static string DelimiterGlyph(string key, bool opening)
    {
        var table = opening ? LatexSymbols.OpenDelimiters : LatexSymbols.CloseDelimiters;
        if (table.TryGetValue(key, out var glyph))
        {
            return glyph;
        }

        // A symbol-command delimiter (\langle, \lceil, …) resolves through the symbol map.
        if (LatexSymbols.Map.TryGetValue(key, out var sym))
        {
            return sym;
        }

        return key == "." ? string.Empty : key;
    }

    /// <summary>
    /// <c>\begin{env} … \end{env}</c> for matrix environments. Cells are split on
    /// <c>&amp;</c>, rows on <c>\\</c>. pmatrix/bmatrix/Bmatrix/vmatrix carry the
    /// bracket pair; plain matrix has none. The aligned/align/cases family routes
    /// to <see cref="ParseEqArray"/> (an OMML <c>m:eqArr</c>) instead.
    /// </summary>
    private MathNode ParseEnvironment()
    {
        var env = ReadRawGroupText();

        // The aligned / cases family becomes an equation array, not a matrix.
        switch (env)
        {
            case "aligned" or "align" or "align*" or "alignedat" or "alignat" or "alignat*"
                or "gathered" or "gather" or "gather*" or "split" or "multline" or "multline*":
                return ParseEqArray(env, string.Empty, string.Empty);
            case "cases":
                return ParseEqArray(env, "{", string.Empty);
            case "rcases":
                return ParseEqArray(env, string.Empty, "}");
            default:
                break;
        }

        var (open, close) = env switch
        {
            "pmatrix" => ("(", ")"),
            "bmatrix" => ("[", "]"),
            "Bmatrix" => ("{", "}"),
            "vmatrix" => ("|", "|"),
            "Vmatrix" => ("‖", "‖"),
            _ => (string.Empty, string.Empty),
        };

        return new MathMatrix(open, close, ReadEnvironmentRows());
    }

    /// <summary>
    /// An aligned/cases environment: the same <c>&amp;</c>-cell, <c>\\</c>-row grid
    /// as a matrix, but emitted as an OMML <c>m:eqArr</c> (left-aligned rows with
    /// alignment columns) optionally wrapped in an open/close brace.
    /// <c>alignedat</c>/<c>alignat</c> carry a mandatory <c>{n}</c> column-count
    /// argument that is consumed and ignored (OMML sizes columns automatically).
    /// </summary>
    private MathNode ParseEqArray(string env, string open, string close)
    {
        if (env is "alignedat" or "alignat" or "alignat*")
        {
            _ = ReadRawGroupText(); // consume the {n} column-pair count
        }

        return new MathEqArray(open, close, ReadEnvironmentRows());
    }

    /// <summary>
    /// <c>\cases{ a &amp; b \\ c &amp; d }</c> brace-shorthand: the rows live in a
    /// single braced group rather than a <c>\begin…\end</c> pair, but split on
    /// <c>&amp;</c>/<c>\\</c> identically and emit the same braced equation array.
    /// </summary>
    private MathNode ParseCasesShorthand()
    {
        if (Peek.Kind != LatexTokenKind.OpenBrace)
        {
            return new MathEqArray("{", string.Empty, [[new MathList([])]]);
        }

        Advance(); // {
        var rows = ReadEnvironmentRows(untilCloseBrace: true);
        Expect(LatexTokenKind.CloseBrace);
        return new MathEqArray("{", string.Empty, rows);
    }

    /// <summary>
    /// Collects environment rows: cells split on <c>&amp;</c>, rows on <c>\\</c>,
    /// stopping at the matching <c>\end{…}</c> (default) or a <c>}</c> (shorthand).
    /// Shared by matrix, aligned and cases parsing so they segment identically.
    /// </summary>
    private List<IReadOnlyList<MathNode>> ReadEnvironmentRows(bool untilCloseBrace = false)
    {
        var rows = new List<IReadOnlyList<MathNode>>();
        var currentRow = new List<MathNode>();
        var cell = new List<MathNode>();
        var sawRowBreak = false;

        void FlushCell()
        {
            currentRow.Add(new MathList([.. cell]));
            cell = [];
        }

        while (!AtEnd)
        {
            if (untilCloseBrace && Peek.Kind == LatexTokenKind.CloseBrace)
            {
                break;
            }

            if (!untilCloseBrace && Peek.Kind == LatexTokenKind.Command && Peek.Value == "end")
            {
                Advance();
                _ = ReadRawGroupText(); // consume {env}
                break;
            }

            if (Peek.Kind == LatexTokenKind.Ampersand)
            {
                Advance();
                FlushCell();
                continue;
            }

            if (Peek.Kind == LatexTokenKind.RowBreak)
            {
                Advance();
                FlushCell();
                rows.Add(currentRow);
                currentRow = [];
                sawRowBreak = true;
                continue;
            }

            // Scripts inside a cell bind to the preceding atom, exactly as in a sequence.
            if (Peek.Kind is LatexTokenKind.Caret or LatexTokenKind.Underscore)
            {
                var baseNode = cell.Count > 0 ? cell[^1] : new MathText(string.Empty);
                if (cell.Count > 0)
                {
                    cell.RemoveAt(cell.Count - 1);
                }

                cell.Add(ParseScripts(baseNode));
                continue;
            }

            cell.Add(ParseAtom());
        }

        // Close the final cell/row unless the last token was a row break with no trailing content.
        var trailingRowEmpty = sawRowBreak && cell.Count == 0 && currentRow.Count == 0;
        if (!trailingRowEmpty)
        {
            FlushCell();
            rows.Add(currentRow);
        }

        if (rows.Count == 0)
        {
            rows.Add([new MathList([])]);
        }

        return rows;
    }

    /// <summary>
    /// <c>\overbrace{body}^{label}</c> / <c>\underbrace{body}_{label}</c>: the body
    /// is the required argument; a brace label is the optional script on the far
    /// side (above for overbrace, below for underbrace).
    /// </summary>
    private MathNode ParseBrace(bool below)
    {
        var body = ParseArgument();
        MathNode? label = null;

        // The label is attached with ^ (overbrace) or _ (underbrace); accept either
        // and treat it as the brace label rather than an ordinary script.
        if (Peek.Kind is LatexTokenKind.Caret or LatexTokenKind.Underscore)
        {
            Advance();
            label = ParseArgument();
        }

        return new MathBrace(body, below, label);
    }

    /// <summary>
    /// Reads a braced group as raw concatenated text (for <c>\text{…}</c>,
    /// <c>\begin{…}</c>): no math interpretation, commands lose their backslash.
    /// </summary>
    private string ReadRawGroupText()
    {
        if (Peek.Kind != LatexTokenKind.OpenBrace)
        {
            // Unbraced single token (rare): take its text.
            return Peek.Kind == LatexTokenKind.Text ? Advance().Value : string.Empty;
        }

        Advance(); // {
        var sb = new System.Text.StringBuilder();
        var depth = 1;
        while (!AtEnd && depth > 0)
        {
            var token = Advance();
            switch (token.Kind)
            {
                case LatexTokenKind.OpenBrace:
                    depth++;
                    break;
                case LatexTokenKind.CloseBrace:
                    depth--;
                    if (depth > 0)
                    {
                        sb.Append('}');
                    }

                    break;
                case LatexTokenKind.Command:
                    // A symbol command inside \text becomes its glyph; otherwise its name.
                    sb.Append(LatexSymbols.Map.TryGetValue(token.Value, out var g) ? g : token.Value);
                    break;
                case LatexTokenKind.End:
                    break;
                default:
                    sb.Append(token.Value);
                    break;
            }
        }

        return sb.ToString();
    }

    private void Expect(LatexTokenKind kind)
    {
        if (Peek.Kind == kind)
        {
            Advance();
        }

        // Missing closer: tolerate (degrade gracefully) rather than throw.
    }

    /// <summary>Formats a node's plain text for diagnostics (unused in OMML emit, kept for tests).</summary>
    internal static string DebugText(MathNode node) => node switch
    {
        MathText t => t.Value,
        MathLiteral l => l.Value,
        MathList list => string.Concat(list.Items.Select(DebugText)),
        _ => node.GetType().Name,
    };
}
