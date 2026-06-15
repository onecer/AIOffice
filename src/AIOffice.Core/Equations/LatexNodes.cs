namespace AIOffice.Core.Equations;

/// <summary>
/// The math expression tree produced by <see cref="LatexParser"/> and consumed
/// by an OMML emitter. Every node is a pure data record; the tree is the
/// format-neutral middle layer between LaTeX source and OMML. Lives in Core (M10)
/// so both the Word and Pptx handlers consume the one shared converter.
/// </summary>
public abstract record MathNode;

/// <summary>An ordered list of sibling nodes (the body of a group or argument).</summary>
public sealed record MathList(IReadOnlyList<MathNode> Items) : MathNode;

/// <summary>A literal run of math text (variables, numbers, operators, mapped symbols).</summary>
public sealed record MathText(string Value) : MathNode;

/// <summary>Upright literal text from <c>\text{…}</c> (a non-italic run).</summary>
public sealed record MathLiteral(string Value) : MathNode;

/// <summary>An upright named operator/function such as <c>sin</c> or <c>lim</c>.</summary>
public sealed record MathFunctionName(string Name) : MathNode;

/// <summary><c>base^sup</c>, <c>base_sub</c> or <c>base_sub^sup</c>.</summary>
public sealed record MathScript(MathNode Base, MathNode? Sub, MathNode? Sup) : MathNode;

/// <summary><c>\frac{num}{den}</c>.</summary>
public sealed record MathFraction(MathNode Numerator, MathNode Denominator) : MathNode;

/// <summary><c>\sqrt{radicand}</c> or <c>\sqrt[degree]{radicand}</c> (degree null = square root).</summary>
public sealed record MathRadical(MathNode? Degree, MathNode Radicand) : MathNode;

/// <summary>An n-ary operator (<c>\sum</c>, <c>\int</c>, …) with optional limits and a body.</summary>
public sealed record MathNary(string Operator, MathNode? Sub, MathNode? Sup, MathNode Body) : MathNode;

/// <summary>A bracketed group <c>\left( … \right)</c> or bare <c>( … )</c>; chars are the literal glyphs.</summary>
public sealed record MathDelimiter(string Open, string Close, MathNode Body) : MathNode;

/// <summary>An accented base: <c>\hat{x}</c>, <c>\vec{v}</c>, … (Char is the combining glyph).</summary>
public sealed record MathAccent(string Char, MathNode Base) : MathNode;

/// <summary>An overbar (<c>\bar</c>/<c>\overline</c>) or underline (<c>\underline</c>): a bar above or below the base.</summary>
public sealed record MathBar(MathNode Base, bool Below = false) : MathNode;

/// <summary>A matrix/array: rows of cells, each cell a node, with the bracket pair (empty = plain matrix).</summary>
public sealed record MathMatrix(string Open, string Close, IReadOnlyList<IReadOnlyList<MathNode>> Rows) : MathNode;

/// <summary>
/// An equation array (<c>\begin{aligned}</c>, <c>\begin{cases}</c>, …): rows
/// separated by <c>\\</c>, each row split into alignment columns on <c>&amp;</c>.
/// Open/Close hold an optional surrounding brace pair (<c>{</c> for cases, empty
/// for aligned). Emits an OMML <c>m:eqArr</c> (optionally inside a delimiter).
/// </summary>
public sealed record MathEqArray(string Open, string Close, IReadOnlyList<IReadOnlyList<MathNode>> Rows) : MathNode;

/// <summary>A binomial coefficient <c>\binom{n}{k}</c> — a no-bar fraction inside parentheses.</summary>
public sealed record MathBinomial(MathNode Top, MathNode Bottom) : MathNode;

/// <summary>
/// A horizontal brace group: <c>\overbrace{…}</c> (above) or <c>\underbrace{…}</c>
/// (below), with an optional script label attached on the far side via <c>^</c>/<c>_</c>.
/// </summary>
public sealed record MathBrace(MathNode Base, bool Below, MathNode? Label) : MathNode;

/// <summary>
/// A function/operator carrying a limit placed directly below it (<c>\lim_{x\to0}</c>,
/// <c>\max_i</c>): emits an OMML <c>m:limLow</c> so the bound sits under the operator.
/// </summary>
public sealed record MathLowerLimit(MathNode Base, MathNode Limit) : MathNode;
