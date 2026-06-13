namespace AIOffice.Word.Equations;

/// <summary>
/// The math expression tree produced by <see cref="LatexParser"/> and consumed
/// by the OMML emitter. Every node is a pure data record; the tree is the
/// format-neutral middle layer between LaTeX source and DocumentFormat.OpenXml.Math.
/// </summary>
internal abstract record MathNode;

/// <summary>An ordered list of sibling nodes (the body of a group or argument).</summary>
internal sealed record MathList(IReadOnlyList<MathNode> Items) : MathNode;

/// <summary>A literal run of math text (variables, numbers, operators, mapped symbols).</summary>
internal sealed record MathText(string Value) : MathNode;

/// <summary>Upright literal text from <c>\text{…}</c> (a non-italic run).</summary>
internal sealed record MathLiteral(string Value) : MathNode;

/// <summary>An upright named operator/function such as <c>sin</c> or <c>lim</c>.</summary>
internal sealed record MathFunctionName(string Name) : MathNode;

/// <summary><c>base^sup</c>, <c>base_sub</c> or <c>base_sub^sup</c>.</summary>
internal sealed record MathScript(MathNode Base, MathNode? Sub, MathNode? Sup) : MathNode;

/// <summary><c>\frac{num}{den}</c>.</summary>
internal sealed record MathFraction(MathNode Numerator, MathNode Denominator) : MathNode;

/// <summary><c>\sqrt{radicand}</c> or <c>\sqrt[degree]{radicand}</c> (degree null = square root).</summary>
internal sealed record MathRadical(MathNode? Degree, MathNode Radicand) : MathNode;

/// <summary>An n-ary operator (<c>\sum</c>, <c>\int</c>, …) with optional limits and a body.</summary>
internal sealed record MathNary(string Operator, MathNode? Sub, MathNode? Sup, MathNode Body) : MathNode;

/// <summary>A bracketed group <c>\left( … \right)</c> or bare <c>( … )</c>; chars are the literal glyphs.</summary>
internal sealed record MathDelimiter(string Open, string Close, MathNode Body) : MathNode;

/// <summary>An accented base: <c>\hat{x}</c>, <c>\vec{v}</c>, … (Char is the combining glyph).</summary>
internal sealed record MathAccent(string Char, MathNode Base) : MathNode;

/// <summary>An overbar: <c>\bar{x}</c> / <c>\overline{…}</c>.</summary>
internal sealed record MathBar(MathNode Base) : MathNode;

/// <summary>A matrix/array: rows of cells, each cell a node, with the bracket pair (empty = plain matrix).</summary>
internal sealed record MathMatrix(string Open, string Close, IReadOnlyList<IReadOnlyList<MathNode>> Rows) : MathNode;
