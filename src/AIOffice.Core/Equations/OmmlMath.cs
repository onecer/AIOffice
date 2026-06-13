using System.Xml.Linq;

namespace AIOffice.Core.Equations;

/// <summary>
/// The M10 shared LaTeX→OMML converter: a pure <see cref="System.Xml.Linq"/>
/// producer that turns a parsed <see cref="MathNode"/> tree into an
/// <c>m:oMath</c> <see cref="XElement"/>. It carries NO DocumentFormat.OpenXml
/// dependency, so both the Word handler (which loads the XElement into a
/// <c>DocumentFormat.OpenXml.Math.OfficeMath</c>) and the Pptx handler (which
/// drops it into a slide text body) consume the one converter. The emitted
/// markup is the same OMML element model Word writes natively, so the
/// OpenXmlValidator passes with zero errors after it is loaded into either
/// document.
/// </summary>
public static class OmmlMath
{
    /// <summary>The OMML (officeDocument math) namespace; the <c>m:</c> prefix the SDK uses.</summary>
    public static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>The standard <c>m</c> prefix declaration carried on the root so consumers serialize identically to the SDK.</summary>
    public const string Prefix = "m";

    /// <summary>Parses <paramref name="latex"/> and emits the <c>m:oMath</c> element plus any unrecognized tokens.</summary>
    public static (XElement OMath, IReadOnlyList<string> UnknownTokens) FromLatex(string latex)
    {
        var parsed = LatexParser.Parse(latex);
        return (ToOMath(parsed.Root), parsed.UnknownTokens);
    }

    /// <summary>Builds an <c>m:oMath</c> element for one parsed equation tree.</summary>
    public static XElement ToOMath(MathNode root)
    {
        var math = new XElement(M + "oMath");
        foreach (var element in EmitNodes(root))
        {
            math.Add(element);
        }

        // An empty equation still needs a run so m:oMath is schema-valid and visible.
        if (!math.HasElements)
        {
            math.Add(Run(string.Empty));
        }

        return math;
    }

    /// <summary>The OMML elements for one node (a list expands to its concatenated children).</summary>
    private static IEnumerable<XElement> EmitNodes(MathNode node)
    {
        switch (node)
        {
            case MathList list:
                foreach (var item in list.Items)
                {
                    foreach (var element in EmitNodes(item))
                    {
                        yield return element;
                    }
                }

                break;

            case MathText text:
                if (text.Value.Length > 0)
                {
                    yield return Run(text.Value);
                }

                break;

            case MathLiteral literal:
                yield return Run(literal.Value, normalText: true);
                break;

            case MathFunctionName fn:
                yield return Run(fn.Name, normalText: true);
                break;

            default:
                yield return EmitObject(node);
                break;
        }
    }

    /// <summary>A single composite math object (fraction, script, radical, …).</summary>
    private static XElement EmitObject(MathNode node) => node switch
    {
        MathFraction frac => Fraction(frac),
        MathRadical rad => Radical(rad),
        MathNary nary => Nary(nary),
        MathScript script => Script(script),
        MathDelimiter delim => Delimiter(delim),
        MathAccent accent => Accent(accent),
        MathBar bar => Bar(bar),
        MathMatrix matrix => Matrix(matrix),
        _ => Run(string.Empty),
    };

    // ---------------------------------------------------------------- atoms

    /// <summary>An on/off OMML element <c>&lt;m:name m:val="1|0"/&gt;</c>.</summary>
    private static XElement OnOff(string name, bool on) =>
        new(M + name, new XAttribute(M + "val", on ? "1" : "0"));

    /// <summary>A value-carrying OMML element <c>&lt;m:name m:val="…"/&gt;</c>.</summary>
    private static XElement Val(string name, string value) =>
        new(M + name, new XAttribute(M + "val", value));

    /// <summary>A math run <c>m:r/m:t</c> holding literal text; normalText marks it upright (\text, functions).</summary>
    private static XElement Run(string text, bool normalText = false)
    {
        var run = new XElement(M + "r");
        if (normalText)
        {
            // m:nor flips the run to upright (non-italic) text — the way OMML stores
            // \text{…} and function names. m:rPr accepts only m:lit and m:nor here.
            run.Add(new XElement(M + "rPr", new XElement(M + "nor")));
        }

        var t = new XElement(M + "t", text);
        if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1])))
        {
            t.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        }

        run.Add(t);
        return run;
    }

    /// <summary>Wraps a node's emitted content in an argument element (m:num, m:den, m:e, …).</summary>
    private static XElement Argument(string name, MathNode node)
    {
        var arg = new XElement(M + name);
        foreach (var element in EmitNodes(node))
        {
            arg.Add(element);
        }

        return arg;
    }

    // ----------------------------------------------------------- composites

    private static XElement Fraction(MathFraction frac) => new(
        M + "f",
        Argument("num", frac.Numerator),
        Argument("den", frac.Denominator));

    private static XElement Radical(MathRadical rad)
    {
        if (rad.Degree is null)
        {
            // Square root: m:radPr/m:degHide=1, plus an empty m:deg the schema still requires.
            return new XElement(
                M + "rad",
                new XElement(M + "radPr", OnOff("degHide", true)),
                new XElement(M + "deg"),
                Argument("e", rad.Radicand));
        }

        return new XElement(
            M + "rad",
            Argument("deg", rad.Degree),
            Argument("e", rad.Radicand));
    }

    private static XElement Nary(MathNary nary)
    {
        var hasSub = nary.Sub is not null;
        var hasSup = nary.Sup is not null;
        var props = new XElement(
            M + "naryPr",
            Val("chr", nary.Operator),
            Val("limLoc", "subSup"),
            OnOff("grow", true),
            OnOff("subHide", !hasSub),
            OnOff("supHide", !hasSup));

        return new XElement(
            M + "nary",
            props,
            Argument("sub", nary.Sub ?? new MathList([])),
            Argument("sup", nary.Sup ?? new MathList([])),
            Argument("e", nary.Body));
    }

    private static XElement Script(MathScript script)
    {
        if (script is { Sub: not null, Sup: not null })
        {
            return new XElement(
                M + "sSubSup",
                new XElement(M + "sSubSupPr"),
                Argument("e", script.Base),
                Argument("sub", script.Sub),
                Argument("sup", script.Sup));
        }

        if (script.Sup is not null)
        {
            return new XElement(
                M + "sSup",
                new XElement(M + "sSupPr"),
                Argument("e", script.Base),
                Argument("sup", script.Sup));
        }

        if (script.Sub is not null)
        {
            return new XElement(
                M + "sSub",
                new XElement(M + "sSubPr"),
                Argument("e", script.Base),
                Argument("sub", script.Sub));
        }

        // No scripts attached (degenerate; the parser never builds this): emit the base atom.
        return EmitNodes(script.Base).FirstOrDefault() ?? Run(string.Empty);
    }

    private static XElement Delimiter(MathDelimiter delim)
    {
        var props = new XElement(
            M + "dPr",
            Val("begChr", delim.Open.Length > 0 ? delim.Open : "."),
            Val("endChr", delim.Close.Length > 0 ? delim.Close : "."),
            OnOff("grow", true));

        return new XElement(M + "d", props, Argument("e", delim.Body));
    }

    private static XElement Accent(MathAccent accent) => new(
        M + "acc",
        new XElement(M + "accPr", Val("chr", accent.Char)),
        Argument("e", accent.Base));

    private static XElement Bar(MathBar bar) => new(
        M + "bar",
        new XElement(M + "barPr", Val("pos", "top")),
        Argument("e", bar.Base));

    /// <summary>
    /// A matrix. A plain <c>matrix</c> emits a bare <c>m:m</c>; a bracketed
    /// environment (pmatrix/bmatrix/…) wraps it in a grow-operator delimiter so
    /// the brackets size to the rows, exactly as Word stores them.
    /// </summary>
    private static XElement Matrix(MathMatrix matrix)
    {
        var columns = matrix.Rows.Count > 0 ? matrix.Rows.Max(r => r.Count) : 1;

        var matrixColumn = new XElement(
            M + "mc",
            new XElement(
                M + "mcPr",
                Val("count", columns.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                Val("mcJc", "center")));

        var mProps = new XElement(
            M + "mPr",
            Val("baseJc", "center"),
            OnOff("plcHide", true),
            new XElement(M + "mcs", matrixColumn));

        var m = new XElement(M + "m", mProps);

        foreach (var row in matrix.Rows)
        {
            var matrixRow = new XElement(M + "mr");
            for (var c = 0; c < columns; c++)
            {
                var cell = c < row.Count ? row[c] : new MathList([]);
                matrixRow.Add(Argument("e", cell));
            }

            m.Add(matrixRow);
        }

        if (matrix.Open.Length == 0 && matrix.Close.Length == 0)
        {
            return m;
        }

        var props = new XElement(
            M + "dPr",
            Val("begChr", matrix.Open.Length > 0 ? matrix.Open : "."),
            Val("endChr", matrix.Close.Length > 0 ? matrix.Close : "."),
            OnOff("grow", true));

        return new XElement(M + "d", props, new XElement(M + "e", m));
    }
}
