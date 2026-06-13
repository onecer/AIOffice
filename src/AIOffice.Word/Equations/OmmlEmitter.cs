using DocumentFormat.OpenXml;
using M = DocumentFormat.OpenXml.Math;

namespace AIOffice.Word.Equations;

/// <summary>
/// Emits validator-clean OMML (<c>m:oMath</c> and below) from the math tree. The
/// emitter targets the DocumentFormat.OpenXml.Math element model directly, so the
/// output is real OOXML — Word renders it natively and the OpenXmlValidator
/// passes with zero errors. Argument types (<c>m:num</c>, <c>m:den</c>, <c>m:e</c>,
/// …) wrap their child run/list exactly as the schema requires.
/// </summary>
internal static class OmmlEmitter
{
    /// <summary>OMML on/off elements take an <c>EnumValue&lt;BooleanValues&gt;</c>, not a bare bool.</summary>
    private static M.BooleanValues Bool(bool on) => on ? M.BooleanValues.True : M.BooleanValues.False;

    /// <summary>Builds an <c>m:oMath</c> for one parsed equation.</summary>
    public static M.OfficeMath ToOfficeMath(MathNode root)
    {
        var math = new M.OfficeMath();
        foreach (var element in EmitNodes(root))
        {
            math.AppendChild(element);
        }

        // An empty equation still needs a run so m:oMath is schema-valid and visible.
        if (!math.HasChildren)
        {
            math.AppendChild(Run(string.Empty));
        }

        return math;
    }

    /// <summary>The OMML elements for one node (a list expands to its concatenated children).</summary>
    private static IEnumerable<OpenXmlElement> EmitNodes(MathNode node)
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
    private static OpenXmlElement EmitObject(MathNode node) => node switch
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

    /// <summary>A math run <c>m:r/m:t</c> holding literal text; normalText marks it upright (\text, functions).</summary>
    private static M.Run Run(string text, bool normalText = false)
    {
        var run = new M.Run();
        if (normalText)
        {
            // m:nor flips the run to upright (non-italic) text — the way Word stores
            // \text{…} and function names. m:rPr accepts only m:lit and m:nor here.
            run.MathRunProperties = new M.RunProperties(new M.NormalText());
        }

        var t = new M.Text(text);
        if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1])))
        {
            t.Space = SpaceProcessingModeValues.Preserve;
        }

        run.AppendChild(t);
        return run;
    }

    /// <summary>Wraps a node's emitted content in an argument element (m:num, m:den, m:e, …).</summary>
    private static T Argument<T>(MathNode node)
        where T : OpenXmlElement, new()
    {
        var arg = new T();
        foreach (var element in EmitNodes(node))
        {
            arg.AppendChild(element);
        }

        return arg;
    }

    // ----------------------------------------------------------- composites

    private static M.Fraction Fraction(MathFraction frac) => new(
        Argument<M.Numerator>(frac.Numerator),
        Argument<M.Denominator>(frac.Denominator));

    private static M.Radical Radical(MathRadical rad)
    {
        var radical = new M.Radical();
        if (rad.Degree is null)
        {
            // Square root: m:radPr/m:degHide=1, plus an empty m:deg the schema still requires.
            radical.RadicalProperties = new M.RadicalProperties(new M.HideDegree { Val = Bool(true) });
            radical.Degree = new M.Degree();
        }
        else
        {
            radical.Degree = Argument<M.Degree>(rad.Degree);
        }

        radical.Base = Argument<M.Base>(rad.Radicand);
        return radical;
    }

    private static M.Nary Nary(MathNary nary)
    {
        var hasSub = nary.Sub is not null;
        var hasSup = nary.Sup is not null;
        var props = new M.NaryProperties(
            new M.AccentChar { Val = nary.Operator },
            new M.LimitLocation { Val = M.LimitLocationValues.SubscriptSuperscript },
            new M.GrowOperators { Val = Bool(true) },
            new M.HideSubArgument { Val = Bool(!hasSub) },
            new M.HideSuperArgument { Val = Bool(!hasSup) });

        return new M.Nary(
            props,
            Argument<M.SubArgument>(nary.Sub ?? new MathList([])),
            Argument<M.SuperArgument>(nary.Sup ?? new MathList([])),
            Argument<M.Base>(nary.Body));
    }

    private static OpenXmlElement Script(MathScript script)
    {
        if (script is { Sub: not null, Sup: not null })
        {
            return new M.SubSuperscript(
                new M.SubSuperscriptProperties(),
                Argument<M.Base>(script.Base),
                Argument<M.SubArgument>(script.Sub),
                Argument<M.SuperArgument>(script.Sup));
        }

        if (script.Sup is not null)
        {
            return new M.Superscript(
                new M.SuperscriptProperties(),
                Argument<M.Base>(script.Base),
                Argument<M.SuperArgument>(script.Sup));
        }

        if (script.Sub is not null)
        {
            return new M.Subscript(
                new M.SubscriptProperties(),
                Argument<M.Base>(script.Base),
                Argument<M.SubArgument>(script.Sub));
        }

        // No scripts attached (degenerate; the parser never builds this): emit the base atom.
        return EmitNodes(script.Base).FirstOrDefault() ?? Run(string.Empty);
    }

    private static M.Delimiter Delimiter(MathDelimiter delim)
    {
        var props = new M.DelimiterProperties(
            new M.BeginChar { Val = delim.Open.Length > 0 ? delim.Open : "." },
            new M.EndChar { Val = delim.Close.Length > 0 ? delim.Close : "." },
            new M.GrowOperators { Val = Bool(true) });

        var d = new M.Delimiter(props);
        d.AppendChild(Argument<M.Base>(delim.Body));
        return d;
    }

    private static M.Accent Accent(MathAccent accent) => new(
        new M.AccentProperties(new M.AccentChar { Val = accent.Char }),
        Argument<M.Base>(accent.Base));

    private static M.Bar Bar(MathBar bar) => new(
        new M.BarProperties(new M.Position { Val = M.VerticalJustificationValues.Top }),
        Argument<M.Base>(bar.Base));

    /// <summary>
    /// A matrix. A plain <c>matrix</c> emits a bare <c>m:m</c>; a bracketed
    /// environment (pmatrix/bmatrix/…) wraps it in a grow-operator delimiter so
    /// the brackets size to the rows, exactly as Word stores them.
    /// </summary>
    private static OpenXmlElement Matrix(MathMatrix matrix)
    {
        var columns = matrix.Rows.Count > 0 ? matrix.Rows.Max(r => r.Count) : 1;

        var columnProps = new M.MatrixColumns();
        columnProps.AppendChild(new M.MatrixColumn(
            new M.MatrixColumnProperties(
                new M.MatrixColumnCount { Val = columns },
                new M.MatrixColumnJustification { Val = M.HorizontalAlignmentValues.Center })));

        var m = new M.Matrix(new M.MatrixProperties(
            new M.BaseJustification { Val = M.VerticalAlignmentValues.Center },
            new M.HidePlaceholder { Val = Bool(true) },
            columnProps));

        foreach (var row in matrix.Rows)
        {
            var matrixRow = new M.MatrixRow();
            for (var c = 0; c < columns; c++)
            {
                var cell = c < row.Count ? row[c] : new MathList([]);
                matrixRow.AppendChild(Argument<M.Base>(cell));
            }

            m.AppendChild(matrixRow);
        }

        if (matrix.Open.Length == 0 && matrix.Close.Length == 0)
        {
            return m;
        }

        var props = new M.DelimiterProperties(
            new M.BeginChar { Val = matrix.Open.Length > 0 ? matrix.Open : "." },
            new M.EndChar { Val = matrix.Close.Length > 0 ? matrix.Close : "." },
            new M.GrowOperators { Val = Bool(true) });
        var baseArg = new M.Base();
        baseArg.AppendChild(m);
        return new M.Delimiter(props, baseArg);
    }
}
