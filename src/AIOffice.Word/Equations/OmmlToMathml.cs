using System.Text;
using DocumentFormat.OpenXml;
using M = DocumentFormat.OpenXml.Math;

namespace AIOffice.Word.Equations;

/// <summary>
/// Converts the OMML this project emits (and the common constructs Word writes)
/// into presentation MathML (<c>&lt;math&gt;</c>). Used by the HTML renderer so
/// equations display natively in a browser. The converter walks the OMML element
/// model the same way the emitter built it, mapping each math object to its
/// MathML counterpart (mfrac, msqrt/mroot, msub/msup/msubsup, mfenced, mover,
/// mtable). Unhandled shapes degrade to their inner text inside an <c>&lt;mtext&gt;</c>.
/// </summary>
internal static class OmmlToMathml
{
    /// <summary>Renders an <c>m:oMath</c> as a complete MathML <c>&lt;math&gt;</c> element string.</summary>
    public static string ToMathml(M.OfficeMath oMath, bool display)
    {
        var sb = new StringBuilder();
        sb.Append("<math xmlns=\"http://www.w3.org/1998/Math/MathML\"");
        if (display)
        {
            sb.Append(" display=\"block\"");
        }

        sb.Append('>');
        AppendRow(sb, oMath.ChildElements);
        sb.Append("</math>");
        return sb.ToString();
    }

    /// <summary>A sequence of math elements; wrapped in mrow only when there is more than one.</summary>
    private static void AppendRow(StringBuilder sb, IEnumerable<OpenXmlElement> elements)
    {
        var items = elements.Where(IsMathContent).ToList();
        if (items.Count == 1)
        {
            AppendElement(sb, items[0]);
            return;
        }

        sb.Append("<mrow>");
        foreach (var item in items)
        {
            AppendElement(sb, item);
        }

        sb.Append("</mrow>");
    }

    private static bool IsMathContent(OpenXmlElement element) => element switch
    {
        M.RunProperties or M.ControlProperties => false,
        M.FractionProperties or M.RadicalProperties or M.NaryProperties or M.DelimiterProperties => false,
        M.SubscriptProperties or M.SuperscriptProperties or M.SubSuperscriptProperties => false,
        M.AccentProperties or M.BarProperties or M.MatrixProperties or M.ParagraphProperties => false,
        _ => true,
    };

    private static void AppendElement(StringBuilder sb, OpenXmlElement element)
    {
        switch (element)
        {
            case M.Run run:
                AppendRun(sb, run);
                break;

            case M.Fraction fraction:
                sb.Append("<mfrac>");
                AppendArgument(sb, fraction.Numerator);
                AppendArgument(sb, fraction.Denominator);
                sb.Append("</mfrac>");
                break;

            case M.Radical radical:
                AppendRadical(sb, radical);
                break;

            case M.Superscript sup:
                sb.Append("<msup>");
                AppendArgument(sb, sup.Base);
                AppendArgument(sb, sup.SuperArgument);
                sb.Append("</msup>");
                break;

            case M.Subscript sub:
                sb.Append("<msub>");
                AppendArgument(sb, sub.Base);
                AppendArgument(sb, sub.SubArgument);
                sb.Append("</msub>");
                break;

            case M.SubSuperscript subsup:
                sb.Append("<msubsup>");
                AppendArgument(sb, subsup.Base);
                AppendArgument(sb, subsup.SubArgument);
                AppendArgument(sb, subsup.SuperArgument);
                sb.Append("</msubsup>");
                break;

            case M.Nary nary:
                AppendNary(sb, nary);
                break;

            case M.Delimiter delimiter:
                AppendDelimiter(sb, delimiter);
                break;

            case M.Accent accent:
                sb.Append("<mover accent=\"true\">");
                AppendArgument(sb, accent.Base);
                sb.Append("<mo>").Append(Escape(accent.AccentProperties?.AccentChar?.Val?.Value ?? "^")).Append("</mo>");
                sb.Append("</mover>");
                break;

            case M.Bar bar:
                sb.Append("<mover accent=\"true\">");
                AppendArgument(sb, bar.Base);
                sb.Append("<mo>&#x00AF;</mo></mover>"); // combining overline
                break;

            case M.Matrix matrix:
                AppendMatrix(sb, matrix);
                break;

            default:
                if (element.InnerText.Length > 0)
                {
                    sb.Append("<mtext>").Append(Escape(element.InnerText)).Append("</mtext>");
                }

                break;
        }
    }

    /// <summary>An argument element (m:num, m:e, …): its content as a single MathML node/row.</summary>
    private static void AppendArgument(StringBuilder sb, OpenXmlElement? argument)
    {
        if (argument is null)
        {
            sb.Append("<mrow/>");
            return;
        }

        AppendRow(sb, argument.ChildElements);
    }

    /// <summary>
    /// A math run becomes mi/mn/mo tokens: letters are identifiers, digit runs are
    /// numbers, everything else is an operator. m:nor (normal text) runs become
    /// mtext (upright). Splitting on character class keeps spacing/italics correct.
    /// </summary>
    private static void AppendRun(StringBuilder sb, M.Run run)
    {
        var text = run.GetFirstChild<M.Text>()?.Text ?? run.InnerText;
        if (text.Length == 0)
        {
            return;
        }

        var normal = run.MathRunProperties?.GetFirstChild<M.NormalText>() is not null;
        if (normal)
        {
            sb.Append("<mtext>").Append(Escape(text)).Append("</mtext>");
            return;
        }

        var i = 0;
        while (i < text.Length)
        {
            var c = text[i];
            if (char.IsWhiteSpace(c))
            {
                i++;
                continue;
            }

            if (char.IsDigit(c) || c == '.')
            {
                var start = i;
                while (i < text.Length && (char.IsDigit(text[i]) || text[i] == '.'))
                {
                    i++;
                }

                sb.Append("<mn>").Append(Escape(text[start..i])).Append("</mn>");
            }
            else if (char.IsLetter(c))
            {
                sb.Append("<mi>").Append(Escape(c.ToString())).Append("</mi>");
                i++;
            }
            else
            {
                sb.Append("<mo>").Append(Escape(c.ToString())).Append("</mo>");
                i++;
            }
        }
    }

    private static void AppendRadical(StringBuilder sb, M.Radical radical)
    {
        var hideDegree = radical.RadicalProperties?.GetFirstChild<M.HideDegree>()?.Val?.Value == M.BooleanValues.True;
        var degreeEmpty = radical.Degree is null || radical.Degree.InnerText.Length == 0;

        if (hideDegree || degreeEmpty)
        {
            sb.Append("<msqrt>");
            AppendArgument(sb, radical.Base);
            sb.Append("</msqrt>");
        }
        else
        {
            sb.Append("<mroot>");
            AppendArgument(sb, radical.Base);
            AppendArgument(sb, radical.Degree);
            sb.Append("</mroot>");
        }
    }

    private static void AppendNary(StringBuilder sb, M.Nary nary)
    {
        var op = nary.NaryProperties?.GetFirstChild<M.AccentChar>()?.Val?.Value ?? "∫";
        var hasSub = nary.NaryProperties?.GetFirstChild<M.HideSubArgument>()?.Val?.Value != M.BooleanValues.True;
        var hasSup = nary.NaryProperties?.GetFirstChild<M.HideSuperArgument>()?.Val?.Value != M.BooleanValues.True;

        sb.Append("<mrow>");
        if (hasSub && hasSup)
        {
            sb.Append("<munderover><mo>").Append(Escape(op)).Append("</mo>");
            AppendArgument(sb, nary.SubArgument);
            AppendArgument(sb, nary.SuperArgument);
            sb.Append("</munderover>");
        }
        else if (hasSub)
        {
            sb.Append("<munder><mo>").Append(Escape(op)).Append("</mo>");
            AppendArgument(sb, nary.SubArgument);
            sb.Append("</munder>");
        }
        else if (hasSup)
        {
            sb.Append("<mover><mo>").Append(Escape(op)).Append("</mo>");
            AppendArgument(sb, nary.SuperArgument);
            sb.Append("</mover>");
        }
        else
        {
            sb.Append("<mo>").Append(Escape(op)).Append("</mo>");
        }

        AppendArgument(sb, nary.Base);
        sb.Append("</mrow>");
    }

    private static void AppendDelimiter(StringBuilder sb, M.Delimiter delimiter)
    {
        var open = delimiter.DelimiterProperties?.GetFirstChild<M.BeginChar>()?.Val?.Value ?? "(";
        var close = delimiter.DelimiterProperties?.GetFirstChild<M.EndChar>()?.Val?.Value ?? ")";
        if (open == ".")
        {
            open = string.Empty;
        }

        if (close == ".")
        {
            close = string.Empty;
        }

        sb.Append("<mrow><mo>").Append(Escape(open)).Append("</mo>");
        foreach (var baseArg in delimiter.Elements<M.Base>())
        {
            AppendArgument(sb, baseArg);
        }

        sb.Append("<mo>").Append(Escape(close)).Append("</mo></mrow>");
    }

    private static void AppendMatrix(StringBuilder sb, M.Matrix matrix)
    {
        sb.Append("<mtable>");
        foreach (var row in matrix.Elements<M.MatrixRow>())
        {
            sb.Append("<mtr>");
            foreach (var cell in row.Elements<M.Base>())
            {
                sb.Append("<mtd>");
                AppendArgument(sb, cell);
                sb.Append("</mtd>");
            }

            sb.Append("</mtr>");
        }

        sb.Append("</mtable>");
    }

    private static string Escape(string text) => text
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal);
}
