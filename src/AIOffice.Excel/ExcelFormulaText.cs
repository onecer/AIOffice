namespace AIOffice.Excel;

/// <summary>
/// (1.5) Shared formula-text parsing helpers for the write-time evaluators
/// (financial, dynamic-array, scalar). The leading-function-name / inside-parens /
/// top-level-comma-split logic was duplicated privately in each evaluator since
/// v1.4; this is the one place it lives now. Pure string handling, no model
/// access — every evaluator strips a leading <c>=</c> before calling in.
/// </summary>
internal static class ExcelFormulaText
{
    /// <summary>
    /// The leading function name (letters/underscores up to the first <c>(</c>), or
    /// null. ClosedXML serializes "future" functions (XLOOKUP, IFS, LET, TEXTSPLIT,
    /// …) with an <c>_xlfn.</c> / <c>_xlfn._xlws.</c> prefix in the A1 form; the
    /// prefix is stripped so the bare function name is recognized.
    /// </summary>
    public static string? LeadingFunctionName(string formula)
    {
        var text = StripFuturePrefix((formula.StartsWith('=') ? formula[1..] : formula).TrimStart());
        var i = 0;
        while (i < text.Length && (char.IsLetter(text[i]) || char.IsDigit(text[i]) || text[i] == '_'))
        {
            i++;
        }

        return i > 0 && i < text.Length && text[i] == '(' ? text[..i] : null;
    }

    /// <summary>
    /// Strips ClosedXML's <c>_xlfn.</c> and <c>_xlfn._xlws.</c> future-function
    /// prefixes from the start of a formula body so the evaluators parse the bare
    /// function name and arguments.
    /// </summary>
    public static string StripFuturePrefix(string body)
    {
        var text = body.TrimStart();
        while (true)
        {
            if (text.StartsWith("_xlfn.", StringComparison.Ordinal))
            {
                text = text[6..];
            }
            else if (text.StartsWith("_xlws.", StringComparison.Ordinal))
            {
                text = text[6..];
            }
            else
            {
                return text;
            }
        }
    }

    /// <summary>The substring inside the outermost parentheses, or empty when there are none.</summary>
    public static string InsideParens(string body)
    {
        var text = body.TrimStart();
        var open = text.IndexOf('(', StringComparison.Ordinal);
        return open >= 0 && text.EndsWith(')') ? text[(open + 1)..^1] : string.Empty;
    }

    /// <summary>Splits an argument list on top-level commas (respecting nested parens/braces and quotes).</summary>
    public static List<string> SplitArgs(string inside)
    {
        var args = new List<string>();
        if (inside.Trim().Length == 0)
        {
            return args;
        }

        var depth = 0;
        var inQuote = false;
        var start = 0;
        for (var i = 0; i < inside.Length; i++)
        {
            var ch = inside[i];
            if (ch == '"')
            {
                inQuote = !inQuote;
            }
            else if (!inQuote && ch is '(' or '{')
            {
                depth++;
            }
            else if (!inQuote && ch is ')' or '}')
            {
                depth--;
            }
            else if (!inQuote && depth == 0 && ch == ',')
            {
                args.Add(inside[start..i].Trim());
                start = i + 1;
            }
        }

        args.Add(inside[start..].Trim());
        return args;
    }
}
