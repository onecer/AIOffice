using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The M4 shared find/replace op for xlsx:
/// <c>{op:replace, path:&lt;scope&gt;, props:{find, replace, regex?, matchCase?,
/// wholeWord?, inFormulas?}}</c>.
/// <list type="bullet">
/// <item>Scope is a sheet (<c>/Sheet1</c>), a range (<c>/Sheet1/A1:C10</c>) or a
/// single cell.</item>
/// <item>Matching covers TEXT cell values. Numbers, booleans and dates never
/// match (their display text is a formatting concern); replacement results stay
/// literal text — no re-auto-typing.</item>
/// <item><c>inFormulas:true</c> (default false) also matches formula text in
/// its user-facing form (leading <c>=</c> included). Replaced formula cells
/// re-evaluate on save through the normal cached-value pipeline; a replacement
/// that destroys the leading <c>=</c> turns the cell into literal text.</item>
/// <item><c>regex:true</c> uses .NET regex with a 2 s match timeout (timeout →
/// <c>invalid_args</c>); literal mode escapes the pattern AND the replacement
/// (<c>$</c> is never substitution syntax there). <c>wholeWord</c> wraps the
/// pattern in word-boundary lookarounds.</item>
/// <item>Zero matches is ok:true with a <c>find_no_match</c> meta warning; the
/// per-op result carries <c>replacements</c> and up to 20 cell-path locations.</item>
/// </list>
/// </summary>
public sealed partial class ExcelHandler
{
    private const int MaxReplaceLocations = 20;

    private static readonly TimeSpan ReplaceTimeout = TimeSpan.FromSeconds(2);

    private static readonly IReadOnlyList<string> ReplaceProps =
        ["find", "replace", "regex", "matchCase", "wholeWord", "inFormulas"];

    private static void ApplyReplace(
        XLWorkbook workbook, EditOp op, int index, List<object> details, List<Warning> warnings)
    {
        var props = op.Props ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: replace needs props.",
            "Pass props like {\"find\":\"old\",\"replace\":\"new\"}.");

        foreach (var (key, _) in props)
        {
            if (!ReplaceProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: unknown replace prop '{key}'.",
                    "Supported replace props: " + string.Join(", ", ReplaceProps) + ".",
                    candidates: ReplaceProps);
            }
        }

        var find = props.TryGetPropertyValue("find", out var findNode) && findNode is not null
            ? StringPropValue(findNode, "find", index)
            : null;
        if (string.IsNullOrEmpty(find))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: replace needs a non-empty 'find' prop.",
                "Pass the text to search for, e.g. {\"find\":\"old\",\"replace\":\"new\"}.");
        }

        var replacement = props.TryGetPropertyValue("replace", out var replaceNode) && replaceNode is not null
            ? StringPropValue(replaceNode, "replace", index)
            : string.Empty;
        var useRegex = ReplaceFlag(props, "regex", index);
        var matchCase = ReplaceFlag(props, "matchCase", index);
        var wholeWord = ReplaceFlag(props, "wholeWord", index);
        var inFormulas = ReplaceFlag(props, "inFormulas", index);

        var finder = BuildFinder(find, useRegex, matchCase, wholeWord, index);

        // In literal mode '$' in the replacement is literal text, never
        // Regex substitution syntax.
        var substitution = useRegex
            ? replacement
            : replacement.Replace("$", "$$", StringComparison.Ordinal);

        var target = ExcelPaths.Resolve(workbook, op.Path);
        // Materialized: values are mutated while walking the used-cell set.
        IEnumerable<IXLCell> cells = target.Kind switch
        {
            ExcelTargetKind.Sheet => target.Sheet.CellsUsed().ToList(),
            ExcelTargetKind.Range => target.Range!.CellsUsed().ToList(),
            ExcelTargetKind.Cell => [target.Cell!],
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: replace targets a sheet, range or cell, not '{op.Path}'.",
                "Scope it like /Sheet1 (whole sheet) or /Sheet1/A1:C10; address a row as a range (A3:Z3)."),
        };

        var total = 0;
        var locations = new List<string>();
        foreach (var cell in cells)
        {
            string text;
            if (cell.HasFormula)
            {
                if (!inFormulas)
                {
                    continue; // formula results are derived; only inFormulas:true touches them
                }

                text = "=" + cell.FormulaA1; // the user-facing form get/read return
            }
            else if (cell.DataType == XLDataType.Text)
            {
                text = cell.GetText();
            }
            else
            {
                continue; // numbers/booleans/dates never match (display text is formatting)
            }

            var (count, replaced) = ReplaceText(finder, text, substitution, index);
            if (count == 0)
            {
                continue;
            }

            if (cell.HasFormula && replaced.StartsWith('=') && replaced.Length > 1)
            {
                cell.FormulaA1 = replaced; // re-evaluates in the cached-value save
            }
            else
            {
                cell.Value = replaced; // literal text (clears the formula if one was destroyed)
            }

            total += count;
            if (locations.Count < MaxReplaceLocations)
            {
                locations.Add(ExcelPaths.CellPath(cell.Worksheet, cell.Address));
            }
        }

        var canonical = DocPath.Parse(op.Path).ToCanonicalString();
        if (total == 0)
        {
            warnings.Add(new Warning(
                ErrorCodes.FindNoMatch,
                $"ops[{index}]: '{find}' matched nothing in {canonical}."));
        }

        details.Add(new { op = "replace", path = canonical, replacements = total, locations });
    }

    /// <summary>Compiles the search pattern (literal or regex) with the 2 s timeout.</summary>
    private static Regex BuildFinder(string find, bool useRegex, bool matchCase, bool wholeWord, int index)
    {
        var pattern = useRegex ? find : Regex.Escape(find);
        if (wholeWord)
        {
            // Lookarounds instead of \b so patterns that start/end with
            // non-word characters still honor the flag.
            pattern = @"(?<!\w)(?:" + pattern + @")(?!\w)";
        }

        var options = RegexOptions.CultureInvariant | (matchCase ? RegexOptions.None : RegexOptions.IgnoreCase);
        try
        {
            return new Regex(pattern, options, ReplaceTimeout);
        }
        catch (ArgumentException exception)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{find}' is not a valid .NET regular expression: {exception.Message}",
                "Fix the pattern, or drop regex:true to search for the literal text.",
                innerException: exception);
        }
    }

    private static (int Count, string Replaced) ReplaceText(Regex finder, string input, string substitution, int index)
    {
        try
        {
            var count = finder.Count(input);
            return count == 0 ? (0, input) : (count, finder.Replace(input, substitution));
        }
        catch (RegexMatchTimeoutException exception)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: the regex timed out after {ReplaceTimeout.TotalSeconds:0} seconds.",
                "Simplify the pattern (catastrophic backtracking?) or narrow the scope to a smaller range.",
                innerException: exception);
        }
    }

    private static bool ReplaceFlag(JsonObject props, string key, int index) =>
        props.TryGetPropertyValue(key, out var node) && node is not null && BoolPropValue(node, key, index);
}
