using System.Globalization;
using System.Text;

namespace AIOffice.Excel;

/// <summary>
/// The v1.2 named-preset library for the <c>numberFormat</c> prop (additive,
/// contract-frozen surface unchanged). The existing prop already accepts any
/// raw OOXML format code (e.g. <c>"$#,##0.00"</c>); this layer lets it ALSO
/// accept a short stable name (e.g. <c>"currency-usd"</c>) that resolves to its
/// format code. Resolution rules — the whole point is to never break the
/// existing literal-code behavior:
///
/// <list type="bullet">
/// <item>A name in <see cref="Presets"/> resolves to its OOXML format code.</item>
/// <item>Anything else is treated as a LITERAL custom format string and passes
/// through unchanged (the v1.0/v1.1 behavior — a real OOXML code like
/// <c>0.00%</c> or <c>#,##0</c> is never a preset name and is preserved
/// verbatim).</item>
/// </list>
///
/// Because resolution is a pure name→code lookup with literal fallthrough, every
/// number-format string that worked before still works, and only the new short
/// names gain meaning. A misspelled preset is left as a literal (Excel shows it
/// as a custom format) but <see cref="NearestPreset"/> can surface the closest
/// real preset in a suggestion when a caller wants to be helpful.
/// </summary>
internal static class ExcelNumberFormats
{
    /// <summary>
    /// Preset name → OOXML format code. Names are matched case-insensitively.
    /// The codes are the conventional Excel custom-format strings each preset
    /// stands for (locale-neutral: a literal currency symbol, not a locale code,
    /// so the file shows the same on every machine).
    /// </summary>
    public static readonly IReadOnlyDictionary<string, string> Presets =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Money — accounting aligns the symbol and parenthesizes negatives.
            ["accounting-usd"] = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
            ["currency-usd"] = "\"$\"#,##0.00",
            ["currency-eur"] = "\"€\"#,##0.00",
            ["currency-gbp"] = "\"£\"#,##0.00",
            ["currency-jpy"] = "\"¥\"#,##0",

            // Percentages.
            ["percent"] = "0%",
            ["percent2"] = "0.00%",

            // Scientific / fraction.
            ["scientific"] = "0.00E+00",
            ["fraction"] = "# ?/?",

            // Plain numbers.
            ["thousands"] = "#,##0",
            ["thousands2"] = "#,##0.00",
            ["integer"] = "0",
            ["number2"] = "0.00",

            // Dates & times (ISO-leaning, locale-neutral codes).
            ["date-iso"] = "yyyy-mm-dd",
            ["datetime-iso"] = "yyyy-mm-dd hh:mm:ss",
            ["time"] = "hh:mm:ss",
            ["duration"] = "[h]:mm:ss",

            // Text — forces a cell to display its content as entered.
            ["text"] = "@",
        };

    /// <summary>
    /// Resolves a <c>numberFormat</c> prop value to the OOXML format code to
    /// store. A known preset name maps to its code; any other string is a
    /// literal custom format and passes through unchanged.
    /// </summary>
    public static string Resolve(string value) =>
        Presets.TryGetValue(value, out var code) ? code : value;

    /// <summary>True when the value is a registered preset name (vs. a literal code).</summary>
    public static bool IsPreset(string value) => Presets.ContainsKey(value);

    /// <summary>
    /// The closest preset name to a (presumably mistyped) value within a small
    /// edit distance, or null when nothing is near. Used only to enrich
    /// suggestions; resolution itself never depends on it.
    /// </summary>
    public static string? NearestPreset(string value)
    {
        // Only bother for short, name-like tokens — a real format code like
        // "#,##0.00" or "yyyy-mm-dd" is never a near-miss of a preset name.
        if (value.Length == 0 || value.Length > 24 || !LooksLikeName(value))
        {
            return null;
        }

        string? best = null;
        var bestDistance = int.MaxValue;
        foreach (var name in Presets.Keys)
        {
            var distance = ExcelPaths.Levenshtein(value, name);
            if (distance < bestDistance)
            {
                bestDistance = distance;
                best = name;
            }
        }

        // A typo is at most ~a third of the name away; beyond that it is almost
        // certainly an intentional literal format, not a misspelled preset.
        var threshold = Math.Max(1, value.Length / 3);
        return bestDistance <= threshold ? best : null;
    }

    /// <summary>A token that could plausibly be a preset name (letters, digits, hyphens only).</summary>
    private static bool LooksLikeName(string value) =>
        value.All(c => char.IsLetterOrDigit(c) || c == '-');

    /// <summary>
    /// The human-readable preset table for <c>aioffice help number-formats</c>:
    /// each preset name with the OOXML code it resolves to, plus a note that
    /// any other string is a literal custom format. Deterministic ordering.
    /// </summary>
    public static string HelpText()
    {
        var sb = new StringBuilder();
        sb.Append("# Number-format presets\n\n");
        sb.Append("The `numberFormat` prop accepts either a raw OOXML format code (e.g. `\"$#,##0.00\"`) ");
        sb.Append("or one of the named presets below. A preset resolves to its format code; ");
        sb.Append("any other string is kept verbatim as a literal custom format.\n\n");
        sb.Append("| preset | format code |\n");
        sb.Append("|--------|-------------|\n");
        foreach (var (name, code) in Presets.OrderBy(p => p.Key, StringComparer.Ordinal))
        {
            sb.Append(string.Create(CultureInfo.InvariantCulture, $"| `{name}` | `{code}` |\n"));
        }

        sb.Append("\nExamples:\n");
        sb.Append("  aioffice edit book.xlsx --ops '[{\"op\":\"set\",\"path\":\"/Sheet1/B2\",");
        sb.Append("\"props\":{\"numberFormat\":\"currency-usd\"}}]'\n");
        sb.Append("  aioffice edit book.xlsx --ops '[{\"op\":\"set\",\"path\":\"/Sheet1/C2\",");
        sb.Append("\"props\":{\"numberFormat\":\"0.00%\"}}]'  # a literal code still works\n");
        return sb.ToString();
    }

    /// <summary>The preset names in stable order, for schema / candidate lists.</summary>
    public static IReadOnlyList<string> Names => [.. Presets.Keys.OrderBy(k => k, StringComparer.Ordinal)];
}
