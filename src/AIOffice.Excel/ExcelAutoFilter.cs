using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// (1.12, additive) AutoFilter <em>criteria</em>: beyond the M3 bool form (which
/// only toggles the dropdowns), a sheet/table filter can now carry per-column
/// filters that actually HIDE the non-matching rows headlessly — a real applied
/// filter, so a headless reader sees the filtered view and Excel re-applies it on
/// open.
///
/// <para>Wire grammar (all additive; the bool form is unchanged):</para>
/// <list type="bullet">
/// <item><c>{autoFilter:true|false}</c> — enable/clear the filter (unchanged);</item>
/// <item><c>{autoFilter:{column, values:["East","West"]}}</c> — a VALUES filter: keep
///   only rows whose column equals one of <c>values</c>;</item>
/// <item><c>{autoFilter:{column, criteria:"&gt;100"|"&lt;=0"|"&lt;&gt;0"|"*text*"}}</c> — a
///   comparison/text filter;</item>
/// <item><c>{autoFilter:[{column, values|criteria}, …]}</c> — multiple columns (AND).</item>
/// </list>
///
/// <para><c>column</c> resolves against the filter range as a header name (the first
/// row of the range), a column letter, or a 1-based index within the range; a name
/// that does not resolve is <c>invalid_args</c> with the candidates.</para>
///
/// Applied through ClosedXML's <see cref="IXLAutoFilter"/> with <c>reapply</c>, so the
/// rows hide in the saved bytes and the criteria persist in the worksheet's
/// <c>&lt;autoFilter&gt;</c> element. <see cref="Read"/> parses that element back for
/// <c>get</c>.
/// </summary>
internal static partial class ExcelAutoFilter
{
    /// <summary>One validated per-column filter parsed from the wire.</summary>
    internal sealed record ColumnFilter(
        int IndexInRange, // 1-based position within the filter range
        string ColumnLabel, // the resolved column letter, for messages/readback
        IReadOnlyList<string>? Values, // a values filter, or null
        ComparisonFilter? Comparison); // a comparison/text filter, or null

    /// <summary>A parsed comparison/text criterion like <c>"&gt;100"</c> or <c>"*text*"</c>.</summary>
    internal sealed record ComparisonFilter(string Operator, string Operand);

    /// <summary>A leading comparison operator, longest first so <c>&gt;=</c> beats <c>&gt;</c>.</summary>
    private static readonly string[] Operators = [">=", "<=", "<>", "!=", "==", "=", ">", "<"];

    /// <summary>
    /// Applies the criteria form. <paramref name="node"/> is a JSON object (one
    /// column) or array (several). Each entry sets a real filter that hides the
    /// non-matching rows. Returns the applied-prop labels.
    /// </summary>
    public static IReadOnlyList<string> Apply(IXLWorksheet sheet, IXLRange range, JsonNode node, int index)
    {
        var entries = node switch
        {
            JsonArray array => array.Where(n => n is not null).Select(n => n!).ToList(),
            JsonObject obj => [obj],
            _ => throw Invalid(
                index,
                "autoFilter takes true/false, a {column, values|criteria} object, or an array of them.",
                "e.g. {op:set, path:/Sheet1/A1:D20, props:{autoFilter:{column:\"Region\", values:[\"East\",\"West\"]}}}."),
        };

        if (entries.Count == 0)
        {
            throw Invalid(
                index,
                "autoFilter array is empty.",
                "Pass at least one {column, values|criteria} entry, or {autoFilter:true} to just enable the dropdowns.");
        }

        var headers = ReadHeaders(range);
        var filters = entries.Select(e => Parse(e, range, headers, index)).ToList();

        // Enable the filter on the requested range first (idempotent), then apply
        // each column with reapply:true so ClosedXML hides the non-matching rows.
        var auto = range.SetAutoFilter();
        var applied = new List<string> { "autoFilter" };
        foreach (var filter in filters)
        {
            ApplyOne(auto, filter);
            applied.Add("autoFilter:" + filter.ColumnLabel);
        }

        return applied;
    }

    private static void ApplyOne(IXLAutoFilter auto, ColumnFilter filter)
    {
        var column = auto.Column(filter.IndexInRange);
        column.Clear(reapply: false); // start from a clean column (idempotent re-apply)

        if (filter.Values is { } values)
        {
            // A values filter: keep only the rows matching one of the values.
            for (var i = 0; i < values.Count; i++)
            {
                XLCellValue value = values[i];
                // reapply only on the last value so the row hide reflects the full set.
                column.AddFilter(value, reapply: i == values.Count - 1);
            }

            return;
        }

        var comparison = filter.Comparison!;
        ApplyComparison(column, comparison);
    }

    private static void ApplyComparison(IXLFilterColumn column, ComparisonFilter comparison)
    {
        var operand = comparison.Operand;

        // Wildcard text (* / ?) → contains/begins/ends, anchored on the wildcards.
        if (comparison.Operator is "=" or "==" && HasWildcard(operand))
        {
            ApplyWildcard(column, operand);
            return;
        }

        // Numeric operand when it parses as a number; else a string comparison.
        XLCellValue value = double.TryParse(operand, NumberStyles.Any, CultureInfo.InvariantCulture, out var number)
            ? number
            : operand;

        switch (comparison.Operator)
        {
            case ">": column.GreaterThan(value); break;
            case ">=": column.EqualOrGreaterThan(value); break;
            case "<": column.LessThan(value); break;
            case "<=": column.EqualOrLessThan(value); break;
            case "<>":
            case "!=": column.NotEqualTo(value); break;
            default: column.EqualTo(value); break; // "=" / "=="
        }
    }

    /// <summary>A wildcard text criterion maps to begins/ends/contains on the inner text.</summary>
    private static void ApplyWildcard(IXLFilterColumn column, string pattern)
    {
        var starts = pattern.StartsWith('*');
        var ends = pattern.EndsWith('*');
        var inner = pattern.Trim('*');
        if (starts && ends)
        {
            column.Contains(inner);
        }
        else if (ends)
        {
            column.BeginsWith(inner);
        }
        else if (starts)
        {
            column.EndsWith(inner);
        }
        else
        {
            column.EqualTo(pattern); // a bare '?' wildcard — fall back to equality
        }
    }

    private static bool HasWildcard(string text) =>
        text.Contains('*', StringComparison.Ordinal) || text.Contains('?', StringComparison.Ordinal);

    // ----- parsing ----------------------------------------------------------------

    private static ColumnFilter Parse(
        JsonNode node, IXLRange range, IReadOnlyList<string> headers, int index)
    {
        if (node is not JsonObject obj)
        {
            throw Invalid(
                index,
                "each autoFilter entry is a {column, values|criteria} object.",
                "e.g. {column:\"Region\", values:[\"East\"]} or {column:\"Amount\", criteria:\">100\"}.");
        }

        if (!obj.TryGetPropertyValue("column", out var columnNode) || columnNode is null)
        {
            throw Invalid(
                index,
                "an autoFilter entry needs a 'column'.",
                "Identify the column by header name, column letter, or 1-based index within the filter range.");
        }

        var (indexInRange, label) = ResolveColumn(columnNode, range, headers, index);

        var hasValues = obj.TryGetPropertyValue("values", out var valuesNode) && valuesNode is not null;
        var hasCriteria = obj.TryGetPropertyValue("criteria", out var criteriaNode) && criteriaNode is not null;
        if (hasValues == hasCriteria)
        {
            throw Invalid(
                index,
                "an autoFilter entry needs exactly one of 'values' or 'criteria'.",
                "Use {column, values:[\"East\",\"West\"]} for a list filter, or {column, criteria:\">100\"} for a comparison.");
        }

        if (hasValues)
        {
            var values = ParseValues(valuesNode!, index);
            return new ColumnFilter(indexInRange, label, values, null);
        }

        var comparison = ParseCriteria(criteriaNode!, index);
        return new ColumnFilter(indexInRange, label, null, comparison);
    }

    private static IReadOnlyList<string> ParseValues(JsonNode node, int index)
    {
        if (node is not JsonArray array || array.Count == 0)
        {
            throw Invalid(
                index,
                "autoFilter 'values' must be a non-empty array.",
                "Pass the values to keep, e.g. {values:[\"East\",\"West\"]}.");
        }

        var values = new List<string>(array.Count);
        foreach (var item in array)
        {
            values.Add(item switch
            {
                JsonValue v when v.GetValueKind() == JsonValueKind.String => v.GetValue<string>(),
                JsonValue v when v.GetValueKind() == JsonValueKind.Number => v.ToJsonString(),
                JsonValue v when v.GetValueKind() is JsonValueKind.True or JsonValueKind.False =>
                    v.GetValue<bool>() ? "TRUE" : "FALSE",
                _ => throw Invalid(
                    index,
                    "autoFilter 'values' entries must be strings, numbers, or booleans.",
                    "e.g. {values:[\"East\",\"West\"]} or {values:[10,20]}."),
            });
        }

        return values;
    }

    private static ComparisonFilter ParseCriteria(JsonNode node, int index)
    {
        if (node is not JsonValue value || value.GetValueKind() != JsonValueKind.String)
        {
            throw Invalid(
                index,
                "autoFilter 'criteria' must be a string.",
                "e.g. {criteria:\">100\"}, {criteria:\"<=0\"}, {criteria:\"<>0\"}, or {criteria:\"*text*\"}.");
        }

        var text = value.GetValue<string>().Trim();
        if (text.Length == 0)
        {
            throw Invalid(
                index,
                "autoFilter 'criteria' is empty.",
                "Pass a comparison like \">100\" or a wildcard like \"*east*\".");
        }

        foreach (var op in Operators)
        {
            if (text.StartsWith(op, StringComparison.Ordinal))
            {
                var operand = text[op.Length..].Trim();
                if (operand.Length == 0)
                {
                    throw Invalid(
                        index,
                        $"autoFilter criteria '{text}' has the operator '{op}' but no value.",
                        "e.g. \">100\" keeps values above 100.");
                }

                return new ComparisonFilter(op, operand);
            }
        }

        // No leading operator: a bare value is equality (wildcards make it text).
        return new ComparisonFilter("=", text);
    }

    // ----- column resolution ------------------------------------------------------

    /// <summary>The header texts of the filter range's first row (used to resolve names).</summary>
    private static IReadOnlyList<string> ReadHeaders(IXLRange range)
    {
        var headers = new List<string>(range.ColumnCount());
        var headerRow = range.FirstRow();
        for (var i = 1; i <= range.ColumnCount(); i++)
        {
            headers.Add(headerRow.Cell(i).GetFormattedString());
        }

        return headers;
    }

    /// <summary>
    /// Resolves the wire <c>column</c> (header name | column letter | 1-based index)
    /// to a 1-based position WITHIN the filter range. Out-of-range / unknown names
    /// are <c>invalid_args</c> with the available headers/letters as candidates.
    /// </summary>
    private static (int IndexInRange, string Label) ResolveColumn(
        JsonNode columnNode, IXLRange range, IReadOnlyList<string> headers, int index)
    {
        var columnCount = range.ColumnCount();
        var firstColumn = range.RangeAddress.FirstAddress.ColumnNumber;

        if (columnNode is JsonValue value && value.GetValueKind() == JsonValueKind.Number &&
            value.TryGetValue<int>(out var number))
        {
            if (number < 1 || number > columnCount)
            {
                throw OutOfRange(index, number.ToString(CultureInfo.InvariantCulture), columnCount, headers);
            }

            return (number, ExcelCharts.ColumnLetters(firstColumn + number - 1));
        }

        var text = (columnNode as JsonValue)?.GetValueKind() == JsonValueKind.String
            ? columnNode.GetValue<string>().Trim()
            : throw Invalid(
                index,
                "autoFilter 'column' must be a header name, a column letter, or a 1-based index.",
                "e.g. {column:\"Region\"}, {column:\"A\"}, or {column:1}.");

        // 1) header name (case-insensitive).
        for (var i = 0; i < headers.Count; i++)
        {
            if (string.Equals(headers[i], text, StringComparison.OrdinalIgnoreCase))
            {
                return (i + 1, ExcelCharts.ColumnLetters(firstColumn + i));
            }
        }

        // 2) a 1-based index written as a string.
        if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var asNumber))
        {
            if (asNumber < 1 || asNumber > columnCount)
            {
                throw OutOfRange(index, text, columnCount, headers);
            }

            return (asNumber, ExcelCharts.ColumnLetters(firstColumn + asNumber - 1));
        }

        // 3) a column letter — must land inside the filter range.
        if (ColumnLetterPattern().IsMatch(text))
        {
            var absolute = new CellRef(text.ToUpperInvariant(), 1).ColumnNumber;
            var inRange = absolute - firstColumn + 1;
            if (inRange < 1 || inRange > columnCount)
            {
                throw OutOfRange(index, text, columnCount, headers);
            }

            return (inRange, text.ToUpperInvariant());
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: autoFilter column '{text}' is not a header, column letter, or index in {range.RangeAddress}.",
            "Use a header name from the first row, a column letter inside the range, or a 1-based index.",
            candidates: Candidates(headers));
    }

    private static AiofficeException OutOfRange(
        int index, string column, int columnCount, IReadOnlyList<string> headers) =>
        new(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: autoFilter column '{column}' is outside the filter range (1..{columnCount}).",
            $"The filter range has {columnCount} column(s); pass a 1-based index within it.",
            candidates: Candidates(headers));

    private static IReadOnlyList<string> Candidates(IReadOnlyList<string> headers) =>
        headers.Where(h => h.Length > 0).ToList();

    private static AiofficeException Invalid(int index, string message, string suggestion) =>
        new(ErrorCodes.InvalidArgs, $"ops[{index}]: {message}", suggestion);

    [GeneratedRegex("^[A-Za-z]{1,3}$")]
    private static partial Regex ColumnLetterPattern();

    // ----- readback (raw <autoFilter> XML) ----------------------------------------

    /// <summary>
    /// Reads the active per-column filters back for <c>get</c> from the worksheet's
    /// <c>&lt;autoFilter&gt;</c> element (ClosedXML's public surface exposes the type
    /// but not the values, so aioffice parses the raw element). Returns null when no
    /// column carries a criterion (a plain enabled filter reports just its range).
    /// </summary>
    public static IReadOnlyList<object>? Read(string file, string sheetName)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart;
        var sheetElement = workbookPart?.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheetElement?.Id?.Value is not { } relId ||
            workbookPart!.GetPartById(relId) is not WorksheetPart worksheetPart ||
            worksheetPart.Worksheet?.Elements<S.AutoFilter>().FirstOrDefault() is not { } autoFilter)
        {
            return null;
        }

        var reference = autoFilter.Reference?.Value;
        var firstColumn = FirstColumnOf(reference);
        var columns = new List<object>();
        foreach (var filterColumn in autoFilter.Elements<S.FilterColumn>())
        {
            var colId = (int?)(filterColumn.ColumnId?.Value) ?? 0;
            var label = firstColumn > 0 ? ExcelCharts.ColumnLetters(firstColumn + (int)colId) : null;

            if (filterColumn.Elements<S.Filters>().FirstOrDefault() is { } filters)
            {
                var values = filters.Elements<S.Filter>()
                    .Select(f => f.Val?.Value)
                    .Where(v => v is not null)
                    .Cast<string>()
                    .ToList();
                columns.Add(new { column = label, indexInRange = (int)colId + 1, kind = "values", values });
            }
            else if (filterColumn.Elements<S.CustomFilters>().FirstOrDefault() is { } custom)
            {
                var criteria = custom.Elements<S.CustomFilter>()
                    .Select(c => new
                    {
                        // The OOXML wire operator name (e.g. "greaterThan"); the SDK's
                        // struct-enum ToString is opaque, so read InnerText. Absent
                        // attribute defaults to "equal" per the schema.
                        @operator = string.IsNullOrEmpty(c.Operator?.InnerText) ? "equal" : c.Operator!.InnerText,
                        value = c.Val?.Value,
                    })
                    .ToList<object>();
                columns.Add(new { column = label, indexInRange = (int)colId + 1, kind = "custom", criteria });
            }
        }

        return columns.Count > 0 ? columns : null;
    }

    /// <summary>The 1-based absolute first column number of a filter ref like "A1:D20".</summary>
    private static int FirstColumnOf(string? reference)
    {
        if (string.IsNullOrEmpty(reference))
        {
            return 0;
        }

        var firstCell = reference.Split(':')[0];
        var letters = new string(firstCell.TakeWhile(char.IsLetter).ToArray());
        return letters.Length is >= 1 and <= 3 ? new CellRef(letters.ToUpperInvariant(), 1).ColumnNumber : 0;
    }
}
