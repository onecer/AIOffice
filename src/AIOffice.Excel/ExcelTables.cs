using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// Excel structured tables (ListObjects) — the M6 ListObject feature, all
/// ClosedXML-native (<c>IXLRange.CreateTable</c> / <c>IXLTable</c>).
///
/// <para>Add: <c>{op:add, path:/Sheet1/A1:D20, type:table,
/// props:{name:"Sales", style:"medium2", headerRow:true, totalsRow:false,
/// bandedRows:true}}</c> turns a range (first row = headers) into a real
/// structured table with a built-in table style. <c>totals</c> maps column
/// names to a totals function (<c>{"Sales":"sum"}</c>), which also turns the
/// totals row on. Structured references in cell formulas
/// (<c>=SUM(Sales[Sales])</c>) evaluate because the table exists in the model
/// before the save pass runs.</para>
///
/// <para>Get: <c>/Sheet1/table[@name=Sales]</c> describes the table (range,
/// style, columns, header/totals flags, banding). Remove: drops the
/// ListObject but KEEPS the cell data (ClosedXML's <c>Clear(...)</c> would
/// erase values, so the table is unregistered while its cells survive).</para>
///
/// <para>Honest capability notes:</para>
/// <list type="bullet">
/// <item>Styles are the built-in Excel gallery (light/medium/dark 1..N). The
/// short forms <c>"medium2"</c>, <c>"light1"</c>, <c>"none"</c> map onto
/// <see cref="XLTableTheme"/>; an unknown style returns invalid_args listing
/// the accepted forms.</item>
/// <item>Totals functions are sum/average/count/countNumbers/min/max/
/// stdDev/var (Excel's built-in set); an unknown function returns invalid_args.
/// A totals column that is not in the table returns invalid_args naming the
/// real columns.</item>
/// </list>
/// </summary>
internal static class ExcelTables
{
    /// <summary>The totals-row functions Excel exposes in the dropdown, by wire name.</summary>
    private static readonly IReadOnlyDictionary<string, XLTotalsRowFunction> TotalsFunctions =
        new Dictionary<string, XLTotalsRowFunction>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = XLTotalsRowFunction.None,
            ["sum"] = XLTotalsRowFunction.Sum,
            ["average"] = XLTotalsRowFunction.Average,
            ["count"] = XLTotalsRowFunction.Count,
            ["countNumbers"] = XLTotalsRowFunction.CountNumbers,
            ["min"] = XLTotalsRowFunction.Minimum,
            ["max"] = XLTotalsRowFunction.Maximum,
            ["stdDev"] = XLTotalsRowFunction.StandardDeviation,
            ["var"] = XLTotalsRowFunction.Variance,
        };

    public static IReadOnlyList<string> TotalsFunctionNames => [.. TotalsFunctions.Keys];

    // ----- add ----------------------------------------------------------------

    public static object Add(ExcelTarget target, EditOp op, int index)
    {
        if (target.Kind != ExcelTargetKind.Range)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add table needs a range path like /Sheet1/A1:C10 (first row = headers).",
                "Address the data including its header row.");
        }

        var props = op.Props;
        var name = StringProp(props, "name");
        var headerRow = BoolProp(props, "headerRow") ?? true;

        IXLTable table;
        try
        {
            table = name is null
                ? target.Range!.CreateTable()
                : target.Range!.CreateTable(name);
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: could not create the table: {exception.Message}",
                "Table names must be unique in the workbook and the range must not overlap an existing table.",
                innerException: exception);
        }

        if (!headerRow)
        {
            table.SetShowHeaderRow(false);
        }

        if (StyleProp(props, index) is { } theme)
        {
            table.Theme = theme;
        }

        if (BoolProp(props, "bandedRows") is { } banded)
        {
            table.SetShowRowStripes(banded);
        }

        if (BoolProp(props, "bandedColumns") is { } bandedColumns)
        {
            table.SetShowColumnStripes(bandedColumns);
        }

        var totals = ApplyTotals(table, props, index);

        return new
        {
            op = "add",
            type = "table",
            path = ExcelPaths.TablePath(target.Sheet, table.Name),
            name = table.Name,
            range = table.RangeAddress.ToString(),
            style = ThemeName(table.Theme),
            columns = table.Fields.Select(f => f.Name).ToList(),
            headerRow = table.ShowHeaderRow,
            totalsRow = table.ShowTotalsRow,
            bandedRows = table.Theme != XLTableTheme.None && totals.ShowRowStripes,
            totals = totals.Applied.Count > 0 ? totals.Applied : null,
        };
    }

    private readonly record struct TotalsResult(Dictionary<string, string> Applied, bool ShowRowStripes);

    /// <summary>
    /// Applies <c>totalsRow</c> (turn the row on/off) and a <c>totals</c> map of
    /// column name → function. Setting any totals function turns the totals row
    /// on. Returns the applied functions for the response.
    /// </summary>
    private static TotalsResult ApplyTotals(IXLTable table, JsonObject? props, int index)
    {
        var applied = new Dictionary<string, string>(StringComparer.Ordinal);
        var totalsRow = BoolProp(props, "totalsRow");
        var totalsNode = props?["totals"] as JsonObject;

        if (totalsNode is not null && totalsNode.Count > 0)
        {
            table.SetShowTotalsRow(true);
            foreach (var (column, functionNode) in totalsNode)
            {
                var field = FieldOrThrow(table, column, index);
                var functionName = functionNode?.GetValue<string>()
                    ?? throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{index}]: totals['{column}'] needs a function name.",
                        "Use one of: " + string.Join(", ", TotalsFunctionNames) + ".");
                if (!TotalsFunctions.TryGetValue(functionName, out var function))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{index}]: unknown totals function '{functionName}' for column '{column}'.",
                        "Use one of: " + string.Join(", ", TotalsFunctionNames) + ".",
                        candidates: TotalsFunctionNames);
                }

                field.TotalsRowFunction = function;
                applied[field.Name] = functionName;
            }
        }
        else if (totalsRow == true)
        {
            table.SetShowTotalsRow(true);
        }
        else if (totalsRow == false)
        {
            table.SetShowTotalsRow(false);
        }

        return new TotalsResult(applied, table.Theme.ToString().Length > 0);
    }

    // ----- set totals (post-creation edit) ------------------------------------

    /// <summary>
    /// Edits a table's totals-row functions and labels after creation. The
    /// <paramref name="totals"/> object maps a column name to
    /// <c>{ function?, label? }</c>. <c>function:"none"</c> clears the column's
    /// function; an empty-string label clears the label. Setting any totals turns
    /// the totals row on (parity with how Excel auto-shows the row). Returns the
    /// applied edits per column for the response.
    ///
    /// <para>A totals cell holds EITHER a built-in function OR a custom label,
    /// never both (Excel's own model): setting a non-empty label clears the
    /// column's function, and setting a function clears a custom label. This keeps
    /// the column's totals state unambiguous and round-trips cleanly through the
    /// OOXML reader.</para>
    ///
    /// <para>After setting a function the totals cell's SUBTOTAL formula is
    /// materialized (<c>TotalsCell.FormulaA1</c>): ClosedXML's
    /// evaluate-before-save pass otherwise strips a loaded table's
    /// <c>totalsRowFunction</c>, and touching the formula makes it survive — the
    /// same shape the table-add path produces.</para>
    /// </summary>
    public static object ApplySetTotals(ExcelTarget target, IXLTable table, JsonObject totals, int index)
    {
        if (totals.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: totals needs at least one column → {{function?, label?}} entry.",
                "Use {totals:{\"Qty\":{function:\"sum\", label:\"Total Qty\"}}}.");
        }

        table.SetShowTotalsRow(true);

        var applied = new Dictionary<string, object>(StringComparer.Ordinal);
        foreach (var (column, settingNode) in totals)
        {
            if (settingNode is not JsonObject setting)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: totals['{column}'] needs an object like {{function:\"sum\", label:\"Total\"}}.",
                    "Map each column to a {function?, label?} object.");
            }

            var field = FieldOrThrow(table, column, index);
            var entry = new Dictionary<string, string?>(StringComparer.Ordinal);
            var hasFunction = setting.TryGetPropertyValue("function", out var functionNode) && functionNode is not null;
            var hasLabel = setting.TryGetPropertyValue("label", out var labelNode) && labelNode is not null;

            if (!hasFunction && !hasLabel)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: totals['{column}'] needs a function or label.",
                    "Provide {function:\"sum\"} or {label:\"Total\"}.");
            }

            // Resolve the column's final state first. A totals cell holds EITHER a
            // function OR a custom label, never both — when this op sets both, the
            // label wins (the OOXML reader gives a label precedence). Start from the
            // column's current state so function-only / label-only edits leave the
            // other side untouched.
            XLTotalsRowFunction function = field.TotalsRowFunction;
            string? label = string.IsNullOrEmpty(field.TotalsRowLabel) ? null : field.TotalsRowLabel;

            if (hasFunction)
            {
                if (functionNode is not JsonValue functionValue || functionValue.GetValueKind() != JsonValueKind.String)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{index}]: totals['{column}'].function must be a string.",
                        "Use one of: " + string.Join(", ", TotalsFunctionNames) + ".",
                        candidates: TotalsFunctionNames);
                }

                var functionName = functionNode!.GetValue<string>();
                if (!TotalsFunctions.TryGetValue(functionName, out var resolved))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{index}]: unknown totals function '{functionName}' for column '{column}'.",
                        "Use one of: " + string.Join(", ", TotalsFunctionNames) + ".",
                        candidates: TotalsFunctionNames);
                }

                function = resolved;
                label = null; // a function owns the cell; drop a stray label
            }

            if (hasLabel)
            {
                if (labelNode is not JsonValue labelValue || labelValue.GetValueKind() != JsonValueKind.String)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{index}]: totals['{column}'].label must be a string.",
                        "Use a string label, or \"\" to clear it.");
                }

                var requested = labelNode!.GetValue<string>();
                if (requested.Length == 0)
                {
                    label = null; // empty string clears the label
                }
                else
                {
                    function = XLTotalsRowFunction.None; // a custom label replaces a function
                    label = requested;
                }
            }

            // Apply in an order that survives ClosedXML's writer: clear the label
            // BEFORE setting the function (setting a null label after a function
            // resets the function in the model), then materialize the SUBTOTAL so
            // the evaluate-before-save pass keeps it (matches the table-add path).
            field.TotalsRowLabel = label;
            field.TotalsRowFunction = function;
            if (function != XLTotalsRowFunction.None)
            {
                _ = field.TotalsCell.FormulaA1;
                entry["function"] = TotalsName(function);
            }

            if (label is not null)
            {
                entry["label"] = label;
            }

            applied[field.Name] = entry;
        }

        return new
        {
            op = "set",
            type = "table",
            path = ExcelPaths.TablePath(target.Sheet, table.Name),
            name = table.Name,
            totalsRow = table.ShowTotalsRow,
            totals = applied,
        };
    }

    private static IXLTableField FieldOrThrow(IXLTable table, string columnName, int index)
    {
        var field = table.Fields.FirstOrDefault(f =>
            string.Equals(f.Name, columnName, StringComparison.OrdinalIgnoreCase));
        if (field is not null)
        {
            return field;
        }

        var columns = table.Fields.Select(f => f.Name).ToList();
        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: table '{table.Name}' has no column '{columnName}'.",
            "Totals are keyed by column name; available columns are listed as candidates.",
            candidates: columns);
    }

    // ----- get ----------------------------------------------------------------

    public static IXLTable Find(ExcelTarget target)
    {
        var table = target.Sheet.Tables.FirstOrDefault(t =>
            string.Equals(t.Name, target.TableName, StringComparison.OrdinalIgnoreCase));
        if (table is not null)
        {
            return table;
        }

        var candidates = target.Sheet.Tables
            .Select(t => ExcelPaths.TablePath(target.Sheet, t.Name))
            .ToList();
        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No table named '{target.TableName}' on sheet '{target.Sheet.Name}'.",
            candidates.Count > 0
                ? "Table names are matched case-insensitively; pick one of the candidates."
                : "This sheet has no tables; add one with {op:add, type:table, path:" +
                  ExcelPaths.SheetPath(target.Sheet) + "/A1:D20, props:{name:\"Sales\"}}.",
            candidates: candidates.Count > 0 ? candidates : [ExcelPaths.SheetPath(target.Sheet)]);
    }

    public static object Describe(IXLWorksheet sheet, IXLTable table) => new
    {
        path = ExcelPaths.TablePath(sheet, table.Name),
        kind = "table",
        sheet = sheet.Name,
        name = table.Name,
        range = table.RangeAddress.ToString(),
        style = ThemeName(table.Theme),
        headerRow = table.ShowHeaderRow,
        totalsRow = table.ShowTotalsRow,
        bandedRows = table.ShowRowStripes,
        bandedColumns = table.ShowColumnStripes,
        columns = table.Fields.Select(f => new
        {
            name = f.Name,
            totalsFunction = TotalsName(f.TotalsRowFunction),
            totalsLabel = string.IsNullOrEmpty(f.TotalsRowLabel) ? null : f.TotalsRowLabel,
        }).ToList(),
    };

    // ----- remove (keep the data) ---------------------------------------------

    /// <summary>
    /// Drops the ListObject registration but keeps the cell data: ClosedXML's
    /// <c>IXLTables.Remove</c> unregisters the table without clearing cells.
    /// </summary>
    public static object Remove(ExcelTarget target)
    {
        var table = Find(target);
        var name = table.Name;
        var range = table.RangeAddress.ToString();
        target.Sheet.Tables.Remove(name);
        return new
        {
            op = "remove",
            path = ExcelPaths.TablePath(target.Sheet, name),
            removed = "table",
            name,
            note = "data kept",
            range,
        };
    }

    // ----- style mapping ------------------------------------------------------

    /// <summary>
    /// Maps a short style name to a built-in table theme. Accepts <c>none</c>,
    /// <c>light1..21</c>, <c>medium1..28</c>, <c>dark1..11</c>, and the full
    /// <c>TableStyleMedium2</c> form. Null when no style prop was supplied.
    /// </summary>
    private static XLTableTheme? StyleProp(JsonObject? props, int index)
    {
        var style = StringProp(props, "style");
        if (style is null)
        {
            return null;
        }

        if (string.Equals(style, "none", StringComparison.OrdinalIgnoreCase))
        {
            return XLTableTheme.None;
        }

        var canonical = CanonicalThemeName(style);
        try
        {
            var theme = XLTableTheme.FromName(canonical);
            if (theme is not null)
            {
                return theme;
            }
        }
        catch (ArgumentException)
        {
            // fall through to the typed error below
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: unknown table style '{style}'.",
            "Use none, light1-21, medium1-28, dark1-11 (e.g. \"medium2\"), or the full name like \"TableStyleMedium2\".",
            candidates: ["none", "light1", "medium2", "medium9", "dark1"]);
    }

    /// <summary>"medium2" → "TableStyleMedium2"; a full name passes through.</summary>
    private static string CanonicalThemeName(string style)
    {
        if (style.StartsWith("TableStyle", StringComparison.OrdinalIgnoreCase))
        {
            return style;
        }

        foreach (var tier in new[] { "Light", "Medium", "Dark" })
        {
            if (style.StartsWith(tier, StringComparison.OrdinalIgnoreCase) &&
                style.Length > tier.Length &&
                int.TryParse(style.AsSpan(tier.Length), NumberStyles.None, CultureInfo.InvariantCulture, out var n))
            {
                return string.Create(CultureInfo.InvariantCulture, $"TableStyle{tier}{n}");
            }
        }

        return style; // let FromName reject it
    }

    private static string ThemeName(XLTableTheme theme) =>
        theme == XLTableTheme.None ? "none" : theme.Name;

    private static string? TotalsName(XLTotalsRowFunction function) => function switch
    {
        XLTotalsRowFunction.None => null,
        XLTotalsRowFunction.Sum => "sum",
        XLTotalsRowFunction.Average => "average",
        XLTotalsRowFunction.Count => "count",
        XLTotalsRowFunction.CountNumbers => "countNumbers",
        XLTotalsRowFunction.Minimum => "min",
        XLTotalsRowFunction.Maximum => "max",
        XLTotalsRowFunction.StandardDeviation => "stdDev",
        XLTotalsRowFunction.Variance => "var",
        XLTotalsRowFunction.Custom => "custom",
        _ => function.ToString().ToLowerInvariant(),
    };

    // ----- prop helpers -------------------------------------------------------

    private static string? StringProp(JsonObject? props, string key) =>
        props?.TryGetPropertyValue(key, out var node) == true && node is JsonValue value &&
        value.GetValueKind() == System.Text.Json.JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static bool? BoolProp(JsonObject? props, string key) =>
        props?.TryGetPropertyValue(key, out var node) == true && node is JsonValue value &&
        value.GetValueKind() is System.Text.Json.JsonValueKind.True or System.Text.Json.JsonValueKind.False
            ? value.GetValue<bool>()
            : null;
}
