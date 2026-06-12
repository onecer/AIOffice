using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The xlsx data-validation layer (M5, ClosedXML-native) — how agents build
/// forms and dropdowns. Supported kinds: <c>list</c> (literal values or a
/// <c>sourceRange</c>), <c>wholeNumber</c>, <c>decimal</c>, <c>date</c>,
/// <c>textLength</c>; anything else is a typed <c>unsupported_feature</c>.
/// Rules are addressed as <c>/Sheet1/dataValidation[i]</c> (1-based, sheet
/// order); <c>read --view structure</c> lists them, <c>get</c> describes one,
/// <c>remove</c> deletes one. Like conditional formats, there is no in-place
/// <c>set</c> — remove and re-add.
/// </summary>
internal static class ExcelDataValidations
{
    /// <summary>The validation kinds aioffice can create.</summary>
    public static readonly IReadOnlyList<string> Kinds = ["list", "wholeNumber", "decimal", "date", "textLength"];

    private static readonly IReadOnlyList<string> Operators =
        ["between", "notBetween", ">", "<", ">=", "<=", "==", "!="];

    private static readonly IReadOnlyList<string> CommonProps =
        ["kind", "allowBlank", "inputTitle", "inputMessage", "errorTitle", "errorMessage", "errorStyle"];

    private static readonly IReadOnlyList<string> ListProps = [.. CommonProps, "values", "sourceRange"];

    private static readonly IReadOnlyList<string> ComparisonProps = [.. CommonProps, "operator", "value", "value2"];

    private static readonly IReadOnlyList<string> ErrorStyles = ["stop", "warning", "information"];

    // ----- add ---------------------------------------------------------------

    /// <summary>Validates and applies an <c>add dataValidation</c> op; returns the details entry.</summary>
    public static object Add(XLWorkbook workbook, ExcelTarget target, EditOp op, int opIndex)
    {
        var range = target.Kind switch
        {
            ExcelTargetKind.Range => target.Range!,
            ExcelTargetKind.Cell => target.Cell!.AsRange(),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add dataValidation targets a cell or range path like /Sheet1/B2:B50.",
                "Address the cells to validate, e.g. {op:add, type:dataValidation, path:/Sheet1/B2:B50, " +
                "props:{kind:\"list\", values:[\"todo\",\"doing\",\"done\"]}}."),
        };

        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add dataValidation needs props.",
                "Pass props like {\"kind\":\"list\",\"values\":[\"todo\",\"doing\",\"done\"]} or " +
                "{\"kind\":\"wholeNumber\",\"operator\":\"between\",\"value\":1,\"value2\":10}.");
        }

        var kind = OptionalString(props, "kind") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add dataValidation needs a 'kind'.",
            "Supported kinds: " + string.Join(", ", Kinds) + ".",
            candidates: Kinds);

        var validation = range.CreateDataValidation();
        switch (kind)
        {
            case "list":
                GuardProps(props, ListProps, opIndex);
                AddList(workbook, target.Sheet, validation, props, opIndex);
                break;
            case "wholeNumber":
            case "decimal":
            case "date":
            case "textLength":
                GuardProps(props, ComparisonProps, opIndex);
                AddComparison(validation, kind, props, opIndex);
                break;
            default:
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"ops[{opIndex}]: dataValidation kind '{kind}' is not supported yet.",
                    "Supported kinds: " + string.Join(", ", Kinds) + ". For time or custom-formula rules, " +
                    "a decimal or textLength rule is the usual stand-in.",
                    candidates: Kinds);
        }

        ApplyCommonProps(validation, props, opIndex);

        var index = target.Sheet.DataValidations.Count();
        return new
        {
            op = "add",
            type = "dataValidation",
            path = ExcelPaths.DataValidationPath(target.Sheet, index),
            dvKind = kind,
            range = range.RangeAddress.ToString(),
        };
    }

    private static void AddList(
        XLWorkbook workbook, IXLWorksheet sheet, IXLDataValidation validation, JsonObject props, int opIndex)
    {
        var hasValues = props.TryGetPropertyValue("values", out var valuesNode) && valuesNode is not null;
        var sourceRange = OptionalString(props, "sourceRange");
        if (hasValues == (sourceRange is not null))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a list rule takes EITHER 'values' OR 'sourceRange'.",
                "Pass literal values ({\"values\":[\"todo\",\"doing\",\"done\"]}) or point at a range " +
                "({\"sourceRange\":\"/Sheet1/Z1:Z10\"}).");
        }

        if (hasValues)
        {
            if (valuesNode is not JsonArray array || array.Count == 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: 'values' must be a non-empty array of strings.",
                    "Pass e.g. {\"values\":[\"todo\",\"doing\",\"done\"]}.");
            }

            var items = new List<string>(array.Count);
            foreach (var node in array)
            {
                var item = node is JsonValue v && v.GetValueKind() == JsonValueKind.String
                    ? v.GetValue<string>()
                    : node?.ToJsonString(JsonDefaults.Options);
                if (string.IsNullOrEmpty(item))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{opIndex}]: list values must be non-empty strings.",
                        "Pass e.g. {\"values\":[\"todo\",\"doing\",\"done\"]}.");
                }

                if (item.Contains(',') || item.Contains('"'))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{opIndex}]: list value '{item}' contains a comma or quote, which the xlsx " +
                        "literal-list format cannot store.",
                        "Put the values in cells instead and reference them: " +
                        "{\"sourceRange\":\"/Sheet1/Z1:Z10\"}.");
                }

                items.Add(item);
            }

            var literal = "\"" + string.Join(",", items) + "\"";
            if (literal.Length > 255)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: the literal list is {literal.Length} characters; Excel caps it at 255.",
                    "Put the values in cells instead and reference them: {\"sourceRange\":\"/Sheet1/Z1:Z10\"}.");
            }

            validation.List(literal, inCellDropdown: true);
            return;
        }

        var resolved = ResolveSourceRange(workbook, sheet, sourceRange!, opIndex);
        validation.List(resolved, inCellDropdown: true);
    }

    /// <summary>A sourceRange is a path (<c>/Sheet1/Z1:Z10</c>) or a bare range on the rule's own sheet.</summary>
    private static IXLRange ResolveSourceRange(XLWorkbook workbook, IXLWorksheet sheet, string text, int opIndex)
    {
        var pathText = text.StartsWith('/') ? text : ExcelPaths.SheetPath(sheet) + "/" + text.ToUpperInvariant();
        var target = ExcelPaths.Resolve(workbook, pathText);
        return target.Kind switch
        {
            ExcelTargetKind.Range => target.Range!,
            ExcelTargetKind.Cell => target.Cell!.AsRange(),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'sourceRange' must address a cell or range, not '{text}'.",
                "Pass e.g. {\"sourceRange\":\"/Sheet1/Z1:Z10\"} (or a bare Z1:Z10 on the rule's sheet)."),
        };
    }

    private static void AddComparison(IXLDataValidation validation, string kind, JsonObject props, int opIndex)
    {
        var op = OptionalString(props, "operator") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: {kind} needs an 'operator'.",
            "Supported operators: " + string.Join(", ", Operators) + ".",
            candidates: Operators);
        if (!Operators.Contains(op, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: unknown operator '{op}'.",
                "Supported operators: " + string.Join(", ", Operators) + ".",
                candidates: Operators);
        }

        var needsPair = op is "between" or "notBetween";
        var hasValue2 = props.TryGetPropertyValue("value2", out var v2) && v2 is not null;
        if (needsPair && !hasValue2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: operator '{op}' needs both value and value2.",
                "Pass e.g. {\"operator\":\"between\",\"value\":1,\"value2\":10}.");
        }

        if (!needsPair && hasValue2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'value2' only applies to between/notBetween.",
                "Drop value2, or switch the operator to between.");
        }

        switch (kind)
        {
            case "wholeNumber":
                ApplyOperator(
                    validation.WholeNumber, op,
                    RequiredInt(props, "value", opIndex),
                    needsPair ? RequiredInt(props, "value2", opIndex) : 0);
                break;
            case "decimal":
                ApplyOperator(
                    validation.Decimal, op,
                    RequiredNumber(props, "value", opIndex),
                    needsPair ? RequiredNumber(props, "value2", opIndex) : 0d);
                break;
            case "date":
                ApplyOperator(
                    validation.Date, op,
                    RequiredDate(props, "value", opIndex),
                    needsPair ? RequiredDate(props, "value2", opIndex) : default);
                break;
            default: // textLength
                ApplyOperator(
                    validation.TextLength, op,
                    RequiredInt(props, "value", opIndex),
                    needsPair ? RequiredInt(props, "value2", opIndex) : 0);
                break;
        }
    }

    private static void ApplyOperator(XLValidationCriteria criteria, string op, int value, int value2) =>
        ApplyOperatorText(
            criteria, op,
            value.ToString(CultureInfo.InvariantCulture),
            value2.ToString(CultureInfo.InvariantCulture));

    private static void ApplyOperator(XLValidationCriteria criteria, string op, double value, double value2) =>
        ApplyOperatorText(
            criteria, op,
            value.ToString(CultureInfo.InvariantCulture),
            value2.ToString(CultureInfo.InvariantCulture));

    private static void ApplyOperator(XLValidationCriteria criteria, string op, DateTime value, DateTime value2) =>
        ApplyOperatorText(
            criteria, op,
            value.ToOADate().ToString(CultureInfo.InvariantCulture),
            value2 == default ? "0" : value2.ToOADate().ToString(CultureInfo.InvariantCulture));

    private static void ApplyOperatorText(XLValidationCriteria criteria, string op, string value, string value2)
    {
        switch (op)
        {
            case "between":
                criteria.Between(value, value2);
                break;
            case "notBetween":
                criteria.NotBetween(value, value2);
                break;
            case ">":
                criteria.GreaterThan(value);
                break;
            case "<":
                criteria.LessThan(value);
                break;
            case ">=":
                criteria.EqualOrGreaterThan(value);
                break;
            case "<=":
                criteria.EqualOrLessThan(value);
                break;
            case "==":
                criteria.EqualTo(value);
                break;
            default: // "!=" — the operator list was validated above
                criteria.NotEqualTo(value);
                break;
        }
    }

    private static void ApplyCommonProps(IXLDataValidation validation, JsonObject props, int opIndex)
    {
        if (props.TryGetPropertyValue("allowBlank", out var blankNode) && blankNode is not null)
        {
            validation.IgnoreBlanks = blankNode.GetValue<bool>();
        }

        if (OptionalString(props, "inputTitle") is { } inputTitle)
        {
            validation.InputTitle = inputTitle;
            validation.ShowInputMessage = true;
        }

        if (OptionalString(props, "inputMessage") is { } inputMessage)
        {
            validation.InputMessage = inputMessage;
            validation.ShowInputMessage = true;
        }

        if (OptionalString(props, "errorTitle") is { } errorTitle)
        {
            validation.ErrorTitle = errorTitle;
            validation.ShowErrorMessage = true;
        }

        if (OptionalString(props, "errorMessage") is { } errorMessage)
        {
            validation.ErrorMessage = errorMessage;
            validation.ShowErrorMessage = true;
        }

        if (OptionalString(props, "errorStyle") is { } styleText)
        {
            validation.ErrorStyle = styleText switch
            {
                "stop" => XLErrorStyle.Stop,
                "warning" => XLErrorStyle.Warning,
                "information" => XLErrorStyle.Information,
                _ => throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown errorStyle '{styleText}'.",
                    "Use one of: " + string.Join(", ", ErrorStyles) + ".",
                    candidates: ErrorStyles),
            };
            validation.ShowErrorMessage = true;
        }
    }

    // ----- find / describe ------------------------------------------------------

    /// <summary>Finds <c>dataValidation[i]</c> on a sheet or throws <c>invalid_path</c> with real candidates.</summary>
    public static IXLDataValidation Find(ExcelTarget target)
    {
        var rules = target.Sheet.DataValidations.ToList();
        var index = target.DataValidationIndex!.Value;
        if (index > rules.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No dataValidation[{index}] on sheet '{target.Sheet.Name}' ({rules.Count} rule(s) exist).",
                rules.Count > 0
                    ? "Rule indices are 1-based per sheet; run 'aioffice read --view structure' to list them."
                    : "This sheet has no data validations; add one with {op:add, type:dataValidation, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + "/B2:B50, props:{kind:\"list\", values:[…]}}.",
                candidates: rules.Count > 0
                    ? [.. Enumerable.Range(1, rules.Count).Select(i => ExcelPaths.DataValidationPath(target.Sheet, i))]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return rules[index - 1];
    }

    /// <summary>One rule as agents see it (get and read --view structure).</summary>
    public static object Describe(IXLWorksheet sheet, IXLDataValidation validation, int index)
    {
        var kind = KindName(validation.AllowedValues);
        var isList = validation.AllowedValues == XLAllowedValues.List;
        List<string>? listValues = null;
        string? sourceRange = null;
        if (isList)
        {
            var formula = validation.Value;
            if (formula.Length >= 2 && formula[0] == '"' && formula[^1] == '"')
            {
                listValues = [.. formula[1..^1].Split(',')];
            }
            else
            {
                sourceRange = formula;
            }
        }

        return new
        {
            path = ExcelPaths.DataValidationPath(sheet, index),
            kind = "dataValidation",
            sheet = sheet.Name,
            ranges = validation.Ranges.Select(r => r.RangeAddress.ToString()).ToList(),
            dvKind = kind,
            @operator = isList || validation.AllowedValues == XLAllowedValues.AnyValue
                ? null
                : OperatorName(validation.Operator),
            value = isList ? null : ValueText(validation, validation.MinValue),
            value2 = !isList && validation.Operator is XLOperator.Between or XLOperator.NotBetween
                ? ValueText(validation, validation.MaxValue)
                : null,
            values = listValues,
            sourceRange,
            allowBlank = validation.IgnoreBlanks,
            inputTitle = NullIfEmpty(validation.InputTitle),
            inputMessage = NullIfEmpty(validation.InputMessage),
            errorTitle = NullIfEmpty(validation.ErrorTitle),
            errorMessage = NullIfEmpty(validation.ErrorMessage),
            errorStyle = validation.ErrorStyle switch
            {
                XLErrorStyle.Warning => "warning",
                XLErrorStyle.Information => "information",
                _ => "stop",
            },
        };
    }

    /// <summary>Every rule on a sheet, for read --view structure.</summary>
    public static List<object> List(IXLWorksheet sheet) =>
        [.. sheet.DataValidations.Select((validation, i) => Describe(sheet, validation, i + 1))];

    private static string KindName(XLAllowedValues allowed) => allowed switch
    {
        XLAllowedValues.List => "list",
        XLAllowedValues.WholeNumber => "wholeNumber",
        XLAllowedValues.Decimal => "decimal",
        XLAllowedValues.Date => "date",
        XLAllowedValues.TextLength => "textLength",
        XLAllowedValues.Time => "time",
        XLAllowedValues.Custom => "custom",
        _ => "anyValue",
    };

    private static string OperatorName(XLOperator op) => op switch
    {
        XLOperator.Between => "between",
        XLOperator.NotBetween => "notBetween",
        XLOperator.GreaterThan => ">",
        XLOperator.LessThan => "<",
        XLOperator.EqualOrGreaterThan => ">=",
        XLOperator.EqualOrLessThan => "<=",
        XLOperator.EqualTo => "==",
        _ => "!=",
    };

    /// <summary>Date rules store OADate serials; report them back as ISO dates.</summary>
    private static string? ValueText(IXLDataValidation validation, string stored)
    {
        if (string.IsNullOrEmpty(stored))
        {
            return null;
        }

        if (validation.AllowedValues == XLAllowedValues.Date &&
            double.TryParse(stored, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
        {
            return DateTime.FromOADate(serial).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        return stored;
    }

    private static string? NullIfEmpty(string? text) => string.IsNullOrEmpty(text) ? null : text;

    private static void GuardProps(JsonObject props, IReadOnlyList<string> allowed, int opIndex)
    {
        foreach (var (key, _) in props)
        {
            if (!allowed.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown dataValidation prop '{key}' for this kind.",
                    "Supported props here: " + string.Join(", ", allowed) + ".",
                    candidates: allowed);
            }
        }
    }

    private static double RequiredNumber(JsonObject props, string key, int opIndex)
    {
        if (props.TryGetPropertyValue(key, out var node) && node is JsonValue value)
        {
            if (value.GetValueKind() == JsonValueKind.Number && value.TryGetValue<double>(out var d))
            {
                return d;
            }

            if (value.GetValueKind() == JsonValueKind.Number)
            {
                return Convert.ToDouble(value.GetValue<object>(), CultureInfo.InvariantCulture);
            }

            if (value.GetValueKind() == JsonValueKind.String &&
                double.TryParse(value.GetValue<string>(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: '{key}' must be a number.",
            $"Pass e.g. {{\"{key}\":10}}.");
    }

    private static int RequiredInt(JsonObject props, string key, int opIndex)
    {
        var number = RequiredNumber(props, key, opIndex);
        if (number != Math.Floor(number) || number is < int.MinValue or > int.MaxValue)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{key}' must be a whole number for this kind.",
                $"Pass an integer like {{\"{key}\":10}}, or switch kind to decimal.");
        }

        return (int)number;
    }

    private static DateTime RequiredDate(JsonObject props, string key, int opIndex)
    {
        if (props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
            value.GetValueKind() == JsonValueKind.String &&
            DateTime.TryParse(
                value.GetValue<string>(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
        {
            return parsed;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: '{key}' must be an ISO date for a date rule.",
            $"Pass e.g. {{\"{key}\":\"2026-01-31\"}}.");
    }

    private static string? OptionalString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;
}
