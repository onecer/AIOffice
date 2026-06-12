using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// Typed-value bridging between the JSON wire surface and ClosedXML.
/// Auto-typing rules for strings: a leading <c>=</c> means formula, ISO dates
/// (<c>2024-05-01</c> / <c>2024-05-01T08:30:00</c>) become DateTime,
/// <c>true</c>/<c>false</c> become booleans, parseable numbers become numbers,
/// everything else stays text. Pass <c>valueType: "text"</c> to keep a literal.
/// </summary>
internal static partial class ExcelValues
{
    [GeneratedRegex(@"^\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2}(\.\d+)?)?)?$")]
    private static partial Regex IsoDate();

    public static readonly IReadOnlyList<string> ValueTypes = ["auto", "text", "number", "boolean", "dateTime"];

    /// <summary>A parsed wire value: either a literal cell value or a formula string.</summary>
    public sealed record ParsedValue(XLCellValue Value, string? Formula)
    {
        public bool IsFormula => Formula is not null;
    }

    /// <summary>Parses a JSON scalar into a typed cell value or formula.</summary>
    public static ParsedValue Parse(JsonNode? node, string? valueType = null)
    {
        valueType ??= "auto";
        if (!ValueTypes.Contains(valueType, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown valueType '{valueType}'.",
                "Use one of: " + string.Join(", ", ValueTypes) + ".",
                candidates: ValueTypes);
        }

        if (node is null)
        {
            return new ParsedValue(Blank.Value, null);
        }

        return node.GetValueKind() switch
        {
            JsonValueKind.Null => new ParsedValue(Blank.Value, null),
            JsonValueKind.Number => Coerce(ToDouble(node), valueType),
            JsonValueKind.True => Coerce(true, valueType),
            JsonValueKind.False => Coerce(false, valueType),
            JsonValueKind.String => ParseString(node.GetValue<string>(), valueType),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A cell value must be a JSON scalar (string, number, boolean or null).",
                "For ranges pass a 2D array under the 'values' prop instead of 'value'."),
        };
    }

    /// <summary>
    /// JSON numbers arrive either as parsed elements or as in-memory nodes
    /// backed by int/long/decimal; bridge them all to double.
    /// </summary>
    private static double ToDouble(JsonNode node) =>
        node is JsonValue value && value.TryGetValue<double>(out var d)
            ? d
            : Convert.ToDouble(node.GetValue<object>(), CultureInfo.InvariantCulture);

    private static ParsedValue Coerce(XLCellValue value, string valueType) => valueType switch
    {
        "text" => new ParsedValue(value.ToString(CultureInfo.InvariantCulture), null),
        _ => new ParsedValue(value, null),
    };

    private static ParsedValue ParseString(string text, string valueType)
    {
        if (valueType == "text")
        {
            return new ParsedValue(text, null);
        }

        if (text.StartsWith('='))
        {
            return new ParsedValue(Blank.Value, text);
        }

        switch (valueType)
        {
            case "number":
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var forcedNumber))
                {
                    return new ParsedValue(forcedNumber, null);
                }

                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"'{text}' is not a number but valueType is 'number'.",
                    "Pass a numeric literal like 42 or 3.14, or drop valueType for auto-typing.");

            case "boolean":
                if (bool.TryParse(text, out var forcedBool))
                {
                    return new ParsedValue(forcedBool, null);
                }

                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"'{text}' is not a boolean but valueType is 'boolean'.",
                    "Pass true or false, or drop valueType for auto-typing.");

            case "dateTime":
                if (TryParseIsoDate(text, out var forcedDate))
                {
                    return new ParsedValue(forcedDate, null);
                }

                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"'{text}' is not an ISO date but valueType is 'dateTime'.",
                    "Use ISO 8601 like 2024-05-01 or 2024-05-01T08:30:00.");

            default: // auto
                if (bool.TryParse(text, out var b))
                {
                    return new ParsedValue(b, null);
                }

                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
                {
                    return new ParsedValue(n, null);
                }

                if (TryParseIsoDate(text, out var d))
                {
                    return new ParsedValue(d, null);
                }

                return new ParsedValue(text, null);
        }
    }

    private static bool TryParseIsoDate(string text, out DateTime value)
    {
        value = default;
        return IsoDate().IsMatch(text) &&
               DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
    }

    /// <summary>Converts a cell value to its JSON wire shape (dates as ISO strings, errors as their display text).</summary>
    public static object? ToJson(XLCellValue value) => value.Type switch
    {
        XLDataType.Blank => null,
        XLDataType.Boolean => value.GetBoolean(),
        XLDataType.Number => value.GetNumber(),
        XLDataType.Text => value.GetText(),
        XLDataType.Error => value.ToString(CultureInfo.InvariantCulture),
        XLDataType.DateTime => value.GetDateTime().ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture),
        XLDataType.TimeSpan => value.GetTimeSpan().ToString("c", CultureInfo.InvariantCulture),
        _ => value.ToString(CultureInfo.InvariantCulture),
    };

    /// <summary>The wire name of a cell value's type.</summary>
    public static string TypeName(XLDataType type) => type switch
    {
        XLDataType.Blank => "blank",
        XLDataType.Boolean => "boolean",
        XLDataType.Number => "number",
        XLDataType.Text => "text",
        XLDataType.Error => "error",
        XLDataType.DateTime => "dateTime",
        XLDataType.TimeSpan => "timeSpan",
        _ => "unknown",
    };

    /// <summary>
    /// The number-format-applied display text of a cell. Never throws: cells whose
    /// formula the engine cannot evaluate fall back to their formula text.
    /// </summary>
    public static string SafeFormatted(IXLCell cell)
    {
        try
        {
            return cell.GetFormattedString();
        }
        catch (Exception)
        {
            return cell.HasFormula ? "=" + cell.FormulaA1 : string.Empty;
        }
    }

    /// <summary>Escapes one CSV field (RFC 4180 style).</summary>
    public static string CsvEscape(string field) =>
        field.Contains(',') || field.Contains('"') || field.Contains('\n') || field.Contains('\r')
            ? "\"" + field.Replace("\"", "\"\"", StringComparison.Ordinal) + "\""
            : field;

    /// <summary>Renders a JSON scalar as plain text for template substitution.</summary>
    public static string TemplateText(JsonNode? node) => node switch
    {
        null => string.Empty,
        JsonValue v when v.GetValueKind() == JsonValueKind.String => v.GetValue<string>(),
        _ => node.ToJsonString(JsonDefaults.Options),
    };

    /// <summary>Appends an HTML-escaped string.</summary>
    public static StringBuilder AppendEscaped(StringBuilder sb, string text)
    {
        foreach (var c in text)
        {
            _ = c switch
            {
                '<' => sb.Append("&lt;"),
                '>' => sb.Append("&gt;"),
                '&' => sb.Append("&amp;"),
                '"' => sb.Append("&quot;"),
                _ => sb.Append(c),
            };
        }

        return sb;
    }
}
