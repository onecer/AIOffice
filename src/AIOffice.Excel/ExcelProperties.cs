using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// Document properties (M7): the workbook's core OOXML properties (title,
/// author, subject…) plus typed custom properties. Core properties come from
/// ClosedXML's <see cref="XLWorkbookProperties"/> (the CoreFilePropertiesPart);
/// custom properties from <see cref="IXLWorkbook.CustomProperties"/> (the
/// CustomFilePropertiesPart), typed string/number/boolean/dateTime.
///
/// <para>Read: <c>read --view properties</c> returns both blocks. Set:
/// <c>{op:set, path:"/properties", props:{title:"Q3", custom:{Region:"EU",
/// Reviewed:true}}}</c>. Get mirrors read. A custom value of <c>null</c>
/// deletes that custom property.</para>
///
/// <para>The <c>_aioffice_cellStyles</c> registry property
/// (<see cref="ExcelCellStyles.RegistryProperty"/>) is aioffice bookkeeping and
/// is hidden from the custom-property surface so agents never see or clobber
/// it.</para>
/// </summary>
internal static class ExcelProperties
{
    /// <summary>The core properties agents can read and set, in display order.</summary>
    public static readonly IReadOnlyList<string> CoreKeys =
    [
        "title", "subject", "author", "manager", "company", "category",
        "keywords", "comments", "status", "lastModifiedBy",
    ];

    // ----- read / get ---------------------------------------------------------

    /// <summary>
    /// Both property blocks (core + custom) for read --view properties and
    /// get /properties. M10: the core/custom blocks nest under a "properties" key
    /// (data.properties.core.title) so the envelope shape is identical across
    /// docx/xlsx/pptx — the unified pre-1.0 contract shape.
    /// </summary>
    public static object Describe(XLWorkbook workbook)
    {
        var p = workbook.Properties;
        return new
        {
            path = "/properties",
            kind = "properties",
            properties = new
            {
                core = new
                {
                    title = NullIfEmpty(p.Title),
                    subject = NullIfEmpty(p.Subject),
                    author = NullIfEmpty(p.Author),
                    manager = NullIfEmpty(p.Manager),
                    company = NullIfEmpty(p.Company),
                    category = NullIfEmpty(p.Category),
                    keywords = NullIfEmpty(p.Keywords),
                    comments = NullIfEmpty(p.Comments),
                    status = NullIfEmpty(p.Status),
                    lastModifiedBy = NullIfEmpty(p.LastModifiedBy),
                    created = p.Created == default ? null : Iso(p.Created),
                    modified = p.Modified == default ? null : Iso(p.Modified),
                },
                custom = CustomMap(workbook),
            },
        };
    }

    /// <summary>The user-visible custom properties (the aioffice registry is hidden).</summary>
    public static JsonObject CustomMap(XLWorkbook workbook)
    {
        var map = new JsonObject();
        foreach (var property in workbook.CustomProperties
                     .Where(p => !string.Equals(p.Name, ExcelCellStyles.RegistryProperty, StringComparison.Ordinal))
                     .OrderBy(p => p.Name, StringComparer.Ordinal))
        {
            map[property.Name] = CustomValueNode(property);
        }

        return map;
    }

    private static JsonNode? CustomValueNode(IXLCustomProperty property) => property.Type switch
    {
        XLCustomPropertyType.Number => JsonValue.Create(Convert.ToDouble(property.Value, CultureInfo.InvariantCulture)),
        XLCustomPropertyType.Boolean => JsonValue.Create(Convert.ToBoolean(property.Value, CultureInfo.InvariantCulture)),
        XLCustomPropertyType.Date => JsonValue.Create(
            Iso(Convert.ToDateTime(property.Value, CultureInfo.InvariantCulture))),
        _ => JsonValue.Create(property.Value?.ToString() ?? string.Empty),
    };

    // ----- set ----------------------------------------------------------------

    /// <summary>
    /// Applies a <c>set /properties</c> op. Core keys map to ClosedXML's typed
    /// workbook properties; the optional <c>custom</c> object carries typed
    /// custom properties (a null value deletes the key). Returns the list of
    /// applied property names for the response.
    /// </summary>
    public static List<string> ApplySet(XLWorkbook workbook, EditOp op, int index)
    {
        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: set /properties needs props.",
                "Pass core props like {\"title\":\"Q3\",\"author\":\"Onecer\"} and/or {\"custom\":{\"Region\":\"EU\"}}.");
        }

        var applied = new List<string>();
        var p = workbook.Properties;

        foreach (var (key, node) in props)
        {
            if (string.Equals(key, "custom", StringComparison.Ordinal))
            {
                continue; // handled below
            }

            if (!CoreKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: unknown property '{key}'.",
                    "Core props: " + string.Join(", ", CoreKeys) + ". Custom props go under \"custom\": {...}.",
                    candidates: CoreKeys);
            }

            var text = node is null ? string.Empty : ScalarString(node, key, index);
            SetCore(p, key, text);
            applied.Add(key);
        }

        if (props.TryGetPropertyValue("custom", out var customNode) && customNode is not null)
        {
            if (customNode is not JsonObject customObject)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: 'custom' must be a JSON object of name → value.",
                    "Pass e.g. {\"custom\":{\"Region\":\"EU\",\"Reviewed\":true,\"Score\":9.5}}.");
            }

            foreach (var (name, valueNode) in customObject)
            {
                applied.Add(SetCustom(workbook, name, valueNode, index));
            }
        }

        return applied;
    }

    private static void SetCore(XLWorkbookProperties p, string key, string value)
    {
        switch (key)
        {
            case "title": p.Title = value; break;
            case "subject": p.Subject = value; break;
            case "author": p.Author = value; break;
            case "manager": p.Manager = value; break;
            case "company": p.Company = value; break;
            case "category": p.Category = value; break;
            case "keywords": p.Keywords = value; break;
            case "comments": p.Comments = value; break;
            case "status": p.Status = value; break;
            case "lastModifiedBy": p.LastModifiedBy = value; break;
        }
    }

    /// <summary>Sets (or deletes, on a null value) a typed custom property; returns its display name.</summary>
    private static string SetCustom(XLWorkbook workbook, string name, JsonNode? valueNode, int index)
    {
        if (string.Equals(name, ExcelCellStyles.RegistryProperty, StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{name}' is reserved for aioffice bookkeeping.",
                "Pick a different custom property name.");
        }

        var exists = workbook.CustomProperties.Any(p => string.Equals(p.Name, name, StringComparison.Ordinal));
        if (valueNode is null || valueNode.GetValueKind() == JsonValueKind.Null)
        {
            if (exists)
            {
                workbook.CustomProperties.Delete(name);
            }

            return "custom:" + name;
        }

        if (exists)
        {
            workbook.CustomProperties.Delete(name); // replace: ClosedXML's Add throws on a duplicate key
        }

        switch (valueNode.GetValueKind())
        {
            case JsonValueKind.True:
            case JsonValueKind.False:
                workbook.CustomProperties.Add(name, valueNode.GetValue<bool>());
                break;
            case JsonValueKind.Number:
                workbook.CustomProperties.Add(name, valueNode.GetValue<double>());
                break;
            case JsonValueKind.String:
                var text = valueNode.GetValue<string>();
                if (TryIsoDate(text, out var date))
                {
                    workbook.CustomProperties.Add(name, date);
                }
                else
                {
                    workbook.CustomProperties.Add(name, text);
                }

                break;
            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: custom property '{name}' must be a string, number, boolean or null.",
                    "Pass a JSON scalar, e.g. {\"custom\":{\"Region\":\"EU\"}}; null deletes the property.");
        }

        return "custom:" + name;
    }

    // ----- helpers ------------------------------------------------------------

    private static string ScalarString(JsonNode node, string key, int index)
    {
        if (node is JsonValue value)
        {
            return value.GetValueKind() switch
            {
                JsonValueKind.String => value.GetValue<string>(),
                JsonValueKind.Number => value.GetValue<double>().ToString(CultureInfo.InvariantCulture),
                JsonValueKind.True => "true",
                JsonValueKind.False => "false",
                _ => node.ToJsonString(JsonDefaults.Options),
            };
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: property '{key}' must be a scalar string.",
            "Pass e.g. {\"" + key + "\":\"value\"}.");
    }

    private static bool TryIsoDate(string text, out DateTime value)
    {
        value = default;
        return text.Length >= 10 &&
               (text[4] == '-' && text[7] == '-') &&
               DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
    }

    private static string Iso(DateTime value) =>
        value.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);

    private static string? NullIfEmpty(string? text) => string.IsNullOrEmpty(text) ? null : text;
}
