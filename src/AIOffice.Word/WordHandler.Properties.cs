using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;

namespace AIOffice.Word;

/// <summary>
/// Document properties live on a virtual <c>/properties</c> node: the core
/// properties (title/subject/author…) come from the package's
/// CoreFilePropertiesPart, and typed custom properties (string/number/bool/date)
/// from the CustomFilePropertiesPart. <c>get /properties</c> and
/// <c>read --view properties</c> both project this node; <c>set /properties</c>
/// writes it.
/// </summary>
public sealed partial class WordHandler
{
    /// <summary>The fixed format ID custom-property parts carry (FMTID).</summary>
    private const string CustomPropFmtId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    // ------------------------------------------------------------------- read

    /// <summary>get /properties / read --view properties: {core:{…}, custom:{…}}.</summary>
    private static Dictionary<string, object?> PropertiesShape(WordprocessingDocument doc)
    {
        var core = doc.PackageProperties;
        var coreShape = new Dictionary<string, object?>
        {
            ["title"] = NullIfEmpty(core.Title),
            ["subject"] = NullIfEmpty(core.Subject),
            ["author"] = NullIfEmpty(core.Creator),
            ["keywords"] = NullIfEmpty(core.Keywords),
            ["category"] = NullIfEmpty(core.Category),
            ["comments"] = NullIfEmpty(core.Description),
            ["lastModifiedBy"] = NullIfEmpty(core.LastModifiedBy),
            ["created"] = FormatDate(core.Created),
            ["modified"] = FormatDate(core.Modified),
            ["revision"] = NullIfEmpty(core.Revision),
        };

        var custom = new Dictionary<string, object?>();
        foreach (var prop in CustomProperties(doc))
        {
            if (prop.Name?.Value is { Length: > 0 } name)
            {
                custom[name] = CustomPropertyValue(prop);
            }
        }

        return new Dictionary<string, object?>
        {
            ["core"] = coreShape,
            ["custom"] = custom,
        };
    }

    private static string? NullIfEmpty(string? value) => value is { Length: > 0 } ? value : null;

    private static string? FormatDate(DateTime? value) =>
        value is { } d ? d.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture) : null;

    private static IEnumerable<CustomDocumentProperty> CustomProperties(WordprocessingDocument doc) =>
        doc.CustomFilePropertiesPart?.Properties?.Elements<CustomDocumentProperty>() ?? [];

    /// <summary>The typed value behind one w:property (vt:lpwstr/r8/bool/filetime).</summary>
    private static object? CustomPropertyValue(CustomDocumentProperty prop)
    {
        if (prop.VTBool is { } b)
        {
            return string.Equals(b.Text, "true", StringComparison.OrdinalIgnoreCase) || b.Text == "1";
        }

        if (prop.VTLPWSTR is { } s)
        {
            return s.Text;
        }

        if (prop.VTFileTime is { } ft)
        {
            return ft.Text;
        }

        if (prop.VTDouble?.Text is { } dbl &&
            double.TryParse(dbl, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
        {
            return d;
        }

        if (prop.VTInteger?.Text is { } i &&
            double.TryParse(i, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
        {
            return n;
        }

        // Fall back to whatever scalar text the variant carries.
        return prop.InnerText is { Length: > 0 } text ? text : null;
    }

    // -------------------------------------------------------------------- set

    /// <summary>
    /// <c>{"op":"set","path":"/properties","props":{"title":…,"author":…,
    /// "keywords":"a;b","custom":{"Project":"X","Reviewed":true}}}</c>: core
    /// props go to the CoreFilePropertiesPart, custom props become typed
    /// w:property entries (string/number/bool/date) in the CustomFilePropertiesPart.
    /// </summary>
    private static object ApplySetProperties(WordprocessingDocument doc, EditOp op)
    {
        var props = RequireProps(op);
        var setCore = new List<string>();
        var setCustom = new List<string>();
        JsonObject? customObject = null;

        foreach (var (name, node) in props)
        {
            if (name == "custom")
            {
                customObject = node as JsonObject ?? throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "props.custom must be an object of name:value pairs.",
                    "Example: \"custom\":{\"Project\":\"Acme\",\"Reviewed\":true,\"Budget\":1000}.");
                continue;
            }

            SetCoreProperty(doc, name, node);
            setCore.Add(name);
        }

        if (customObject is not null)
        {
            foreach (var (name, node) in customObject)
            {
                SetCustomProperty(doc, name, node);
                setCustom.Add(name);
            }
        }

        if (setCore.Count == 0 && setCustom.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "set /properties needs at least one core property or a custom object.",
                "Example: {\"op\":\"set\",\"path\":\"/properties\",\"props\":{\"title\":\"Q3 Report\",\"custom\":{\"Reviewed\":true}}}.");
        }

        return new { op = "set", path = "/properties", core = setCore, custom = setCustom };
    }

    private static readonly string[] CorePropertyNames =
        ["title", "subject", "author", "keywords", "category", "comments", "lastModifiedBy", "created", "modified", "revision"];

    private static void SetCoreProperty(WordprocessingDocument doc, string name, JsonNode? node)
    {
        var core = doc.PackageProperties;
        var value = node is null ? string.Empty : NodeToString(node);

        switch (name)
        {
            case "title":
                core.Title = value;
                break;
            case "subject":
                core.Subject = value;
                break;
            case "author" or "creator":
                core.Creator = value;
                break;
            case "keywords":
                core.Keywords = value;
                break;
            case "category":
                core.Category = value;
                break;
            case "comments":
                core.Description = value;
                break;
            case "lastModifiedBy":
                core.LastModifiedBy = value;
                break;
            case "revision":
                core.Revision = value;
                break;
            case "created":
                core.Created = ParsePropertyDate(name, value);
                break;
            case "modified":
                core.Modified = ParsePropertyDate(name, value);
                break;
            default:
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"'{name}' is not a settable core property.",
                    $"Core properties: {string.Join(", ", CorePropertyNames)}. " +
                    "Use a custom property instead: \"custom\":{\"" + name + "\":…}.",
                    candidates: CorePropertyNames);
        }
    }

    private static DateTime ParsePropertyDate(string name, string value)
    {
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var parsed))
        {
            return parsed;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"'{name}' is not a valid date: '{value}'.",
            "Pass an ISO 8601 date like \"2026-06-13\" or \"2026-06-13T10:00:00Z\".");
    }

    // ----------------------------------------------------------- custom props

    private static void SetCustomProperty(WordprocessingDocument doc, string name, JsonNode? node)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "A custom property needs a non-empty name.",
                "Example: \"custom\":{\"Project\":\"Acme\"}.");
        }

        var part = doc.CustomFilePropertiesPart ?? doc.AddCustomFilePropertiesPart();
        var properties = part.Properties ??= new DocumentFormat.OpenXml.CustomProperties.Properties();

        // Remove any existing property of the same name (case-insensitive, as Word treats them).
        foreach (var existing in properties.Elements<CustomDocumentProperty>()
                     .Where(p => string.Equals(p.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
                     .ToList())
        {
            existing.Remove();
        }

        // null clears the property (no value element written back).
        if (node is null)
        {
            RenumberCustomProperties(properties);
            return;
        }

        var prop = new CustomDocumentProperty
        {
            FormatId = CustomPropFmtId,
            Name = name,
        };
        prop.Append(CustomVariant(node));
        properties.AppendChild(prop);
        RenumberCustomProperties(properties);
    }

    /// <summary>Maps a JSON scalar to the matching vt:* variant (bool/number/date/string).</summary>
    private static DocumentFormat.OpenXml.OpenXmlElement CustomVariant(JsonNode node)
    {
        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var b))
            {
                return new VTBool(b ? "true" : "false");
            }

            if (value.TryGetValue<double>(out var d))
            {
                return new VTDouble(d.ToString(CultureInfo.InvariantCulture));
            }

            if (value.TryGetValue<string>(out var s))
            {
                // An ISO date string with a date marker becomes a filetime variant.
                if (LooksLikeDate(s) &&
                    DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var dt))
                {
                    return new VTFileTime(dt.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
                }

                return new VTLPWSTR(s);
            }
        }

        // Objects/arrays serialize to their JSON text as a string property.
        return new VTLPWSTR(NodeToString(node));
    }

    private static bool LooksLikeDate(string s) =>
        s.Length >= 8 && s.Length <= 32 && s.Contains('-', StringComparison.Ordinal) &&
        (s[0] is >= '0' and <= '9');

    /// <summary>Custom properties carry sequential pids starting at 2 (1 + 1 per property).</summary>
    private static void RenumberCustomProperties(DocumentFormat.OpenXml.CustomProperties.Properties properties)
    {
        var pid = 1;
        foreach (var prop in properties.Elements<CustomDocumentProperty>())
        {
            pid++;
            prop.PropertyId = pid;
        }
    }
}
