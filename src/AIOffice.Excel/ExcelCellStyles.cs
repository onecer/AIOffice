using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// Named cell styles (M7): reusable, named bundles of formatting an agent
/// defines once and applies by name to any cell/range.
///
/// <para>ClosedXML 0.105 has no API for the workbook's named-style table
/// (<c>cellStyles</c>/<c>cellStyleXfs</c>), so the style DEFINITIONS live in a
/// single custom workbook property (<c>_aioffice_cellStyles</c>, a JSON map)
/// which survives ClosedXML's saves untouched. That registry is the source of
/// truth for <c>read --view styles</c>, <c>get /style[@name=…]</c> and
/// <c>remove</c>.</para>
///
/// <para>Applying a style — <c>{op:set, path:/Sheet1/B2:B10,
/// props:{cellStyle:"Currency-Red"}}</c> — writes the style's concrete
/// formatting (numberFormat, bold, fill, color, border) straight onto the
/// target cells through ClosedXML, so the cached display text reflects the
/// number format immediately and the result is byte-for-byte what Excel would
/// show. A raw post-save pass (<see cref="MaterializeAfterSave"/>) ALSO emits a
/// real <c>cellStyle</c> entry per registered name so Excel surfaces the style
/// in its gallery; the file stays OpenXmlValidator-clean.</para>
/// </summary>
internal static partial class ExcelCellStyles
{
    /// <summary>The custom-property key the style registry is stored under.</summary>
    public const string RegistryProperty = "_aioffice_cellStyles";

    private static readonly IReadOnlyList<string> AddProps =
        ["name", "numberFormat", "bold", "italic", "fill", "color", "border"];

    private static readonly IReadOnlyList<string> BorderStyles =
        ["none", "thin", "medium", "thick", "dashed", "dotted", "double", "hair"];

    /// <summary>A single named style's formatting (every facet optional).</summary>
    public sealed record Definition
    {
        public required string Name { get; init; }

        public string? NumberFormat { get; init; }

        public bool? Bold { get; init; }

        public bool? Italic { get; init; }

        public string? Fill { get; init; }

        public string? Color { get; init; }

        public string? Border { get; init; }
    }

    /// <summary>The <c>/style[@name=X]</c> form, peeled off the shared grammar like names/pivots.</summary>
    [GeneratedRegex(@"^/(?i:style)\[@name=(?:'(?<quoted>(?:[^']|'')+)'|(?<bare>[^\]]+))\]$")]
    private static partial Regex StylePath();

    /// <summary>Names agents may use (no control chars, 1-255 chars, not a builtin name).</summary>
    [GeneratedRegex(@"^[^\x00-\x1F\[\]]{1,255}$")]
    private static partial Regex ValidName();

    /// <summary>Detects the <c>/style[@name=X]</c> path; the name is unescaped.</summary>
    public static bool TryParsePath(string pathText, out string name)
    {
        name = string.Empty;
        var match = StylePath().Match(pathText);
        if (!match.Success)
        {
            return false;
        }

        name = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal)
            : match.Groups["bare"].Value;
        return true;
    }

    /// <summary>The canonical path aioffice emits for a named style.</summary>
    public static string PathOf(string name) => $"/style[@name={ExcelPaths.QuoteSheet(name)}]";

    // ----- registry (custom-property JSON map) --------------------------------

    /// <summary>Reads the style registry from the workbook (empty when unset).</summary>
    public static Dictionary<string, Definition> Load(XLWorkbook workbook)
    {
        var result = new Dictionary<string, Definition>(StringComparer.OrdinalIgnoreCase);
        if (!workbook.CustomProperties.Any(p =>
                string.Equals(p.Name, RegistryProperty, StringComparison.Ordinal)))
        {
            return result;
        }

        var raw = workbook.CustomProperty(RegistryProperty)?.GetValue<string>();
        if (string.IsNullOrEmpty(raw))
        {
            return result;
        }

        JsonNode? parsed;
        try
        {
            parsed = JsonNode.Parse(raw);
        }
        catch (JsonException)
        {
            return result; // a corrupt registry never crashes a read; it reads as empty
        }

        if (parsed is not JsonArray array)
        {
            return result;
        }

        foreach (var node in array)
        {
            if (node is not JsonObject obj || GetString(obj, "name") is not { } name)
            {
                continue;
            }

            result[name] = new Definition
            {
                Name = name,
                NumberFormat = GetString(obj, "numberFormat"),
                Bold = GetBool(obj, "bold"),
                Italic = GetBool(obj, "italic"),
                Fill = GetString(obj, "fill"),
                Color = GetString(obj, "color"),
                Border = GetString(obj, "border"),
            };
        }

        return result;
    }

    private static void Store(XLWorkbook workbook, IReadOnlyCollection<Definition> styles)
    {
        var array = new JsonArray();
        foreach (var style in styles)
        {
            var obj = new JsonObject { ["name"] = style.Name };
            if (style.NumberFormat is not null)
            {
                obj["numberFormat"] = style.NumberFormat;
            }

            if (style.Bold is { } bold)
            {
                obj["bold"] = bold;
            }

            if (style.Italic is { } italic)
            {
                obj["italic"] = italic;
            }

            if (style.Fill is not null)
            {
                obj["fill"] = style.Fill;
            }

            if (style.Color is not null)
            {
                obj["color"] = style.Color;
            }

            if (style.Border is not null)
            {
                obj["border"] = style.Border;
            }

            array.Add(obj);
        }

        var json = array.ToJsonString(JsonDefaults.Options);

        // ClosedXML's Add throws on a duplicate key; replace by delete + add.
        if (workbook.CustomProperties.Any(p => string.Equals(p.Name, RegistryProperty, StringComparison.Ordinal)))
        {
            workbook.CustomProperties.Delete(RegistryProperty);
        }

        workbook.CustomProperties.Add(RegistryProperty, json);
    }

    // ----- add ----------------------------------------------------------------

    /// <summary>
    /// Applies an <c>add cellStyle</c> op: registers a named style in the
    /// workbook's style registry. The op path is the document root marker the
    /// edit layer routes here (<c>/styles</c>).
    /// </summary>
    public static object Add(XLWorkbook workbook, EditOp op, int index)
    {
        var props = op.Props ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: add cellStyle needs props.",
            "Pass props like {\"name\":\"Currency-Red\",\"numberFormat\":\"$#,##0.00\",\"bold\":true}.");

        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: unknown cellStyle prop '{key}'.",
                    "Supported cellStyle props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }

        var name = GetString(props, "name");
        if (string.IsNullOrWhiteSpace(name) || !ValidName().IsMatch(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{name}' is not a usable style name.",
                "Names are 1-255 characters with no control characters or square brackets, e.g. \"Currency-Red\".");
        }

        var definition = ParseDefinition(name, props, index);

        var styles = Load(workbook);
        if (styles.ContainsKey(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: a cell style named '{name}' already exists.",
                "Remove it first ({op:remove, path:" + PathOf(name) + "}) or pick a different name.");
        }

        styles[name] = definition;
        Store(workbook, styles.Values.ToList());

        return new
        {
            op = "add",
            type = "cellStyle",
            path = PathOf(name),
            name,
            numberFormat = definition.NumberFormat,
            bold = definition.Bold,
            italic = definition.Italic,
            fill = definition.Fill,
            color = definition.Color,
            border = definition.Border,
        };
    }

    private static Definition ParseDefinition(string name, JsonObject props, int index)
    {
        var numberFormat = GetString(props, "numberFormat");
        var fill = GetString(props, "fill");
        var color = GetString(props, "color");
        var border = GetString(props, "border");

        // Validate colors/border up front so a bad style never reaches a cell.
        if (fill is not null)
        {
            _ = ParseColorOrThrow(fill, "fill", index);
        }

        if (color is not null)
        {
            _ = ParseColorOrThrow(color, "color", index);
        }

        if (border is not null && !BorderStyles.Contains(border, StringComparer.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: unknown border style '{border}'.",
                "Use one of: " + string.Join(", ", BorderStyles) + ".",
                candidates: BorderStyles);
        }

        return new Definition
        {
            Name = name,
            NumberFormat = numberFormat,
            Bold = GetBool(props, "bold"),
            Italic = GetBool(props, "italic"),
            Fill = fill,
            Color = color,
            Border = border,
        };
    }

    // ----- apply (set props:{cellStyle:X}) ------------------------------------

    /// <summary>
    /// Stamps a registered named style onto a style holder (cell/range/row/col).
    /// Throws <c>invalid_path</c> with candidates when the name is unknown.
    /// </summary>
    public static void Apply(XLWorkbook workbook, IXLStyle style, string styleName, int index)
    {
        var styles = Load(workbook);
        if (!styles.TryGetValue(styleName, out var definition))
        {
            var candidates = styles.Values
                .OrderBy(s => ExcelPaths.Levenshtein(styleName, s.Name))
                .Select(s => PathOf(s.Name))
                .Take(5)
                .ToList();
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"ops[{index}]: no cell style named '{styleName}' is defined in this workbook.",
                candidates.Count > 0
                    ? "Define it first, or apply one of the existing styles (candidates)."
                    : "Define it first: {op:add, type:cellStyle, props:{name:\"" + styleName +
                      "\", numberFormat:\"$#,##0.00\"}}.",
                candidates: candidates.Count > 0 ? candidates : null);
        }

        ApplyDefinition(style, definition);
    }

    /// <summary>Writes a style definition's facets onto a style holder.</summary>
    public static void ApplyDefinition(IXLStyle style, Definition definition)
    {
        if (definition.NumberFormat is { } numberFormat)
        {
            // A named preset (v1.2) resolves to its OOXML code; a literal code
            // passes through unchanged. The stored definition keeps the original
            // string so 'get' reflects exactly what the caller wrote.
            style.NumberFormat.Format = ExcelNumberFormats.Resolve(numberFormat);
        }

        if (definition.Bold is { } bold)
        {
            style.Font.Bold = bold;
        }

        if (definition.Italic is { } italic)
        {
            style.Font.Italic = italic;
        }

        if (definition.Fill is { } fill)
        {
            style.Fill.BackgroundColor = XLColor.FromHtml(fill);
        }

        if (definition.Color is { } color)
        {
            style.Font.FontColor = XLColor.FromHtml(color);
        }

        if (definition.Border is { } border)
        {
            var borderStyle = BorderStyle(border);
            style.Border.OutsideBorder = borderStyle;
        }
    }

    // ----- read / get / remove ------------------------------------------------

    /// <summary>Every named style as agents see it (read --view styles).</summary>
    public static List<object> ListAll(XLWorkbook workbook) =>
        [.. Load(workbook).Values.OrderBy(s => s.Name, StringComparer.Ordinal).Select(Describe)];

    /// <summary>One named style by name; throws invalid_path with candidates on a miss.</summary>
    public static object Get(XLWorkbook workbook, string name)
    {
        var styles = Load(workbook);
        if (styles.TryGetValue(name, out var definition))
        {
            return Describe(definition);
        }

        var candidates = styles.Values
            .OrderBy(s => ExcelPaths.Levenshtein(name, s.Name))
            .Select(s => PathOf(s.Name))
            .Take(5)
            .ToList();
        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No cell style named '{name}' is defined in this workbook.",
            candidates.Count > 0
                ? "Named styles are matched case-insensitively; pick one of the candidates."
                : "This workbook has no named cell styles; add one with " +
                  "{op:add, type:cellStyle, props:{name:\"Currency-Red\", numberFormat:\"$#,##0.00\"}}.",
            candidates: candidates.Count > 0 ? candidates : null);
    }

    /// <summary>Removes a named style from the registry; throws invalid_path on a miss.</summary>
    public static object Remove(XLWorkbook workbook, string name)
    {
        var styles = Load(workbook);
        if (!styles.Remove(name, out var removed))
        {
            var candidates = styles.Values
                .OrderBy(s => ExcelPaths.Levenshtein(name, s.Name))
                .Select(s => PathOf(s.Name))
                .Take(5)
                .ToList();
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No cell style named '{name}' is defined in this workbook.",
                candidates.Count > 0
                    ? "Pick one of the existing style names (candidates)."
                    : "This workbook has no named cell styles to remove.",
                candidates: candidates.Count > 0 ? candidates : null);
        }

        Store(workbook, styles.Values.ToList());
        return new { op = "remove", path = PathOf(removed.Name), removed = "cellStyle", name = removed.Name };
    }

    public static object Describe(Definition definition) => new
    {
        path = PathOf(definition.Name),
        kind = "cellStyle",
        name = definition.Name,
        numberFormat = definition.NumberFormat,
        bold = definition.Bold,
        italic = definition.Italic,
        fill = definition.Fill,
        color = definition.Color,
        border = definition.Border,
    };

    // ----- raw post-save: surface the styles in Excel's gallery ---------------

    /// <summary>
    /// Emits a real <c>cellStyle</c> entry (and its backing <c>cellStyleXfs</c>
    /// format) for every registered named style, so Excel lists them in the
    /// cell-styles gallery. Idempotent: re-running reconciles the set. No-op
    /// when the registry is empty.
    /// </summary>
    public static void MaterializeAfterSave(string file, IReadOnlyCollection<Definition> styles)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var stylesPart = document.WorkbookPart?.WorkbookStylesPart;
        if (stylesPart?.Stylesheet is not { } stylesheet)
        {
            return;
        }

        var cellStyleFormats = stylesheet.CellStyleFormats ??= NewCellStyleFormats();
        var cellStyles = stylesheet.CellStyles ??= NewCellStyles();

        // The very first cellStyleXfs slot is the document default ("Normal").
        if (!cellStyleFormats.Elements<S.CellFormat>().Any())
        {
            cellStyleFormats.AppendChild(new S.CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 });
        }

        if (!cellStyles.Elements<S.CellStyle>().Any(s => s.BuiltinId?.Value == 0))
        {
            cellStyles.AppendChild(new S.CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 });
        }

        var wanted = styles.Select(s => s.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);

        // Drop any of our previously-emitted named styles that no longer exist,
        // so a remove + re-materialize stays consistent (builtins are untouched).
        foreach (var existing in cellStyles.Elements<S.CellStyle>()
                     .Where(s => s.BuiltinId is null && s.Name?.Value is { } n && !wanted.Contains(n))
                     .ToList())
        {
            existing.Remove();
        }

        var present = cellStyles.Elements<S.CellStyle>()
            .Where(s => s.Name?.Value is not null)
            .Select(s => s.Name!.Value!)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var style in styles)
        {
            if (present.Contains(style.Name))
            {
                continue;
            }

            // One cellStyleXfs slot per named style; we keep its format minimal
            // (number format only) — the concrete cell formatting was already
            // written to the cells, so the gallery link is what matters here.
            var formatId = (uint)cellStyleFormats.Elements<S.CellFormat>().Count();
            var numberFormatId = style.NumberFormat is { } fmt
                ? EnsureNumberFormat(stylesheet, ExcelNumberFormats.Resolve(fmt))
                : 0u;
            var cellFormat = new S.CellFormat { NumberFormatId = numberFormatId, FontId = 0, FillId = 0, BorderId = 0 };
            if (numberFormatId != 0)
            {
                cellFormat.ApplyNumberFormat = true;
            }

            cellStyleFormats.AppendChild(cellFormat);
            cellStyles.AppendChild(new S.CellStyle { Name = style.Name, FormatId = formatId });
        }

        cellStyleFormats.Count = (uint)cellStyleFormats.Elements<S.CellFormat>().Count();
        cellStyles.Count = (uint)cellStyles.Elements<S.CellStyle>().Count();
        stylesheet.Save();
    }

    /// <summary>Finds (or creates) a custom number-format id for a format code.</summary>
    private static uint EnsureNumberFormat(S.Stylesheet stylesheet, string formatCode)
    {
        var numberingFormats = stylesheet.NumberingFormats ??=
            stylesheet.InsertAt(new S.NumberingFormats { Count = 0 }, 0);

        var existing = numberingFormats.Elements<S.NumberingFormat>()
            .FirstOrDefault(n => string.Equals(n.FormatCode?.Value, formatCode, StringComparison.Ordinal));
        if (existing?.NumberFormatId?.Value is { } id)
        {
            return id;
        }

        var nextId = numberingFormats.Elements<S.NumberingFormat>()
            .Select(n => n.NumberFormatId?.Value ?? 0u)
            .DefaultIfEmpty(163u) // custom number formats start at 164
            .Max() + 1;
        nextId = Math.Max(nextId, 164u);
        numberingFormats.AppendChild(new S.NumberingFormat { NumberFormatId = nextId, FormatCode = formatCode });
        numberingFormats.Count = (uint)numberingFormats.Elements<S.NumberingFormat>().Count();
        return nextId;
    }

    private static S.CellStyleFormats NewCellStyleFormats() => new() { Count = 0 };

    private static S.CellStyles NewCellStyles() => new() { Count = 0 };

    // ----- helpers ------------------------------------------------------------

    private static XLColor ParseColorOrThrow(string html, string prop, int index)
    {
        try
        {
            return XLColor.FromHtml(html);
        }
        catch (Exception exception)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{html}' is not a recognizable {prop} color.",
                "Use hex like #FFEE00 (or #AARRGGBB for alpha).",
                innerException: exception);
        }
    }

    private static XLBorderStyleValues BorderStyle(string name) => name.ToLowerInvariant() switch
    {
        "none" => XLBorderStyleValues.None,
        "thin" => XLBorderStyleValues.Thin,
        "medium" => XLBorderStyleValues.Medium,
        "thick" => XLBorderStyleValues.Thick,
        "dashed" => XLBorderStyleValues.Dashed,
        "dotted" => XLBorderStyleValues.Dotted,
        "double" => XLBorderStyleValues.Double,
        "hair" => XLBorderStyleValues.Hair,
        _ => XLBorderStyleValues.Thin,
    };

    private static string? GetString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static bool? GetBool(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() is JsonValueKind.True or JsonValueKind.False
            ? value.GetValue<bool>()
            : null;
}
