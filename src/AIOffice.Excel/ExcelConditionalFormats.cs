using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The xlsx conditional-formatting layer (ClosedXML native). Five kinds are
/// supported: <c>cellIs</c>, <c>colorScale</c>, <c>dataBar</c>,
/// <c>containsText</c>, <c>iconSet</c>; everything else is a typed
/// <c>unsupported_feature</c>.
///
/// Measured ClosedXML 0.105 quirks this slice corrects in a post-save pass
/// (<see cref="FixUpAfterSave"/>):
/// <list type="bullet">
/// <item>New data bars get a lowercase pairing GUID in both the base
/// <c>cfRule/extLst/x14:id</c> and the worksheet-level <c>x14:cfRule@id</c>;
/// the schema pattern requires uppercase hex, so the validator flags both.
/// The pass uppercases the pair (pairing stays intact). Once fixed on disk,
/// ClosedXML round-trips the ids verbatim.</item>
/// <item>Removing a data bar deletes only the base rule and strands the
/// worksheet-level <c>x14:conditionalFormatting</c> twin; the pass deletes
/// orphaned x14 dataBar rules (other x14-only rule types from foreign
/// producers are left untouched).</item>
/// </list>
/// </summary>
internal static partial class ExcelConditionalFormats
{
    /// <summary>The conditional-format kinds aioffice can create.</summary>
    public static readonly IReadOnlyList<string> Kinds =
        ["cellIs", "colorScale", "dataBar", "containsText", "iconSet"];

    private static readonly IReadOnlyList<string> Operators = [">", "<", ">=", "<=", "==", "!=", "between"];

    private static readonly IReadOnlyList<string> CellIsProps =
        ["kind", "operator", "value", "value2", "fill", "color", "bold"];

    private static readonly IReadOnlyList<string> ColorScaleProps = ["kind", "minColor", "midColor", "maxColor"];

    private static readonly IReadOnlyList<string> DataBarProps = ["kind", "color"];

    private static readonly IReadOnlyList<string> ContainsTextProps = ["kind", "text", "fill", "color", "bold"];

    private static readonly IReadOnlyList<string> IconSetProps = ["kind", "set", "reverse", "showValue"];

    /// <summary>
    /// The icon-set names aioffice accepts, in OOXML spelling, mapped to the
    /// ClosedXML style and the icon count (3/4/5). The wire name IS the OOXML
    /// <c>iconSet</c> attribute value, so agents and the file speak the same set
    /// vocabulary. Insertion order is the order the supported-set list reports.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, (XLIconSetStyle Style, int Count)> IconSets =
        new Dictionary<string, (XLIconSetStyle, int)>(StringComparer.Ordinal)
        {
            ["3Arrows"] = (XLIconSetStyle.ThreeArrows, 3),
            ["3ArrowsGray"] = (XLIconSetStyle.ThreeArrowsGray, 3),
            ["3Flags"] = (XLIconSetStyle.ThreeFlags, 3),
            ["3TrafficLights1"] = (XLIconSetStyle.ThreeTrafficLights1, 3),
            ["3TrafficLights2"] = (XLIconSetStyle.ThreeTrafficLights2, 3),
            ["3Signs"] = (XLIconSetStyle.ThreeSigns, 3),
            ["3Symbols"] = (XLIconSetStyle.ThreeSymbols, 3),
            ["3Symbols2"] = (XLIconSetStyle.ThreeSymbols2, 3),
            ["4Arrows"] = (XLIconSetStyle.FourArrows, 4),
            ["4ArrowsGray"] = (XLIconSetStyle.FourArrowsGray, 4),
            ["4RedToBlack"] = (XLIconSetStyle.FourRedToBlack, 4),
            ["4Rating"] = (XLIconSetStyle.FourRating, 4),
            ["4TrafficLights"] = (XLIconSetStyle.FourTrafficLights, 4),
            ["5Arrows"] = (XLIconSetStyle.FiveArrows, 5),
            ["5ArrowsGray"] = (XLIconSetStyle.FiveArrowsGray, 5),
            ["5Rating"] = (XLIconSetStyle.FiveRating, 5),
            ["5Quarters"] = (XLIconSetStyle.FiveQuarters, 5),
        };

    [GeneratedRegex("^#?(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$")]
    private static partial Regex HexColor();

    // ----- add ---------------------------------------------------------------

    /// <summary>
    /// Validates and applies an <c>add conditionalFormat</c> op on the in-memory
    /// workbook. Returns the details entry for the envelope.
    /// </summary>
    public static object Add(ExcelTarget target, EditOp op, int opIndex)
    {
        var range = target.Kind switch
        {
            ExcelTargetKind.Range => target.Range!,
            ExcelTargetKind.Cell => target.Cell!.AsRange(),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add conditionalFormat targets a range path like /Sheet1/A1:C10.",
                "Address the cells the rule should color, e.g. {op:add, type:conditionalFormat, " +
                "path:/Sheet1/A1:C10, props:{kind:\"cellIs\", operator:\">\", value:100, fill:\"FFC7CE\"}}."),
        };

        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add conditionalFormat needs props.",
                "Pass props like {\"kind\":\"cellIs\",\"operator\":\">\",\"value\":100,\"fill\":\"FFC7CE\"}.");
        }

        var kind = OptionalString(props, "kind") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add conditionalFormat needs a 'kind'.",
            "Supported kinds: " + string.Join(", ", Kinds) + ".",
            candidates: Kinds);

        switch (kind)
        {
            case "cellIs":
                AddCellIs(range, props, opIndex);
                break;
            case "colorScale":
                AddColorScale(range, props, opIndex);
                break;
            case "dataBar":
                AddDataBar(range, props, opIndex);
                break;
            case "containsText":
                AddContainsText(range, props, opIndex);
                break;
            case "iconSet":
                AddIconSet(range, props, opIndex);
                break;
            default:
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"ops[{opIndex}]: conditionalFormat kind '{kind}' is not supported yet.",
                    "Supported kinds: cellIs, colorScale, dataBar, containsText, iconSet. For " +
                    "formula-based rules, a colorScale or cellIs rule is the usual stand-in.",
                    candidates: Kinds);
        }

        var index = target.Sheet.ConditionalFormats.Count();
        return new
        {
            op = "add",
            type = "conditionalFormat",
            path = ExcelPaths.ConditionalFormatPath(target.Sheet, index),
            kind,
            range = range.RangeAddress.ToString(),
        };
    }

    private static void AddCellIs(IXLRange range, JsonObject props, int opIndex)
    {
        GuardProps(props, CellIsProps, opIndex);
        var op = OptionalString(props, "operator") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: cellIs needs an 'operator'.",
            "Supported operators: " + string.Join(", ", Operators) + ".",
            candidates: Operators);
        if (!Operators.Contains(op, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: unknown cellIs operator '{op}'.",
                "Supported operators: " + string.Join(", ", Operators) + ".",
                candidates: Operators);
        }

        var value = RequiredNumber(props, "value", opIndex);
        var value2 = OptionalNumber(props, "value2", opIndex);
        if (op == "between" && value2 is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: operator 'between' needs both value and value2.",
                "Pass e.g. {\"operator\":\"between\",\"value\":10,\"value2\":20}.");
        }

        if (op != "between" && value2 is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'value2' only applies to the between operator.",
                "Drop value2, or switch the operator to between.");
        }

        var format = range.AddConditionalFormat();
        var style = op switch
        {
            ">" => format.WhenGreaterThan(value),
            "<" => format.WhenLessThan(value),
            ">=" => format.WhenEqualOrGreaterThan(value),
            "<=" => format.WhenEqualOrLessThan(value),
            "==" => format.WhenEquals(value),
            "!=" => format.WhenNotEquals(value),
            _ => format.WhenBetween(value, value2!.Value),
        };
        ApplyStyle(style, props, opIndex);
    }

    private static void AddColorScale(IXLRange range, JsonObject props, int opIndex)
    {
        GuardProps(props, ColorScaleProps, opIndex);
        var min = ParseColor(RequiredColorText(props, "minColor", opIndex), opIndex);
        var max = ParseColor(RequiredColorText(props, "maxColor", opIndex), opIndex);
        var midText = OptionalString(props, "midColor");

        var scale = range.AddConditionalFormat().ColorScale().LowestValue(min);
        if (midText is null)
        {
            scale.HighestValue(max);
        }
        else
        {
            scale
                .Midpoint(XLCFContentType.Percentile, 50, ParseColor(midText, opIndex))
                .HighestValue(max);
        }
    }

    private static void AddDataBar(IXLRange range, JsonObject props, int opIndex)
    {
        GuardProps(props, DataBarProps, opIndex);
        var color = ParseColor(RequiredColorText(props, "color", opIndex), opIndex);
        range.AddConditionalFormat().DataBar(color, showBarOnly: false).LowestValue().HighestValue();
    }

    private static void AddContainsText(IXLRange range, JsonObject props, int opIndex)
    {
        GuardProps(props, ContainsTextProps, opIndex);
        var text = OptionalString(props, "text");
        if (string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: containsText needs a non-empty 'text'.",
                "Pass e.g. {\"kind\":\"containsText\",\"text\":\"overdue\",\"fill\":\"FFC7CE\"}.");
        }

        ApplyStyle(range.AddConditionalFormat().WhenContains(text), props, opIndex);
    }

    /// <summary>
    /// Adds an iconSet rule: each cell gets one of N icons (3, 4 or 5 depending
    /// on the set) chosen by where its value falls against evenly-spaced percent
    /// thresholds. The first threshold is always 0% (covers the lowest band);
    /// the rest split the 0..100 range evenly (3 icons → 0/33/67, 4 → 0/25/50/75,
    /// 5 → 0/20/40/60/80). <c>reverse</c> flips icon order; <c>showValue</c>
    /// (default true) controls whether the cell value stays visible.
    /// </summary>
    private static void AddIconSet(IXLRange range, JsonObject props, int opIndex)
    {
        GuardProps(props, IconSetProps, opIndex);
        var setName = OptionalString(props, "set") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: iconSet needs a 'set'.",
            "Supported sets: " + string.Join(", ", IconSets.Keys) + ".",
            candidates: [.. IconSets.Keys]);

        if (!IconSets.TryGetValue(setName, out var icon))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{opIndex}]: iconSet '{setName}' is not supported.",
                "Supported sets: " + string.Join(", ", IconSets.Keys) +
                ". The leading digit is the icon count (3, 4 or 5).",
                candidates: [.. IconSets.Keys]);
        }

        var reverse = OptionalBool(props, "reverse") ?? false;
        var showValue = OptionalBool(props, "showValue") ?? true;

        var rule = range.AddConditionalFormat().IconSet(icon.Style, reverse, !showValue);
        foreach (var threshold in EvenPercentThresholds(icon.Count))
        {
            rule.AddValue(
                XLCFIconSetOperator.EqualOrGreaterThan,
                threshold.ToString(CultureInfo.InvariantCulture),
                XLCFContentType.Percent);
        }
    }

    /// <summary>Evenly-spaced percent thresholds starting at 0 (one per icon).</summary>
    private static IEnumerable<int> EvenPercentThresholds(int iconCount)
    {
        for (var i = 0; i < iconCount; i++)
        {
            yield return (int)Math.Round(i * 100.0 / iconCount, MidpointRounding.AwayFromZero);
        }
    }

    /// <summary>Applies fill/color/bold to a rule style; at least one is required.</summary>
    private static void ApplyStyle(IXLStyle style, JsonObject props, int opIndex)
    {
        var applied = false;
        if (OptionalString(props, "fill") is { } fill)
        {
            style.Fill.BackgroundColor = ParseColor(fill, opIndex);
            applied = true;
        }

        if (OptionalString(props, "color") is { } color)
        {
            style.Font.FontColor = ParseColor(color, opIndex);
            applied = true;
        }

        if (props.TryGetPropertyValue("bold", out var boldNode) && boldNode is JsonValue boldValue &&
            boldValue.GetValueKind() is JsonValueKind.True or JsonValueKind.False)
        {
            style.Font.Bold = boldValue.GetValue<bool>();
            applied = true;
        }

        if (!applied)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: the rule has no visible effect; pass at least one of fill, color, bold.",
                "Example: {\"fill\":\"FFC7CE\"} colors matching cells red.");
        }
    }

    // ----- find / describe ----------------------------------------------------

    /// <summary>
    /// Finds the conditional format a resolved target addresses (1-based, in
    /// the sheet's priority order, which is creation order). Throws
    /// <c>invalid_path</c> with the sheet's actual rule paths as candidates.
    /// </summary>
    public static IXLConditionalFormat Find(ExcelTarget target)
    {
        var formats = target.Sheet.ConditionalFormats.ToList();
        var index = target.ConditionalFormatIndex!.Value;
        if (index > formats.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No conditionalFormat[{index}] on sheet '{target.Sheet.Name}' ({formats.Count} rule(s) exist).",
                formats.Count > 0
                    ? "Rule indices are 1-based per sheet in priority order; run 'aioffice read --view structure' to list them."
                    : "This sheet has no conditional formats; add one with {op:add, type:conditionalFormat, " +
                      "path:" + ExcelPaths.SheetPath(target.Sheet) + "/A1:C10, props:{kind:\"cellIs\", …}}.",
                candidates: formats.Count > 0
                    ? [.. Enumerable.Range(1, formats.Count).Select(i => ExcelPaths.ConditionalFormatPath(target.Sheet, i))]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return formats[index - 1];
    }

    /// <summary>One rule as agents see it (get and read --view structure).</summary>
    public static object Describe(IXLWorksheet sheet, IXLConditionalFormat format, int index)
    {
        var kind = KindName(format.ConditionalFormatType);
        return new
        {
            path = ExcelPaths.ConditionalFormatPath(sheet, index),
            kind = "conditionalFormat",
            sheet = sheet.Name,
            cfKind = kind,
            ranges = format.Ranges.Select(r => r.RangeAddress.ToString()).ToList(),
            @operator = format.ConditionalFormatType == XLConditionalFormatType.CellIs
                ? OperatorName(format.Operator)
                : null,
            value = ValueAt(format, 1),
            value2 = format.ConditionalFormatType == XLConditionalFormatType.CellIs ? ValueAt(format, 2) : null,
            text = format.ConditionalFormatType == XLConditionalFormatType.ContainsText ? ValueAt(format, 1) : null,
            fill = HexOf(format.Style.Fill.BackgroundColor),
            color = format.ConditionalFormatType == XLConditionalFormatType.DataBar
                ? ColorAt(format, 1)
                : FontHexOf(format.Style),
            bold = format.Style.Font.Bold ? true : (bool?)null,
            minColor = format.ConditionalFormatType == XLConditionalFormatType.ColorScale ? ColorAt(format, 1) : null,
            midColor = format.ConditionalFormatType == XLConditionalFormatType.ColorScale && format.Colors.Count == 3
                ? ColorAt(format, 2)
                : null,
            maxColor = format.ConditionalFormatType == XLConditionalFormatType.ColorScale
                ? ColorAt(format, format.Colors.Count)
                : null,
            set = format.ConditionalFormatType == XLConditionalFormatType.IconSet
                ? IconSetName(format.IconSetStyle)
                : null,
            reverse = format.ConditionalFormatType == XLConditionalFormatType.IconSet && format.ReverseIconOrder
                ? true
                : (bool?)null,
            showValue = format.ConditionalFormatType == XLConditionalFormatType.IconSet
                ? !format.ShowIconOnly
                : (bool?)null,
        };
    }

    /// <summary>Wire name of a rule type (the five supported kinds keep their op-prop spelling).</summary>
    public static string KindName(XLConditionalFormatType type) => type switch
    {
        XLConditionalFormatType.CellIs => "cellIs",
        XLConditionalFormatType.ColorScale => "colorScale",
        XLConditionalFormatType.DataBar => "dataBar",
        XLConditionalFormatType.ContainsText => "containsText",
        XLConditionalFormatType.IconSet => "iconSet",
        _ => char.ToLowerInvariant(type.ToString()[0]) + type.ToString()[1..],
    };

    /// <summary>Wire <c>set</c> name for a ClosedXML icon-set style (null for unmapped styles).</summary>
    private static string? IconSetName(XLIconSetStyle style)
    {
        foreach (var (name, value) in IconSets)
        {
            if (value.Style == style)
            {
                return name;
            }
        }

        return null;
    }

    private static string? OperatorName(XLCFOperator op) => op switch
    {
        XLCFOperator.GreaterThan => ">",
        XLCFOperator.LessThan => "<",
        XLCFOperator.EqualOrGreaterThan => ">=",
        XLCFOperator.EqualOrLessThan => "<=",
        XLCFOperator.Equal => "==",
        XLCFOperator.NotEqual => "!=",
        XLCFOperator.Between => "between",
        _ => char.ToLowerInvariant(op.ToString()[0]) + op.ToString()[1..],
    };

    private static string? ValueAt(IXLConditionalFormat format, int key) =>
        format.Values.TryGetValue(key, out var formula) ? formula?.Value : null;

    private static string? ColorAt(IXLConditionalFormat format, int key) =>
        format.Colors.TryGetValue(key, out var color) ? HexOf(color) : null;

    /// <summary>Six-digit RGB hex (8-digit when alpha is not FF); null for theme/indexed/unset colors.</summary>
    internal static string? HexOf(XLColor? color)
    {
        if (color is null || !color.HasValue || color.ColorType != XLColorType.Color)
        {
            return null;
        }

        var argb = (uint)color.Color.ToArgb();
        return (argb >> 24) == 0xFF
            ? (argb & 0xFFFFFF).ToString("X6", CultureInfo.InvariantCulture)
            : argb.ToString("X8", CultureInfo.InvariantCulture);
    }

    private static string? FontHexOf(IXLStyle style)
    {
        var hex = HexOf(style.Font.FontColor);
        return hex == "000000" ? null : hex; // the default font color is noise to agents
    }

    // ----- post-save fix-up -----------------------------------------------------

    /// <summary>
    /// Corrects the two measured ClosedXML data-bar defects on the saved file
    /// (lowercase pairing GUIDs; orphaned x14 twins after removal). Cheap scan;
    /// only writes parts that actually change.
    /// </summary>
    public static void FixUpAfterSave(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return;
        }

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            if (worksheetPart.Worksheet is not { } worksheet)
            {
                continue;
            }

            var dirty = false;

            // 1) Uppercase the base-rule pairing ids (cfRule/extLst/x14:id).
            var baseIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var id in worksheet.Descendants<S.ConditionalFormattingRule>()
                         .SelectMany(rule => rule.Descendants<X14.Id>()))
            {
                var upper = id.Text.ToUpperInvariant();
                if (!string.Equals(upper, id.Text, StringComparison.Ordinal))
                {
                    id.Text = upper;
                    dirty = true;
                }

                baseIds.Add(upper);
            }

            // 2) Worksheet-level x14 twins: uppercase matched ids, drop orphaned dataBars.
            foreach (var x14Rule in worksheet.Descendants<X14.ConditionalFormattingRule>().ToList())
            {
                if (x14Rule.Id?.Value is not { } id)
                {
                    continue;
                }

                var upper = id.ToUpperInvariant();
                if (baseIds.Contains(upper))
                {
                    if (!string.Equals(upper, id, StringComparison.Ordinal))
                    {
                        x14Rule.Id = upper;
                        dirty = true;
                    }
                }
                else if (x14Rule.Type?.Value == S.ConditionalFormatValues.DataBar)
                {
                    var parent = x14Rule.Parent;
                    x14Rule.Remove();
                    if (parent is X14.ConditionalFormatting twin &&
                        !twin.Elements<X14.ConditionalFormattingRule>().Any())
                    {
                        twin.Remove();
                    }

                    dirty = true;
                }
            }

            // 3) Remove now-empty x14 containers so the part stays minimal.
            foreach (var container in worksheet.Descendants<X14.ConditionalFormattings>().ToList())
            {
                if (container.HasChildren)
                {
                    continue;
                }

                var extension = container.Parent;
                container.Remove();
                if (extension is S.WorksheetExtension { HasChildren: false } ext)
                {
                    var list = ext.Parent;
                    ext.Remove();
                    if (list is S.WorksheetExtensionList { HasChildren: false } extList)
                    {
                        extList.Remove();
                    }
                }

                dirty = true;
            }

            if (dirty)
            {
                worksheet.Save();
            }
        }
    }

    // ----- small helpers ----------------------------------------------------------

    private static void GuardProps(JsonObject props, IReadOnlyList<string> allowed, int opIndex)
    {
        foreach (var (key, _) in props)
        {
            if (!allowed.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown conditionalFormat prop '{key}' for this kind.",
                    "Supported props here: " + string.Join(", ", allowed) + ".",
                    candidates: allowed);
            }
        }
    }

    private static XLColor ParseColor(string text, int opIndex)
    {
        if (!HexColor().IsMatch(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{text}' is not a usable color.",
                "Pass RGB hex like FFC7CE (a leading # and an AARRGGBB alpha form are also accepted).");
        }

        return XLColor.FromHtml(text.StartsWith('#') ? text : "#" + text);
    }

    private static string RequiredColorText(JsonObject props, string key, int opIndex) =>
        OptionalString(props, key) ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: this conditionalFormat kind needs the '{key}' prop.",
            $"Pass RGB hex, e.g. {{\"{key}\":\"63BE7B\"}}.");

    private static double RequiredNumber(JsonObject props, string key, int opIndex) =>
        OptionalNumber(props, key, opIndex) ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: cellIs needs a numeric '{key}'.",
            $"Pass e.g. {{\"{key}\":100}}.");

    private static double? OptionalNumber(JsonObject props, string key, int opIndex)
    {
        if (!props.TryGetPropertyValue(key, out var node) || node is null)
        {
            return null;
        }

        if (node is JsonValue value)
        {
            if (value.GetValueKind() == JsonValueKind.Number)
            {
                // JsonValue may wrap a CLR int/long/decimal rather than a parsed
                // JSON token; try the numeric types instead of assuming double.
                if (value.TryGetValue<double>(out var asDouble))
                {
                    return asDouble;
                }

                if (value.TryGetValue<int>(out var asInt))
                {
                    return asInt;
                }

                if (value.TryGetValue<long>(out var asLong))
                {
                    return asLong;
                }

                if (value.TryGetValue<float>(out var asFloat))
                {
                    return asFloat;
                }

                if (value.TryGetValue<decimal>(out var asDecimal))
                {
                    return (double)asDecimal;
                }
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
            $"Pass e.g. {{\"{key}\":100}}.");
    }

    private static string? OptionalString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static bool? OptionalBool(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() is JsonValueKind.True or JsonValueKind.False
            ? value.GetValue<bool>()
            : null;
}
