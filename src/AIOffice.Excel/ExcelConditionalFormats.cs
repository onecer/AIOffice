using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The xlsx conditional-formatting layer (ClosedXML native). Eight kinds are
/// supported: <c>cellIs</c>, <c>colorScale</c>, <c>dataBar</c>,
/// <c>containsText</c>, <c>iconSet</c>, and (v1.3, additive) the rule-family
/// completers <c>formula</c> (an <c>=expression</c> rule), <c>topBottom</c>
/// (top/bottom N or N%) and <c>aboveBelowAverage</c> (above/below the range
/// average, optionally ±N standard deviations), plus (v1.21, additive) the
/// value-occurrence pair <c>duplicateValues</c>/<c>uniqueValues</c>; everything
/// else is a typed <c>unsupported_feature</c>.
///
/// Two of the v1.3 kinds save natively through ClosedXML (<c>formula</c> →
/// <c>cfRule@type=expression</c>, <c>topBottom</c> → <c>cfRule@type=top10</c>),
/// but ClosedXML 0.105 throws on the <c>aboveAverage</c> rule type, so
/// <c>aboveBelowAverage</c> rules are authored entirely raw in the post-save
/// pass (<see cref="AverageRuleSpec"/> / <see cref="ApplyAverageRules"/>),
/// mirroring the data-bar fix-up discipline.
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
    [
        "cellIs", "colorScale", "dataBar", "containsText", "iconSet", "formula", "topBottom", "aboveBelowAverage",
        "duplicateValues", "uniqueValues",
    ];

    private static readonly IReadOnlyList<string> Operators = [">", "<", ">=", "<=", "==", "!=", "between"];

    /// <summary>topBottom modes: highest-ranked N (top) or lowest-ranked N (bottom).</summary>
    private static readonly IReadOnlyList<string> TopBottomModes = ["top", "bottom"];

    /// <summary>aboveBelowAverage modes (strict and inclusive of the mean).</summary>
    private static readonly IReadOnlyList<string> AverageModes = ["above", "below", "aboveOrEqual", "belowOrEqual"];

    private static readonly IReadOnlyList<string> CellIsProps =
        ["kind", "operator", "value", "value2", "fill", "color", "bold"];

    private static readonly IReadOnlyList<string> ColorScaleProps =
        ["kind", "minColor", "midColor", "maxColor", "midType", "midValue"];

    /// <summary>
    /// The colorScale 3-color midpoint content types a caller can pin. The wire
    /// name maps to the OOXML <c>cfvo@type</c>: <c>num</c> → a fixed number (legal
    /// even outside the data range, like Excel), <c>percent</c>/<c>percentile</c>
    /// → 0..100 of the range. Omitting both midType and midValue keeps the legacy
    /// Percentile/50 midpoint.
    /// </summary>
    private static readonly IReadOnlyList<string> ColorScaleMidTypes = ["num", "percent", "percentile"];

    private static readonly IReadOnlyList<string> DataBarProps =
        ["kind", "color", "minType", "minValue", "maxType", "maxValue", "showValue"];

    /// <summary>
    /// The data-bar threshold content types a caller can set on the min/max
    /// endpoint. <c>auto</c> keeps Excel's automatic lowest/highest scaling
    /// (the v1.14 default); the rest pin the endpoint to a fixed number, a
    /// percent/percentile of the range, or a formula. The wire name IS the OOXML
    /// <c>cfvo@type</c> spelling for the pinned types (fixed → <c>num</c>).
    /// </summary>
    private static readonly IReadOnlyList<string> DataBarBoundTypes =
        ["auto", "fixed", "percent", "percentile", "formula"];

    private static readonly IReadOnlyList<string> ContainsTextProps = ["kind", "text", "fill", "color", "bold"];

    private static readonly IReadOnlyList<string> IconSetProps = ["kind", "set", "reverse", "showValue", "thresholds"];

    /// <summary>
    /// The icon-set per-icon threshold content types a caller can pin. The wire
    /// name IS the OOXML <c>cfvo@type</c> spelling for the pinned types (fixed →
    /// <c>num</c> stays the colorScale spelling, here re-used as <c>num</c>).
    /// Omitting the whole <c>thresholds</c> array keeps the byte-stable
    /// evenly-spaced percent split.
    /// </summary>
    private static readonly IReadOnlyList<string> IconSetThresholdTypes =
        ["percent", "num", "percentile", "formula"];

    private static readonly IReadOnlyList<string> FormulaProps = ["kind", "formula", "fill", "color", "bold"];

    private static readonly IReadOnlyList<string> TopBottomProps =
        ["kind", "mode", "rank", "percent", "fill", "color", "bold"];

    private static readonly IReadOnlyList<string> AboveBelowAverageProps =
        ["kind", "mode", "stdDev", "fill", "color", "bold"];

    /// <summary>
    /// duplicateValues/uniqueValues take only the shared styling props: the rule
    /// condition is the kind itself (no operator/value/text/formula).
    /// </summary>
    private static readonly IReadOnlyList<string> DuplicateUniqueProps = ["kind", "fill", "color", "bold"];

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
    /// workbook. Returns the details entry for the envelope. The
    /// <c>aboveBelowAverage</c> kind cannot be saved by ClosedXML, so it is
    /// validated here and queued on <paramref name="averageRules"/> for the raw
    /// post-save authoring pass (<see cref="ApplyAverageRules"/>).
    /// </summary>
    public static object Add(
        ExcelTarget target,
        EditOp op,
        int opIndex,
        List<AverageRuleSpec> averageRules,
        List<DataBarThresholdSpec> dataBarRules,
        List<IconSetThresholdSpec> iconSetRules)
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
                AddDataBar(target.Sheet.Name, range, props, opIndex, dataBarRules);
                break;
            case "containsText":
                AddContainsText(range, props, opIndex);
                break;
            case "iconSet":
                AddIconSet(target.Sheet.Name, range, props, opIndex, iconSetRules);
                break;
            case "formula":
                AddFormula(range, props, opIndex);
                break;
            case "topBottom":
                AddTopBottom(range, props, opIndex);
                break;
            case "aboveBelowAverage":
                // Queued for the raw post-save pass (ClosedXML 0.105 cannot save
                // this rule type). The index it will occupy = the rules ClosedXML
                // already holds on the sheet, plus any average rules queued before
                // it on the same sheet, plus one.
                averageRules.Add(ParseAverageRule(target.Sheet.Name, range, props, opIndex));
                break;
            case "duplicateValues":
                AddDuplicateUnique(range, props, opIndex, duplicate: true);
                break;
            case "uniqueValues":
                AddDuplicateUnique(range, props, opIndex, duplicate: false);
                break;
            default:
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"ops[{opIndex}]: conditionalFormat kind '{kind}' is not supported yet.",
                    "Supported kinds: " + string.Join(", ", Kinds) + ".",
                    candidates: Kinds);
        }

        // Index discipline: native rules keep their creation-order positions and
        // the raw average rules are appended after every native rule on the
        // sheet. A native kind's 1-based index is therefore its position among
        // the native rules (which now include it). An average rule's index is
        // every native rule plus its position among the queued average rules
        // (which now include it).
        var nativeCount = target.Sheet.ConditionalFormats.Count();
        var index = kind == "aboveBelowAverage"
            ? nativeCount + averageRules.Count(r =>
                string.Equals(r.SheetName, target.Sheet.Name, StringComparison.OrdinalIgnoreCase))
            : nativeCount;
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

        // The optional caller-controlled midpoint endpoint. Null keeps the legacy
        // Percentile/50 midpoint (byte-stable). midType/midValue only apply to a
        // 3-color scale, so a midpoint type without a midColor is invalid.
        var midpoint = ParseColorScaleMidpoint(props, midText is not null, opIndex);

        var scale = range.AddConditionalFormat().ColorScale().LowestValue(min);
        if (midText is null)
        {
            scale.HighestValue(max);
        }
        else if (midpoint is null)
        {
            scale
                .Midpoint(XLCFContentType.Percentile, 50, ParseColor(midText, opIndex))
                .HighestValue(max);
        }
        else
        {
            scale
                .Midpoint(midpoint.Value.Type, midpoint.Value.Value, ParseColor(midText, opIndex))
                .HighestValue(max);
        }
    }

    /// <summary>
    /// Parses (and validates) the optional colorScale midpoint endpoint. Returns
    /// null when BOTH <c>midType</c> and <c>midValue</c> are absent (the legacy
    /// Percentile/50 default stays byte-stable). When present, both are required,
    /// a midColor must be present (a midpoint type needs a 3-color scale), and a
    /// <c>percent</c>/<c>percentile</c> value is clamped to 0..100 (a <c>num</c>
    /// value outside the data range is legal in Excel — accepted as-is).
    /// </summary>
    private static (XLCFContentType Type, double Value)? ParseColorScaleMidpoint(
        JsonObject props, bool hasMidColor, int opIndex)
    {
        var hasMidType = props.ContainsKey("midType");
        var hasMidValue = props.ContainsKey("midValue");
        if (!hasMidType && !hasMidValue)
        {
            return null;
        }

        if (!hasMidColor)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a midpoint type needs a midColor (midType/midValue apply to a 3-color scale).",
                "Add a 'midColor', or drop midType/midValue for a 2-color scale.");
        }

        var midType = OptionalString(props, "midType");
        if (midType is null)
        {
            // midValue supplied without midType (the only way to reach here with a null
            // type, since absent-both returned above) — name the missing prop, not "".
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a 'midValue' needs a 'midType'.",
                "Pass a midType too, e.g. {\"midType\":\"percent\",\"midValue\":50}. Supported: " +
                string.Join(", ", ColorScaleMidTypes) + ".",
                candidates: ColorScaleMidTypes);
        }

        if (!ColorScaleMidTypes.Contains(midType, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: unknown colorScale 'midType' '{midType}'.",
                "Supported midpoint types: " + string.Join(", ", ColorScaleMidTypes) + ".",
                candidates: ColorScaleMidTypes);
        }

        var midValue = OptionalNumber(props, "midValue", opIndex);
        if (midValue is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: midType '{midType}' needs a numeric 'midValue'.",
                "Pass e.g. {\"midType\":\"" + midType + "\",\"midValue\":50}.");
        }

        if ((midType == "percent" || midType == "percentile") && (midValue < 0 || midValue > 100))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a {midType} 'midValue' must be between 0 and 100.",
                "Pass e.g. {\"midType\":\"" + midType + "\",\"midValue\":50}.");
        }

        var type = midType switch
        {
            "num" => XLCFContentType.Number,
            "percent" => XLCFContentType.Percent,
            _ => XLCFContentType.Percentile,
        };
        return (type, midValue.Value);
    }

    /// <summary>
    /// Adds a data-bar rule. By default (no threshold/showValue prop) the bar
    /// auto-scales lowest→highest with the value shown — the byte-stable v1.14
    /// path. When any of <c>minType</c>/<c>maxType</c>/<c>minValue</c>/
    /// <c>maxValue</c>/<c>showValue</c> is supplied, the rule is still created
    /// through ClosedXML (so its base/x14 pairing is valid), then queued for a
    /// raw post-save pass that rewrites the two cfvo (<c>minType/maxType</c>) and
    /// the <c>@showValue</c> — ClosedXML 0.105 cannot author those endpoints.
    /// </summary>
    private static void AddDataBar(
        string sheetName, IXLRange range, JsonObject props, int opIndex, List<DataBarThresholdSpec> dataBarRules)
    {
        GuardProps(props, DataBarProps, opIndex);
        var color = ParseColor(RequiredColorText(props, "color", opIndex), opIndex);

        // The default path is byte-identical to v1.14: auto min/max cfvo, value
        // shown, no raw rewrite. It MUST stay first-in-code and untouched.
        var threshold = ParseDataBarThreshold(sheetName, range, props, opIndex);
        if (threshold is null)
        {
            range.AddConditionalFormat().DataBar(color, showBarOnly: false).LowestValue().HighestValue();
            return;
        }

        // Create the same ClosedXML rule (valid base/x14 twin); the post-save
        // pass rewrites its cfvo + @showValue from the queued spec.
        range.AddConditionalFormat().DataBar(color, showBarOnly: false).LowestValue().HighestValue();
        dataBarRules.Add(threshold);
    }

    /// <summary>
    /// Parses (and validates) the optional caller-controlled data-bar endpoints.
    /// Returns null when no threshold/showValue prop is supplied (the byte-stable
    /// default). When present, each endpoint is an (type, value) pair: <c>auto</c>
    /// carries no value (Excel's lowest/highest), <c>fixed</c>/<c>percent</c>/
    /// <c>percentile</c> carry a number, and <c>formula</c> carries an expression
    /// string. <c>showValue</c> toggles the cell value's visibility (default on).
    /// </summary>
    private static DataBarThresholdSpec? ParseDataBarThreshold(
        string sheetName, IXLRange range, JsonObject props, int opIndex)
    {
        var hasMinType = props.ContainsKey("minType");
        var hasMaxType = props.ContainsKey("maxType");
        var hasMinValue = props.ContainsKey("minValue");
        var hasMaxValue = props.ContainsKey("maxValue");
        var showValue = OptionalBool(props, "showValue");
        if (!hasMinType && !hasMaxType && !hasMinValue && !hasMaxValue && showValue is null)
        {
            return null;
        }

        // A bare value without its type is ambiguous (auto endpoints take no
        // value); require the type so the endpoint is unambiguous.
        if (hasMinValue && !hasMinType)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'minValue' needs a 'minType' (auto, fixed, percent, percentile or formula).",
                "Pass e.g. {\"minType\":\"fixed\",\"minValue\":0}.",
                candidates: DataBarBoundTypes);
        }

        if (hasMaxValue && !hasMaxType)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'maxValue' needs a 'maxType' (auto, fixed, percent, percentile or formula).",
                "Pass e.g. {\"maxType\":\"percentile\",\"maxValue\":90}.",
                candidates: DataBarBoundTypes);
        }

        var min = ParseDataBarBound(props, "minType", "minValue", hasMinType, opIndex);
        var max = ParseDataBarBound(props, "maxType", "maxValue", hasMaxType, opIndex);
        return new DataBarThresholdSpec(
            sheetName, range.RangeAddress.ToString()!, NormalizeArgb(ParseColor(RequiredColorText(props, "color", opIndex), opIndex)), min, max, showValue);
    }

    /// <summary>
    /// Parses one data-bar endpoint (min or max). When the type key is absent the
    /// endpoint keeps Excel's automatic scaling (<see cref="DataBarBound.Auto"/>).
    /// </summary>
    private static DataBarBound ParseDataBarBound(
        JsonObject props, string typeKey, string valueKey, bool hasType, int opIndex)
    {
        if (!hasType)
        {
            // No type on this endpoint: keep the default auto (lowest/highest).
            // A value without a type was already rejected by the caller.
            return DataBarBound.Auto;
        }

        var type = OptionalString(props, typeKey);
        if (type is null || !DataBarBoundTypes.Contains(type, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: unknown data-bar '{typeKey}' '{type}'.",
                "Supported types: " + string.Join(", ", DataBarBoundTypes) + ".",
                candidates: DataBarBoundTypes);
        }

        if (type == "auto")
        {
            if (props.ContainsKey(valueKey))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: '{typeKey}' is auto, so '{valueKey}' must be omitted.",
                    "Drop the value for an auto endpoint, or switch the type to fixed/percent/percentile/formula.");
            }

            return DataBarBound.Auto;
        }

        if (type == "formula")
        {
            var formula = OptionalString(props, valueKey);
            if (string.IsNullOrWhiteSpace(formula))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: '{typeKey}' is formula, so '{valueKey}' must be a non-empty formula string.",
                    "Pass e.g. {\"" + typeKey + "\":\"formula\",\"" + valueKey + "\":\"$A$1\"}.");
            }

            // Excel's cfvo/@val holds the formula WITHOUT the leading '='.
            var expression = formula.StartsWith('=') ? formula[1..] : formula;
            return new DataBarBound(type, expression, null);
        }

        // fixed / percent / percentile: a number.
        var number = OptionalNumber(props, valueKey, opIndex);
        if (number is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{typeKey}' is {type}, so '{valueKey}' must be a number.",
                "Pass e.g. {\"" + typeKey + "\":\"" + type + "\",\"" + valueKey + "\":50}.");
        }

        if ((type == "percent" || type == "percentile") && (number < 0 || number > 100))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a {type} '{valueKey}' must be between 0 and 100.",
                "Pass e.g. {\"" + typeKey + "\":\"" + type + "\",\"" + valueKey + "\":90}.");
        }

        return new DataBarBound(
            type, number.Value.ToString(CultureInfo.InvariantCulture), number.Value);
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
    /// on the set) chosen by where its value falls against the per-icon
    /// thresholds. By default (no <c>thresholds</c> prop) the bands split the
    /// 0..100 range evenly with the first at 0% (3 icons → 0/33/67, 4 →
    /// 0/25/50/75, 5 → 0/20/40/60/80, all <c>@type=percent</c>) — the byte-stable
    /// v1.19 path. When <c>thresholds</c> is supplied (exactly N entries), the
    /// rule is still created through ClosedXML (valid even-split cfvo), then
    /// queued for a raw post-save pass that rewrites the N cfvo in place from the
    /// caller's spec — ClosedXML 0.105 only authors percent thresholds.
    /// <c>reverse</c> flips icon order; <c>showValue</c> (default true) controls
    /// whether the cell value stays visible.
    /// </summary>
    private static void AddIconSet(
        string sheetName, IXLRange range, JsonObject props, int opIndex, List<IconSetThresholdSpec> iconSetRules)
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

        // The caller-controlled per-icon thresholds (null = the byte-stable even
        // split). Parsed/validated up front so a bad spec fails before the rule
        // is created; queued for the post-save rewrite only when supplied.
        var thresholds = ParseIconSetThresholds(sheetName, range, setName, icon.Count, props, opIndex);

        var rule = range.AddConditionalFormat().IconSet(icon.Style, reverse, !showValue);

        // DEFAULT path (no thresholds): the v1.19 evenly-spaced percent cfvo,
        // first-in-code and byte-identical. MUST stay untouched.
        if (thresholds is null)
        {
            foreach (var threshold in EvenPercentThresholds(icon.Count))
            {
                rule.AddValue(
                    XLCFIconSetOperator.EqualOrGreaterThan,
                    threshold.ToString(CultureInfo.InvariantCulture),
                    XLCFContentType.Percent);
            }

            return;
        }

        // Custom path: create the same valid even-split cfvo (so ClosedXML saves
        // a well-formed rule with N cfvo), then queue the in-place rewrite.
        foreach (var threshold in EvenPercentThresholds(icon.Count))
        {
            rule.AddValue(
                XLCFIconSetOperator.EqualOrGreaterThan,
                threshold.ToString(CultureInfo.InvariantCulture),
                XLCFContentType.Percent);
        }

        iconSetRules.Add(thresholds);
    }

    /// <summary>
    /// Parses (and validates) the optional caller-controlled iconSet per-icon
    /// thresholds. Returns null when <c>thresholds</c> is absent (the byte-stable
    /// even split). When present, the array MUST hold exactly <paramref
    /// name="iconCount"/> entries (one cfvo per icon, lowest band first); each is
    /// a {type, value} pair: <c>num</c>/<c>percent</c>/<c>percentile</c> carry a
    /// number, <c>formula</c> carries an expression string. percent/percentile
    /// values are clamped to 0..100.
    /// </summary>
    private static IconSetThresholdSpec? ParseIconSetThresholds(
        string sheetName, IXLRange range, string setName, int iconCount, JsonObject props, int opIndex)
    {
        if (!props.TryGetPropertyValue("thresholds", out var node) || node is null)
        {
            return null;
        }

        if (node is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: iconSet 'thresholds' must be an array of {{type, value}} entries.",
                "Pass one entry per icon, e.g. {\"thresholds\":[{\"type\":\"percent\",\"value\":0}," +
                "{\"type\":\"percent\",\"value\":50},{\"type\":\"percent\",\"value\":90}]}.");
        }

        if (array.Count != iconCount)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: iconSet '{setName}' has {iconCount} icons, so 'thresholds' needs exactly " +
                $"{iconCount} entries (got {array.Count}).",
                $"Pass one {{type, value}} entry per icon (the lowest band first), {iconCount} in total.");
        }

        var bounds = new List<IconSetBound>(iconCount);
        for (var i = 0; i < array.Count; i++)
        {
            if (array[i] is not JsonObject entry)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: iconSet thresholds[{i}] must be a {{type, value}} object.",
                    "Pass e.g. {\"type\":\"percent\",\"value\":50}.");
            }

            bounds.Add(ParseIconSetBound(entry, i, opIndex));
        }

        return new IconSetThresholdSpec(sheetName, range.RangeAddress.ToString()!, setName, bounds);
    }

    /// <summary>Parses one iconSet threshold entry (a cfvo): validates type, value and range.</summary>
    private static IconSetBound ParseIconSetBound(JsonObject entry, int slot, int opIndex)
    {
        var type = OptionalString(entry, "type");
        if (type is null || !IconSetThresholdTypes.Contains(type, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: unknown iconSet threshold 'type' '{type}' at thresholds[{slot}].",
                "Supported types: " + string.Join(", ", IconSetThresholdTypes) + ".",
                candidates: IconSetThresholdTypes);
        }

        if (type == "formula")
        {
            var formula = OptionalString(entry, "value");
            if (string.IsNullOrWhiteSpace(formula))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: iconSet thresholds[{slot}] type is formula, so 'value' must be a non-empty formula string.",
                    "Pass e.g. {\"type\":\"formula\",\"value\":\"=$A$1\"}.");
            }

            // Excel's cfvo/@val holds the formula WITHOUT the leading '='.
            var expression = formula.StartsWith('=') ? formula[1..] : formula;
            return new IconSetBound(type, expression);
        }

        // num / percent / percentile: a number.
        var number = OptionalNumber(entry, "value", opIndex);
        if (number is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: iconSet thresholds[{slot}] type is {type}, so 'value' must be a number.",
                "Pass e.g. {\"type\":\"" + type + "\",\"value\":50}.");
        }

        if ((type == "percent" || type == "percentile") && (number < 0 || number > 100))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a {type} iconSet threshold 'value' must be between 0 and 100.",
                "Pass e.g. {\"type\":\"" + type + "\",\"value\":50}.");
        }

        return new IconSetBound(type, number.Value.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>Evenly-spaced percent thresholds starting at 0 (one per icon).</summary>
    private static IEnumerable<int> EvenPercentThresholds(int iconCount)
    {
        for (var i = 0; i < iconCount; i++)
        {
            yield return (int)Math.Round(i * 100.0 / iconCount, MidpointRounding.AwayFromZero);
        }
    }

    /// <summary>
    /// Adds an expression rule (<c>cfRule@type=expression</c>): the cells whose
    /// formula evaluates truthy get the fill/color/bold. The formula is a real
    /// Excel expression relative to the range's top-left cell (e.g.
    /// <c>=$A1&gt;$B1</c>); a leading <c>=</c> is optional.
    /// </summary>
    private static void AddFormula(IXLRange range, JsonObject props, int opIndex)
    {
        GuardProps(props, FormulaProps, opIndex);
        var formula = OptionalString(props, "formula");
        if (string.IsNullOrWhiteSpace(formula))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: formula needs a non-empty 'formula'.",
                "Pass an Excel expression relative to the range's top-left cell, e.g. " +
                "{\"kind\":\"formula\",\"formula\":\"=$A1>$B1\",\"fill\":\"FFC7CE\"}.");
        }

        // Excel's cfRule/formula holds the expression WITHOUT the leading '=';
        // accept either form from the agent and normalize.
        var expression = formula.StartsWith('=') ? formula[1..] : formula;
        if (expression.Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: the formula is just '='; it has no expression.",
                "Pass an expression after the '=', e.g. {\"formula\":\"=$A1>$B1\"}.");
        }

        ApplyStyle(range.AddConditionalFormat().WhenIsTrue(expression), props, opIndex);
    }

    /// <summary>
    /// Adds a duplicateValues/uniqueValues rule (<c>cfRule@type=duplicateValues</c>
    /// / <c>uniqueValues</c>): the cells whose value occurs more than once
    /// (duplicate) or exactly once (unique) in the range get the fill/color/bold.
    /// The condition IS the kind — no operator/value/text/formula props.
    /// </summary>
    private static void AddDuplicateUnique(IXLRange range, JsonObject props, int opIndex, bool duplicate)
    {
        GuardProps(props, DuplicateUniqueProps, opIndex);
        var format = range.AddConditionalFormat();
        ApplyStyle(duplicate ? format.WhenIsDuplicate() : format.WhenIsUnique(), props, opIndex);
    }

    /// <summary>
    /// Adds a top/bottom rule (<c>cfRule@type=top10</c>): the highest- (mode
    /// "top") or lowest- (mode "bottom") ranked N values in the range get the
    /// fill/color/bold. With <c>percent:true</c>, N is a percentage of the cell
    /// count rather than a count.
    /// </summary>
    private static void AddTopBottom(IXLRange range, JsonObject props, int opIndex)
    {
        GuardProps(props, TopBottomProps, opIndex);
        var mode = OptionalString(props, "mode") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: topBottom needs a 'mode'.",
            "Supported modes: " + string.Join(", ", TopBottomModes) + ".",
            candidates: TopBottomModes);
        if (!TopBottomModes.Contains(mode, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: unknown topBottom mode '{mode}'.",
                "Supported modes: " + string.Join(", ", TopBottomModes) + ".",
                candidates: TopBottomModes);
        }

        var percent = OptionalBool(props, "percent") ?? false;
        var rank = RequiredRank(props, opIndex, percent);

        var style = mode == "top"
            ? range.AddConditionalFormat().WhenIsTop(rank, percent ? XLTopBottomType.Percent : XLTopBottomType.Items)
            : range.AddConditionalFormat().WhenIsBottom(rank, percent ? XLTopBottomType.Percent : XLTopBottomType.Items);
        ApplyStyle(style, props, opIndex);
    }

    /// <summary>
    /// Parses (but does not apply) an aboveBelowAverage rule for the post-save
    /// raw pass. <c>mode</c> is above/below (strict) or aboveOrEqual/belowOrEqual
    /// (inclusive of the mean); the optional <c>stdDev</c> (a positive integer)
    /// shifts the threshold that many standard deviations away from the average.
    /// </summary>
    private static AverageRuleSpec ParseAverageRule(string sheetName, IXLRange range, JsonObject props, int opIndex)
    {
        GuardProps(props, AboveBelowAverageProps, opIndex);
        var mode = OptionalString(props, "mode") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: aboveBelowAverage needs a 'mode'.",
            "Supported modes: " + string.Join(", ", AverageModes) + ".",
            candidates: AverageModes);
        if (!AverageModes.Contains(mode, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: unknown aboveBelowAverage mode '{mode}'.",
                "Supported modes: " + string.Join(", ", AverageModes) + ".",
                candidates: AverageModes);
        }

        var stdDev = OptionalNumber(props, "stdDev", opIndex);
        var stdDevInt = 0;
        if (stdDev is { } sd)
        {
            if (sd < 0 || sd != Math.Floor(sd))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: 'stdDev' must be a non-negative whole number of standard deviations.",
                    "Pass e.g. {\"stdDev\":1} for one standard deviation above/below; omit it for the plain average.");
            }

            stdDevInt = (int)sd;
        }

        var (fill, color, bold) = ParseDifferentialStyle(props, opIndex);
        var above = mode is "above" or "aboveOrEqual";
        var equalAverage = mode is "aboveOrEqual" or "belowOrEqual";
        return new AverageRuleSpec(
            sheetName,
            range.RangeAddress.ToString()!,
            Above: above,
            EqualAverage: equalAverage,
            StdDev: stdDevInt,
            Fill: fill,
            FontColor: color,
            Bold: bold);
    }

    private static int RequiredRank(JsonObject props, int opIndex, bool percent)
    {
        var rank = OptionalNumber(props, "rank", opIndex) ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: topBottom needs a 'rank' (the N in top/bottom N).",
            "Pass e.g. {\"mode\":\"top\",\"rank\":10,\"fill\":\"FFC7CE\"}.");
        if (rank < 1 || rank != Math.Floor(rank))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'rank' must be a whole number of at least 1.",
                "Pass e.g. {\"rank\":10} for the top/bottom 10.");
        }

        // Excel caps the rank at 1000 (items) or 100 (percent).
        var max = percent ? 100 : 1000;
        if (rank > max)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'rank' must be {max} or less for a {(percent ? "percent" : "count")} rule.",
                percent
                    ? "Excel allows a top/bottom percentage of 1..100."
                    : "Excel allows a top/bottom count of 1..1000.");
        }

        return (int)rank;
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

    /// <summary>
    /// Validates fill/color/bold for a raw-authored rule (aboveBelowAverage),
    /// returning the six-digit RGB hex (ARGB-padded) for the dxf. At least one
    /// must be present, mirroring <see cref="ApplyStyle"/>.
    /// </summary>
    private static (string? Fill, string? FontColor, bool? Bold) ParseDifferentialStyle(JsonObject props, int opIndex)
    {
        string? fill = null, color = null;
        bool? bold = null;
        if (OptionalString(props, "fill") is { } fillText)
        {
            fill = NormalizeArgb(ParseColor(fillText, opIndex));
        }

        if (OptionalString(props, "color") is { } colorText)
        {
            color = NormalizeArgb(ParseColor(colorText, opIndex));
        }

        if (props.TryGetPropertyValue("bold", out var boldNode) && boldNode is JsonValue boldValue &&
            boldValue.GetValueKind() is JsonValueKind.True or JsonValueKind.False)
        {
            bold = boldValue.GetValue<bool>();
        }

        if (fill is null && color is null && bold is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: the rule has no visible effect; pass at least one of fill, color, bold.",
                "Example: {\"fill\":\"FFC7CE\"} colors matching cells red.");
        }

        return (fill, color, bold);
    }

    /// <summary>Eight-digit ARGB hex (FF alpha) for a parsed color, for raw dxf authoring.</summary>
    private static string NormalizeArgb(XLColor color)
    {
        var argb = (uint)color.Color.ToArgb();
        return argb.ToString("X8", CultureInfo.InvariantCulture);
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
    public static object Describe(
        IXLWorksheet sheet,
        IXLConditionalFormat format,
        int index,
        AverageRuleDetail? averageDetail = null,
        DataBarThresholdDetail? dataBarDetail = null,
        IconSetThresholdDetail? iconSetDetail = null)
    {
        var kind = KindName(format.ConditionalFormatType);
        var isTop10 = format.ConditionalFormatType == XLConditionalFormatType.Top10;
        var isDataBar = format.ConditionalFormatType == XLConditionalFormatType.DataBar;
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
            // formula (Expression) rules carry their expression in Values[1].
            formula = format.ConditionalFormatType == XLConditionalFormatType.Expression ? ValueAt(format, 1) : null,
            // topBottom (Top10) rules: highest-ranked = top, Bottom flag = bottom;
            // rank is Values[1]; percent reads the Percent flag.
            mode = isTop10
                ? (format.Bottom ? "bottom" : "top")
                : averageDetail?.Mode,
            rank = isTop10 && int.TryParse(ValueAt(format, 1), out var r) ? r : (int?)null,
            percent = isTop10 ? format.Percent : (bool?)null,
            stdDev = averageDetail is { StdDev: > 0 } ? averageDetail.StdDev : (int?)null,
            fill = HexOf(format.Style.Fill.BackgroundColor),
            color = format.ConditionalFormatType == XLConditionalFormatType.DataBar
                ? ColorAt(format, 1)
                : FontHexOf(format.Style),
            bold = format.Style.Font.Bold ? true : (bool?)null,
            minColor = format.ConditionalFormatType == XLConditionalFormatType.ColorScale ? ColorAt(format, 1) : null,
            midColor = format.ConditionalFormatType == XLConditionalFormatType.ColorScale && format.Colors.Count == 3
                ? ColorAt(format, 2)
                : null,
            // The 3-color midpoint endpoint, emitted ONLY when it is NOT the legacy
            // Percentile/50 default, so legacy colorScale rules read unchanged.
            midType = ColorScaleMidpoint(format)?.Type,
            midValue = ColorScaleMidpoint(format)?.Value,
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
                : isDataBar ? dataBarDetail?.ShowValue : (bool?)null,
            // dataBar thresholds (read raw): the caller-controlled endpoints, or
            // auto for a default bar. Null on every non-dataBar kind.
            minType = isDataBar ? dataBarDetail?.MinType : null,
            minValue = isDataBar ? dataBarDetail?.MinValue : null,
            maxType = isDataBar ? dataBarDetail?.MaxType : null,
            maxValue = isDataBar ? dataBarDetail?.MaxValue : null,
            // iconSet per-icon thresholds (read raw): emitted ONLY when the caller
            // pinned them (non-default), so a legacy even-split iconSet projects no
            // 'thresholds' key. Null on every non-iconSet kind.
            thresholds = format.ConditionalFormatType == XLConditionalFormatType.IconSet
                ? iconSetDetail?.Entries
                : null,
        };
    }

    /// <summary>
    /// The above/below-average sub-details ClosedXML does not expose, read raw
    /// from a worksheet's <c>cfRule[@type=aboveAverage]</c> elements (in priority
    /// order). <c>Mode</c> recombines the rule's aboveAverage + equalAverage
    /// flags into the wire mode; <c>StdDev</c> is the standard-deviation count.
    /// </summary>
    public sealed record AverageRuleDetail(string Mode, int StdDev);

    /// <summary>
    /// The data-bar threshold endpoints + showValue read raw from a saved
    /// <c>cfRule[@type=dataBar]</c> (ClosedXML's data-bar reader does not surface
    /// the cfvo content type/value the way <c>get</c> needs). Keyed by 1-based
    /// sheet rule index (priority order). <c>MinType</c>/<c>MaxType</c> are the
    /// wire spelling (auto for min/max cfvo); the matching value is null for auto.
    /// </summary>
    public sealed record DataBarThresholdDetail(
        string MinType, string? MinValue, string MaxType, string? MaxValue, bool ShowValue);

    /// <summary>
    /// One iconSet per-icon threshold as agents see it (the wire {type, value}
    /// pair). <c>Value</c> is the cfvo <c>@val</c> text: a number's invariant
    /// string, or a formula without its leading <c>=</c>.
    /// </summary>
    public sealed record IconSetThresholdEntry(string Type, string Value);

    /// <summary>
    /// The iconSet per-icon thresholds read raw from a saved
    /// <c>cfRule[@type=iconSet]</c>, in document (lowest-band-first) order — but
    /// ONLY when they are NOT the legacy even-split default, so legacy iconSet
    /// rules surface no thresholds. Keyed by 1-based sheet rule index (priority
    /// order). iconSet has no x14 twin, so this reads the base cfvo directly.
    /// </summary>
    public sealed record IconSetThresholdDetail(IReadOnlyList<IconSetThresholdEntry> Entries);

    /// <summary>
    /// Reads the aboveAverage rules on a sheet (priority order) so the get handler
    /// can enrich the ClosedXML-blind <c>aboveBelowAverage</c> describe with mode
    /// and stdDev. Keyed by 1-based sheet rule index (priority order).
    /// </summary>
    public static IReadOnlyDictionary<int, AverageRuleDetail> ReadAverageDetails(
        SpreadsheetDocument document, string sheetName)
    {
        var result = new Dictionary<int, AverageRuleDetail>();
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is not { } relationshipId ||
            workbookPart!.GetPartById(relationshipId) is not WorksheetPart { Worksheet: { } worksheet })
        {
            return result;
        }

        // All cfRules on the sheet, ordered by their priority attribute — the
        // same ordering ClosedXML surfaces as the 1-based rule index.
        var rules = worksheet.Descendants<S.ConditionalFormattingRule>()
            .OrderBy(r => r.Priority?.Value ?? int.MaxValue)
            .ToList();
        for (var i = 0; i < rules.Count; i++)
        {
            var rule = rules[i];
            if (rule.Type?.Value != S.ConditionalFormatValues.AboveAverage)
            {
                continue;
            }

            var above = rule.AboveAverage?.Value ?? true;
            var equal = rule.EqualAverage?.Value ?? false;
            var mode = (above, equal) switch
            {
                (true, false) => "above",
                (true, true) => "aboveOrEqual",
                (false, false) => "below",
                (false, true) => "belowOrEqual",
            };
            result[i + 1] = new AverageRuleDetail(mode, rule.StdDev?.Value ?? 0);
        }

        return result;
    }

    /// <summary>
    /// Reads the data-bar threshold endpoints + showValue on a sheet (priority
    /// order, the same 1-based index ClosedXML surfaces) so the get handler can
    /// report the caller-controlled endpoints. The base cfRule's two cfvo (MIN
    /// then MAX) carry the wire type; <c>num</c> reads back as <c>fixed</c>; auto
    /// (<c>min</c>/<c>max</c>) reports a null value. Keyed by 1-based rule index.
    /// </summary>
    public static IReadOnlyDictionary<int, DataBarThresholdDetail> ReadDataBarThresholds(
        SpreadsheetDocument document, string sheetName)
    {
        var result = new Dictionary<int, DataBarThresholdDetail>();
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is not { } relationshipId ||
            workbookPart!.GetPartById(relationshipId) is not WorksheetPart { Worksheet: { } worksheet })
        {
            return result;
        }

        var rules = worksheet.Descendants<S.ConditionalFormattingRule>()
            .OrderBy(r => r.Priority?.Value ?? int.MaxValue)
            .ToList();
        for (var i = 0; i < rules.Count; i++)
        {
            if (rules[i].GetFirstChild<S.DataBar>() is not { } dataBar)
            {
                continue;
            }

            var cfvos = dataBar.Elements<S.ConditionalFormatValueObject>().ToList();
            var (minType, minValue) = cfvos.Count >= 1 ? ReadCfvo(cfvos[0]) : ("auto", (string?)null);
            var (maxType, maxValue) = cfvos.Count >= 2 ? ReadCfvo(cfvos[1]) : ("auto", (string?)null);
            // ClosedXML always writes showValue="1"; @showValue="0" only when set.
            var showValue = dataBar.ShowValue?.Value ?? true;
            result[i + 1] = new DataBarThresholdDetail(minType, minValue, maxType, maxValue, showValue);
        }

        return result;
    }

    /// <summary>
    /// Reads the iconSet per-icon thresholds on a sheet (priority order, the same
    /// 1-based index ClosedXML surfaces) so the get handler can report the
    /// caller-controlled bands. Each iconSet cfvo carries the wire type
    /// (<c>num</c>/<c>percent</c>/<c>percentile</c>/<c>formula</c>) and value in
    /// document (lowest-band-first) order. A rule whose cfvo are exactly the
    /// even-split default (all <c>percent</c>, values 0/33/67…) reports NO entry,
    /// so legacy iconSet rules read byte-unchanged. Keyed by 1-based rule index.
    /// </summary>
    public static IReadOnlyDictionary<int, IconSetThresholdDetail> ReadIconSetThresholds(
        SpreadsheetDocument document, string sheetName)
    {
        var result = new Dictionary<int, IconSetThresholdDetail>();
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is not { } relationshipId ||
            workbookPart!.GetPartById(relationshipId) is not WorksheetPart { Worksheet: { } worksheet })
        {
            return result;
        }

        var rules = worksheet.Descendants<S.ConditionalFormattingRule>()
            .OrderBy(r => r.Priority?.Value ?? int.MaxValue)
            .ToList();
        for (var i = 0; i < rules.Count; i++)
        {
            if (rules[i].GetFirstChild<S.IconSet>() is not { } iconSet)
            {
                continue;
            }

            var cfvos = iconSet.Elements<S.ConditionalFormatValueObject>().ToList();
            var entries = cfvos.Select(ReadIconSetCfvo).ToList();
            if (IsEvenPercentSplit(entries))
            {
                // The legacy even split is omitted so legacy files read unchanged;
                // only caller-pinned thresholds surface.
                continue;
            }

            result[i + 1] = new IconSetThresholdDetail(entries);
        }

        return result;
    }

    /// <summary>Wire (type, value) of an iconSet cfvo (num/percent/percentile/formula).</summary>
    private static IconSetThresholdEntry ReadIconSetCfvo(S.ConditionalFormatValueObject cfvo)
    {
        var type = cfvo.Type?.Value;
        var wire = type == S.ConditionalFormatValueObjectValues.Number ? "num"
            : type == S.ConditionalFormatValueObjectValues.Percent ? "percent"
            : type == S.ConditionalFormatValueObjectValues.Percentile ? "percentile"
            : type == S.ConditionalFormatValueObjectValues.Formula ? "formula"
            : "percent";
        return new IconSetThresholdEntry(wire, cfvo.Val?.Value ?? string.Empty);
    }

    /// <summary>True when an iconSet's cfvo are exactly the even-split default (all percent, 0/33/67…).</summary>
    private static bool IsEvenPercentSplit(IReadOnlyList<IconSetThresholdEntry> entries)
    {
        var expected = EvenPercentThresholds(entries.Count).ToList();
        for (var i = 0; i < entries.Count; i++)
        {
            if (entries[i].Type != "percent" ||
                entries[i].Value != expected[i].ToString(CultureInfo.InvariantCulture))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>Wire (type, value) of a base data-bar cfvo (min/max → auto, num → fixed).</summary>
    private static (string Type, string? Value) ReadCfvo(S.ConditionalFormatValueObject cfvo)
    {
        var type = cfvo.Type?.Value;
        if (type == S.ConditionalFormatValueObjectValues.Min || type == S.ConditionalFormatValueObjectValues.Max)
        {
            return ("auto", null);
        }

        var wire = type == S.ConditionalFormatValueObjectValues.Number ? "fixed"
            : type == S.ConditionalFormatValueObjectValues.Percent ? "percent"
            : type == S.ConditionalFormatValueObjectValues.Percentile ? "percentile"
            : type == S.ConditionalFormatValueObjectValues.Formula ? "formula"
            : "auto";
        return (wire, cfvo.Val?.Value);
    }

    /// <summary>Wire name of a rule type (the supported kinds keep their op-prop spelling).</summary>
    public static string KindName(XLConditionalFormatType type) => type switch
    {
        XLConditionalFormatType.CellIs => "cellIs",
        XLConditionalFormatType.ColorScale => "colorScale",
        XLConditionalFormatType.DataBar => "dataBar",
        XLConditionalFormatType.ContainsText => "containsText",
        XLConditionalFormatType.IconSet => "iconSet",
        XLConditionalFormatType.Expression => "formula",
        XLConditionalFormatType.Top10 => "topBottom",
        XLConditionalFormatType.AboveAverage => "aboveBelowAverage",
        // OOXML spelling, NOT the lowercase fallback ('isDuplicate'/'isUnique').
        XLConditionalFormatType.IsDuplicate => "duplicateValues",
        XLConditionalFormatType.IsUnique => "uniqueValues",
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

    /// <summary>
    /// The 3-color colorScale midpoint endpoint (wire type + numeric value), read
    /// from <c>ContentTypes[2]</c>/<c>Values[2]</c>, but ONLY when it is NOT the
    /// legacy Percentile/50 default — so legacy colorScale rules project no
    /// midType/midValue and read byte-unchanged. Null on every other case.
    /// </summary>
    private static (string Type, double Value)? ColorScaleMidpoint(IXLConditionalFormat format)
    {
        if (format.ConditionalFormatType != XLConditionalFormatType.ColorScale || format.Colors.Count != 3)
        {
            return null;
        }

        if (!format.ContentTypes.TryGetValue(2, out var contentType) ||
            !format.Values.TryGetValue(2, out var value) || value?.Value is not { } valueText ||
            !double.TryParse(valueText, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
        {
            return null;
        }

        var wire = contentType switch
        {
            XLCFContentType.Number => "num",
            XLCFContentType.Percent => "percent",
            XLCFContentType.Percentile => "percentile",
            _ => null,
        };

        // The legacy default (Percentile/50) is omitted so legacy files read
        // unchanged; only a caller-pinned midpoint surfaces.
        if (wire is null || (wire == "percentile" && number == 50))
        {
            return null;
        }

        return (wire, number);
    }

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

    // ----- aboveBelowAverage raw authoring (v1.3) ---------------------------------

    /// <summary>
    /// One aboveBelowAverage rule captured at op time for the raw post-save pass.
    /// <c>Above</c> selects above-average cells (false = below); <c>EqualAverage</c>
    /// makes the comparison inclusive of the mean; <c>StdDev</c> &gt; 0 shifts the
    /// threshold that many standard deviations out. Style hex values are 8-digit
    /// ARGB; nulls leave that facet unset.
    /// </summary>
    internal sealed record AverageRuleSpec(
        string SheetName,
        string Sqref,
        bool Above,
        bool EqualAverage,
        int StdDev,
        string? Fill,
        string? FontColor,
        bool? Bold);

    /// <summary>
    /// One data-bar endpoint (a cfvo). <c>Type</c> is the wire spelling
    /// (auto/fixed/percent/percentile/formula); <c>RawValue</c> is the cfvo
    /// <c>@val</c> text (a number's invariant string, or the formula without its
    /// leading <c>=</c>); <c>Number</c> is the parsed numeric value for the
    /// number-typed endpoints (null for auto and formula). The static
    /// <see cref="Auto"/> instance is the byte-stable lowest/highest default.
    /// </summary>
    internal sealed record DataBarBound(string Type, string? RawValue, double? Number)
    {
        public static readonly DataBarBound Auto = new("auto", null, null);
    }

    /// <summary>
    /// A data-bar rule whose caller-supplied thresholds and/or showValue must be
    /// authored raw after the save (ClosedXML 0.105 always writes auto
    /// lowest/highest cfvo and showValue=1). Matched to the saved cfRule by sheet,
    /// sqref and bar color, in creation order. <c>Color</c> is 8-digit ARGB (the
    /// dataBar/color the ClosedXML rule carries); <c>ShowValue</c> is null when
    /// the caller left it at Excel's default (the value stays shown).
    /// </summary>
    internal sealed record DataBarThresholdSpec(
        string SheetName,
        string Sqref,
        string Color,
        DataBarBound Min,
        DataBarBound Max,
        bool? ShowValue);

    /// <summary>
    /// One iconSet threshold (a cfvo). <c>Type</c> is the wire spelling
    /// (num/percent/percentile/formula); <c>RawValue</c> is the cfvo <c>@val</c>
    /// text (a number's invariant string, or a formula without its leading
    /// <c>=</c>).
    /// </summary>
    internal sealed record IconSetBound(string Type, string RawValue);

    /// <summary>
    /// An iconSet rule whose caller-supplied per-icon thresholds must be authored
    /// raw after the save (ClosedXML 0.105 always writes the even-split percent
    /// cfvo). Matched to the saved cfRule by sheet, sqref and icon-set style, in
    /// creation order — a default (no-threshold) iconSet is never queued, so only
    /// the caller-controlled rules are touched. <c>Bounds</c> holds exactly one
    /// entry per icon, lowest band first. iconSet has NO x14 twin, so only the
    /// base cfvo are rewritten (simpler than dataBar).
    /// </summary>
    internal sealed record IconSetThresholdSpec(
        string SheetName,
        string Sqref,
        string SetName,
        IReadOnlyList<IconSetBound> Bounds);

    /// <summary>
    /// Authors the queued aboveBelowAverage rules on the saved file. ClosedXML
    /// 0.105 cannot serialize the <c>aboveAverage</c> rule type, so each rule's
    /// differential format (dxf) is appended to the styles part and a
    /// <c>conditionalFormatting/cfRule[@type=aboveAverage]</c> element is written
    /// on the worksheet with a priority above every existing rule, so the new
    /// rules sit last in priority order (matching the index the add op reported).
    /// </summary>
    public static void ApplyAverageRules(string file, IReadOnlyList<AverageRuleSpec> rules)
    {
        if (rules.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.WorkbookStylesPart?.Stylesheet is not { } stylesheet)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                "The workbook has no styles part to attach the conditional-format style to.",
                "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
        }

        var differentialFormats = stylesheet.GetFirstChild<S.DifferentialFormats>();
        if (differentialFormats is null)
        {
            differentialFormats = new S.DifferentialFormats();
            // dxfs sits after cellStyles and before tableStyles in the schema;
            // appending and letting the SDK reorder is unsafe, so insert before
            // the first known successor or at the end.
            var successor = stylesheet.Elements<OpenXmlElement>()
                .FirstOrDefault(e => e is S.TableStyles or S.Colors or S.StylesheetExtensionList);
            if (successor is null)
            {
                stylesheet.Append(differentialFormats);
            }
            else
            {
                stylesheet.InsertBefore(differentialFormats, successor);
            }
        }

        foreach (var rule in rules)
        {
            var worksheet = WorksheetFor(workbookPart, rule.SheetName);
            var dxfId = (uint)differentialFormats.Count();
            differentialFormats.Append(BuildDifferentialFormat(rule));

            var nextPriority = worksheet.Descendants<S.ConditionalFormattingRule>()
                .Select(r => r.Priority?.Value ?? 0)
                .DefaultIfEmpty(0)
                .Max() + 1;

            var cfRule = new S.ConditionalFormattingRule
            {
                Type = S.ConditionalFormatValues.AboveAverage,
                FormatId = dxfId,
                Priority = nextPriority,
                AboveAverage = rule.Above,
                EqualAverage = rule.EqualAverage ? true : (bool?)null,
            };
            if (rule.StdDev > 0)
            {
                cfRule.StdDev = rule.StdDev;
            }

            var formatting = new S.ConditionalFormatting(cfRule)
            {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = rule.Sqref },
            };
            InsertConditionalFormatting(worksheet, formatting);
        }

        differentialFormats.Count = (uint)differentialFormats.Count();
        stylesheet.Save();
    }

    /// <summary>The dxf for a rule's fill/font, matching ClosedXML's solid-fill-via-bgColor shape.</summary>
    private static S.DifferentialFormat BuildDifferentialFormat(AverageRuleSpec rule)
    {
        var dxf = new S.DifferentialFormat();
        if (rule.Bold is { } bold || rule.FontColor is { } fontColor)
        {
            var font = new S.Font();
            if (rule.Bold == true)
            {
                font.Append(new S.Bold());
            }
            else if (rule.Bold == false)
            {
                // An explicit non-bold override; Excel reads <b val="0"/>.
                font.Append(new S.Bold { Val = false });
            }

            if (rule.FontColor is { } hex)
            {
                font.Append(new S.Color { Rgb = hex });
            }

            dxf.Append(font);
        }

        if (rule.Fill is { } fillHex)
        {
            dxf.Append(new S.Fill(new S.PatternFill(new S.BackgroundColor { Rgb = fillHex })
            {
                PatternType = S.PatternValues.Solid,
            }));
        }

        return dxf;
    }

    /// <summary>
    /// Inserts a conditionalFormatting element in its schema slot. The worksheet
    /// schema puts conditionalFormatting after sheetData / mergeCells and before
    /// dataValidations / hyperlinks; place it before the first known successor.
    /// </summary>
    private static void InsertConditionalFormatting(S.Worksheet worksheet, S.ConditionalFormatting formatting)
    {
        var lastExisting = worksheet.Elements<S.ConditionalFormatting>().LastOrDefault();
        if (lastExisting is not null)
        {
            worksheet.InsertAfter(formatting, lastExisting);
            return;
        }

        var successor = worksheet.Elements<OpenXmlElement>().FirstOrDefault(e =>
            e is S.DataValidations or S.Hyperlinks or S.PrintOptions or S.PageMargins or S.PageSetup
                or S.HeaderFooter or S.RowBreaks or S.ColumnBreaks or S.Drawing or S.LegacyDrawing
                or S.TableParts or S.WorksheetExtensionList);
        if (successor is null)
        {
            worksheet.Append(formatting);
        }
        else
        {
            worksheet.InsertBefore(formatting, successor);
        }
    }

    private static S.Worksheet WorksheetFor(WorkbookPart workbookPart, string sheetName)
    {
        var sheet = workbookPart.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is { } relationshipId &&
            workbookPart.GetPartById(relationshipId) is WorksheetPart { Worksheet: { } worksheet })
        {
            return worksheet;
        }

        throw new AiofficeException(
            ErrorCodes.InternalError,
            $"Sheet '{sheetName}' disappeared between validation and the conditional-format write pass.",
            "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
    }

    // ----- data-bar threshold raw authoring (v1.15) -------------------------------

    /// <summary>
    /// Rewrites the cfvo endpoints and <c>@showValue</c> of the queued data-bar
    /// rules on the saved file. ClosedXML 0.105 always writes auto lowest/highest
    /// cfvo (<c>type=min/max</c> in the base rule, <c>type=num</c> with a 0
    /// formula in the x14 twin) and <c>showValue=1</c>; this pass replaces both
    /// the base and the x14 cfvo from the caller's spec so Excel honors the
    /// thresholds (Excel reads the x14 twin as authoritative when present).
    ///
    /// Specs are matched to the saved cfRules per sheet, by sqref and bar color,
    /// in creation order — a default (no-threshold) data bar is never queued, so
    /// only the caller-controlled rules are touched. Runs before
    /// <see cref="FixUpAfterSave"/>; the rewrite never changes the pairing ids, so
    /// the authored rule survives that pass's orphan-drop (it is not orphaned).
    /// </summary>
    public static void ApplyDataBarThresholds(string file, IReadOnlyList<DataBarThresholdSpec> rules)
    {
        if (rules.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return;
        }

        foreach (var sheetRules in rules.GroupBy(r => r.SheetName, StringComparer.OrdinalIgnoreCase))
        {
            var worksheet = WorksheetFor(workbookPart, sheetRules.Key);

            // The saved base dataBar cfRules and their x14 twins, both in document
            // order, keyed for matching by sqref + bar color.
            var baseBars = worksheet.Descendants<S.ConditionalFormatting>()
                .SelectMany(cf => cf.Elements<S.ConditionalFormattingRule>()
                    .Where(r => r.Type?.Value == S.ConditionalFormatValues.DataBar)
                    .Select(r => (Cf: cf, Rule: r)))
                .ToList();
            var x14Bars = worksheet.Descendants<X14.ConditionalFormattingRule>()
                .Where(r => r.Type?.Value == S.ConditionalFormatValues.DataBar)
                .ToList();

            var consumed = new HashSet<S.ConditionalFormattingRule>();
            foreach (var spec in sheetRules)
            {
                var match = baseBars.FirstOrDefault(b =>
                    !consumed.Contains(b.Rule) &&
                    string.Equals(b.Cf.SequenceOfReferences?.InnerText, spec.Sqref, StringComparison.Ordinal) &&
                    BarColorMatches(b.Rule, spec.Color));
                if (match.Rule is null)
                {
                    throw new AiofficeException(
                        ErrorCodes.InternalError,
                        $"The data-bar rule on '{spec.SheetName}!{spec.Sqref}' vanished before its thresholds could be written.",
                        "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
                }

                consumed.Add(match.Rule);
                RewriteBaseDataBar(match.Rule, spec);
                RewriteX14DataBar(match.Rule, x14Bars, spec);
            }

            worksheet.Save();
        }
    }

    /// <summary>True when a base dataBar cfRule's bar color matches the spec's ARGB hex.</summary>
    private static bool BarColorMatches(S.ConditionalFormattingRule rule, string argb)
    {
        var rgb = rule.GetFirstChild<S.DataBar>()?.GetFirstChild<S.Color>()?.Rgb?.Value;
        return string.Equals(rgb, argb, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Rewrites the base (downlevel) dataBar: its two cfvo (MIN before MAX) and
    /// the <c>@showValue</c>. An auto endpoint keeps ClosedXML's min/max cfvo
    /// untouched; a pinned endpoint becomes the matching <c>num/percent/
    /// percentile/formula</c> cfvo with the spec value.
    /// </summary>
    private static void RewriteBaseDataBar(S.ConditionalFormattingRule rule, DataBarThresholdSpec spec)
    {
        var dataBar = rule.GetFirstChild<S.DataBar>();
        if (dataBar is null)
        {
            return;
        }

        var cfvos = dataBar.Elements<S.ConditionalFormatValueObject>().ToList();
        if (cfvos.Count >= 2)
        {
            ApplyBaseCfvo(cfvos[0], spec.Min, isMin: true);
            ApplyBaseCfvo(cfvos[1], spec.Max, isMin: false);
        }

        // showValue: ClosedXML writes "1"; honor an explicit false, and leave the
        // default ("1") byte-identical when the caller did not set it.
        if (spec.ShowValue == false)
        {
            dataBar.ShowValue = false;
        }
    }

    /// <summary>Sets one base cfvo to the spec endpoint (auto keeps min/max with no value).</summary>
    private static void ApplyBaseCfvo(S.ConditionalFormatValueObject cfvo, DataBarBound bound, bool isMin)
    {
        cfvo.Type = bound.Type switch
        {
            "fixed" => S.ConditionalFormatValueObjectValues.Number,
            "percent" => S.ConditionalFormatValueObjectValues.Percent,
            "percentile" => S.ConditionalFormatValueObjectValues.Percentile,
            "formula" => S.ConditionalFormatValueObjectValues.Formula,
            _ => isMin ? S.ConditionalFormatValueObjectValues.Min : S.ConditionalFormatValueObjectValues.Max,
        };
        if (bound.Type == "auto")
        {
            // Excel's auto endpoints carry val="0" (ClosedXML's default); keep it.
            cfvo.Val = "0";
        }
        else
        {
            cfvo.Val = bound.RawValue;
        }
    }

    /// <summary>
    /// Rewrites the x14 twin's two cfvo (the authoritative endpoints Excel reads)
    /// and <c>@showValue</c>. The x14 cfvo carries its value in a child
    /// <c>xm:f</c>; auto endpoints use <c>min/max</c> with no formula.
    /// </summary>
    private static void RewriteX14DataBar(
        S.ConditionalFormattingRule baseRule, IReadOnlyList<X14.ConditionalFormattingRule> x14Bars, DataBarThresholdSpec spec)
    {
        var id = baseRule.Descendants<X14.Id>().FirstOrDefault()?.Text;
        if (id is null)
        {
            return;
        }

        // The base id and the twin id pair (case-insensitively; FixUpAfterSave
        // uppercases both later).
        var twin = x14Bars.FirstOrDefault(r =>
            string.Equals(r.Id?.Value, id, StringComparison.OrdinalIgnoreCase));
        var dataBar = twin?.GetFirstChild<X14.DataBar>();
        if (dataBar is null)
        {
            return;
        }

        var cfvos = dataBar.Elements<X14.ConditionalFormattingValueObject>().ToList();
        if (cfvos.Count >= 2)
        {
            ApplyX14Cfvo(cfvos[0], spec.Min, isMin: true);
            ApplyX14Cfvo(cfvos[1], spec.Max, isMin: false);
        }

        if (spec.ShowValue == false)
        {
            dataBar.ShowValue = false;
        }
    }

    /// <summary>Sets one x14 cfvo to the spec endpoint (auto keeps min/max with no formula).</summary>
    private static void ApplyX14Cfvo(X14.ConditionalFormattingValueObject cfvo, DataBarBound bound, bool isMin)
    {
        cfvo.Type = bound.Type switch
        {
            "fixed" => X14.ConditionalFormattingValueObjectTypeValues.Numeric,
            "percent" => X14.ConditionalFormattingValueObjectTypeValues.Percent,
            "percentile" => X14.ConditionalFormattingValueObjectTypeValues.Percentile,
            "formula" => X14.ConditionalFormattingValueObjectTypeValues.Formula,
            _ => isMin
                ? X14.ConditionalFormattingValueObjectTypeValues.Min
                : X14.ConditionalFormattingValueObjectTypeValues.Max,
        };

        // The x14 cfvo value lives in a child xm:f. Auto/min/max carry none.
        cfvo.RemoveAllChildren<Xm.Formula>();
        if (bound.Type != "auto" && bound.RawValue is { } raw)
        {
            cfvo.AppendChild(new Xm.Formula(raw));
        }
    }

    // ----- iconSet threshold raw authoring (v1.20) --------------------------------

    /// <summary>
    /// Rewrites the cfvo of the queued iconSet rules on the saved file. ClosedXML
    /// 0.105 always writes the even-split percent cfvo; this pass replaces each of
    /// the N base cfvo IN PLACE (no add/duplicate) from the caller's spec so Excel
    /// honors the per-icon bands. iconSet has NO x14 twin, so only the base cfvo
    /// are touched (simpler than the data-bar pass).
    ///
    /// Specs are matched to the saved cfRules per sheet, by sqref and icon-set
    /// style, in creation order — a default (no-threshold) iconSet is never
    /// queued, so only the caller-controlled rules are rewritten.
    /// </summary>
    public static void ApplyIconSetThresholds(string file, IReadOnlyList<IconSetThresholdSpec> rules)
    {
        if (rules.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return;
        }

        foreach (var sheetRules in rules.GroupBy(r => r.SheetName, StringComparer.OrdinalIgnoreCase))
        {
            var worksheet = WorksheetFor(workbookPart, sheetRules.Key);

            // The saved base iconSet cfRules in document order, keyed for matching
            // by sqref + icon-set style.
            var baseSets = worksheet.Descendants<S.ConditionalFormatting>()
                .SelectMany(cf => cf.Elements<S.ConditionalFormattingRule>()
                    .Where(r => r.Type?.Value == S.ConditionalFormatValues.IconSet)
                    .Select(r => (Cf: cf, Rule: r)))
                .ToList();

            var consumed = new HashSet<S.ConditionalFormattingRule>();
            foreach (var spec in sheetRules)
            {
                var match = baseSets.FirstOrDefault(b =>
                    !consumed.Contains(b.Rule) &&
                    string.Equals(b.Cf.SequenceOfReferences?.InnerText, spec.Sqref, StringComparison.Ordinal) &&
                    IconSetStyleMatches(b.Rule, spec.SetName));
                if (match.Rule is null)
                {
                    throw new AiofficeException(
                        ErrorCodes.InternalError,
                        $"The iconSet rule on '{spec.SheetName}!{spec.Sqref}' vanished before its thresholds could be written.",
                        "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
                }

                consumed.Add(match.Rule);
                RewriteBaseIconSet(match.Rule, spec);
            }

            worksheet.Save();
        }
    }

    /// <summary>True when a base iconSet cfRule's icon-set style matches the spec's wire set name.</summary>
    private static bool IconSetStyleMatches(S.ConditionalFormattingRule rule, string setName)
    {
        var value = rule.GetFirstChild<S.IconSet>()?.IconSetValue?.InnerText;
        return string.Equals(value, setName, StringComparison.Ordinal);
    }

    /// <summary>
    /// Rewrites the base iconSet's N cfvo in place from the spec (one entry per
    /// icon, lowest band first). The cfvo count MUST already equal the icon count
    /// (ClosedXML wrote the even-split set), so each existing cfvo is retyped and
    /// revalued without adding or removing any.
    /// </summary>
    private static void RewriteBaseIconSet(S.ConditionalFormattingRule rule, IconSetThresholdSpec spec)
    {
        var iconSet = rule.GetFirstChild<S.IconSet>();
        if (iconSet is null)
        {
            return;
        }

        var cfvos = iconSet.Elements<S.ConditionalFormatValueObject>().ToList();
        if (cfvos.Count != spec.Bounds.Count)
        {
            // Should never happen (ClosedXML wrote exactly N cfvo); guard anyway.
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"The iconSet rule on '{spec.SheetName}!{spec.Sqref}' has {cfvos.Count} cfvo, expected {spec.Bounds.Count}.",
                "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
        }

        for (var i = 0; i < cfvos.Count; i++)
        {
            ApplyIconSetCfvo(cfvos[i], spec.Bounds[i]);
        }
    }

    /// <summary>Sets one iconSet cfvo to the spec threshold (@type + @val) in place.</summary>
    private static void ApplyIconSetCfvo(S.ConditionalFormatValueObject cfvo, IconSetBound bound)
    {
        cfvo.Type = bound.Type switch
        {
            "num" => S.ConditionalFormatValueObjectValues.Number,
            "percent" => S.ConditionalFormatValueObjectValues.Percent,
            "percentile" => S.ConditionalFormatValueObjectValues.Percentile,
            "formula" => S.ConditionalFormatValueObjectValues.Formula,
            _ => S.ConditionalFormatValueObjectValues.Percent,
        };
        cfvo.Val = bound.RawValue;
    }

    // ----- aboveAverage round-trip preservation (v1.3) ----------------------------

    /// <summary>
    /// True when the file already carries raw <c>aboveAverage</c> conditional
    /// rules. ClosedXML 0.105 reads them but throws when re-serializing, so an
    /// edit batch must strip them from the ClosedXML model before saving and
    /// re-author them raw afterward (see <see cref="CaptureAverageRules"/> /
    /// <see cref="StripAverageRulesFromModel"/>). Cheap part scan.
    /// </summary>
    public static bool FileHasAverageRules(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return false;
        }

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            if (worksheetPart.Worksheet?.Descendants<S.ConditionalFormattingRule>()
                    .Any(r => r.Type?.Value == S.ConditionalFormatValues.AboveAverage) == true)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Captures every raw <c>aboveAverage</c> rule (sqref, flags, stdDev and its
    /// differential fill/font) so a later edit can re-author them after ClosedXML
    /// has rebuilt the file without them. Read straight from the package.
    /// </summary>
    public static List<AverageRuleSpec> CaptureAverageRules(string file)
    {
        var result = new List<AverageRuleSpec>();
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return result;
        }

        var dxfs = workbookPart.WorkbookStylesPart?.Stylesheet?.GetFirstChild<S.DifferentialFormats>()
            ?.Elements<S.DifferentialFormat>().ToList() ?? [];

        foreach (var sheet in workbookPart.Workbook.Descendants<S.Sheet>())
        {
            if (sheet.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart { Worksheet: { } worksheet })
            {
                continue;
            }

            var sheetName = sheet.Name?.Value ?? string.Empty;
            foreach (var formatting in worksheet.Descendants<S.ConditionalFormatting>())
            {
                var sqref = formatting.SequenceOfReferences?.InnerText ?? string.Empty;
                foreach (var rule in formatting.Elements<S.ConditionalFormattingRule>())
                {
                    if (rule.Type?.Value != S.ConditionalFormatValues.AboveAverage)
                    {
                        continue;
                    }

                    var dxf = rule.FormatId?.Value is { } formatId && formatId < dxfs.Count ? dxfs[(int)formatId] : null;
                    var (fill, fontColor, bold) = ReadDifferentialStyle(dxf);
                    result.Add(new AverageRuleSpec(
                        sheetName,
                        sqref,
                        Above: rule.AboveAverage?.Value ?? true,
                        EqualAverage: rule.EqualAverage?.Value ?? false,
                        StdDev: rule.StdDev?.Value ?? 0,
                        Fill: fill,
                        FontColor: fontColor,
                        Bold: bold));
                }
            }
        }

        return result;
    }

    /// <summary>Removes the AboveAverage rules ClosedXML cannot re-serialize from its in-memory model.</summary>
    public static void StripAverageRulesFromModel(XLWorkbook workbook)
    {
        foreach (var sheet in workbook.Worksheets)
        {
            sheet.ConditionalFormats.Remove(cf => cf.ConditionalFormatType == XLConditionalFormatType.AboveAverage);
        }
    }

    /// <summary>The fill/font hex (ARGB) of a differential format, for re-authoring a captured rule.</summary>
    private static (string? Fill, string? FontColor, bool? Bold) ReadDifferentialStyle(S.DifferentialFormat? dxf)
    {
        if (dxf is null)
        {
            return (null, null, null);
        }

        var fill = dxf.GetFirstChild<S.Fill>()?.PatternFill?.BackgroundColor?.Rgb?.Value;
        var font = dxf.GetFirstChild<S.Font>();
        var fontColor = font?.GetFirstChild<S.Color>()?.Rgb?.Value;
        bool? bold = font?.GetFirstChild<S.Bold>() is { } boldElement
            ? boldElement.Val?.Value ?? true
            : null;
        return (fill, fontColor, bold);
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
