using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// (1.7) Workbook calculation settings — the <c>calcPr</c> element:
/// <c>{op:set, path:"/", props:{calculationMode:"auto|manual|autoExceptTables",
/// iterativeCalc:true, maxIterations:100, maxChange:0.001, fullPrecision:true}}</c>.
///
/// These ride the workbook-root set alongside structure protection (a root op can
/// carry both). The <c>calcPr</c> is authored raw in a post-save pass so it is
/// authoritative: the normal save sets <c>fullCalculationOnLoad</c>/
/// <c>forceFullCalculation</c> (so Excel recomputes our cached values on open) and
/// this pass layers the user's mode/iteration settings onto that same element
/// without losing the recalc flags. ClosedXML preserves the element byte-identical
/// otherwise.
///
/// Wire <c>calculationMode</c> maps to OOXML <c>calcMode</c>:
/// <c>auto → auto</c>, <c>manual → manual</c>, <c>autoExceptTables → autoNoTable</c>.
/// </summary>
internal static class ExcelCalculation
{
    /// <summary>The workbook-root set props this layer owns.</summary>
    public static readonly IReadOnlyList<string> Props =
        ["calculationMode", "iterativeCalc", "maxIterations", "maxChange", "fullPrecision"];

    /// <summary>Wire calculationMode → OOXML calcMode attribute value.</summary>
    public static readonly IReadOnlyList<string> Modes = ["auto", "manual", "autoExceptTables"];

    /// <summary>One validated calc-settings change, queued for the post-save raw pass.</summary>
    internal sealed record Spec(
        S.CalculateModeValues? Mode,
        bool? Iterate,
        uint? MaxIterations,
        double? MaxChange,
        bool? FullPrecision);

    /// <summary>True when the workbook-root props include any calc-setting key.</summary>
    public static bool HasCalcProps(JsonObject props) =>
        props.Any(p => Props.Contains(p.Key, StringComparer.Ordinal));

    /// <summary>
    /// Validates the calc props on a workbook-root set and builds a queued spec.
    /// Returns the list of applied prop names for the response.
    /// </summary>
    public static (Spec Spec, List<string> Applied) Parse(JsonObject props, int index)
    {
        var applied = new List<string>();
        S.CalculateModeValues? mode = null;
        bool? iterate = null;
        uint? maxIterations = null;
        double? maxChange = null;
        bool? fullPrecision = null;

        if (props.TryGetPropertyValue("calculationMode", out var modeNode) && modeNode is not null)
        {
            var text = modeNode is JsonValue mv && mv.GetValueKind() == JsonValueKind.String
                ? mv.GetValue<string>()
                : throw Invalid(index, "calculationMode must be a string.", "Use \"auto\", \"manual\" or \"autoExceptTables\".");
            mode = text switch
            {
                "auto" => S.CalculateModeValues.Auto,
                "manual" => S.CalculateModeValues.Manual,
                "autoExceptTables" => S.CalculateModeValues.AutoNoTable,
                _ => throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: unknown calculationMode '{text}'.",
                    "Use \"auto\", \"manual\" or \"autoExceptTables\".",
                    candidates: Modes),
            };
            applied.Add("calculationMode");
        }

        if (props.TryGetPropertyValue("iterativeCalc", out var iterNode) && iterNode is not null)
        {
            iterate = Bool(iterNode, "iterativeCalc", index);
            applied.Add("iterativeCalc");
        }

        if (props.TryGetPropertyValue("maxIterations", out var maxItNode) && maxItNode is not null)
        {
            var value = Int(maxItNode, "maxIterations", index);
            if (value < 1 || value > 32767)
            {
                throw Invalid(index, $"maxIterations must be 1–32767; got {value}.", "Excel caps iterative calc at 32767 iterations.");
            }

            maxIterations = (uint)value;
            applied.Add("maxIterations");
        }

        if (props.TryGetPropertyValue("maxChange", out var maxChNode) && maxChNode is not null)
        {
            var value = Double(maxChNode, "maxChange", index);
            if (value < 0)
            {
                throw Invalid(index, $"maxChange must be 0 or more; got {value.ToString(CultureInfo.InvariantCulture)}.",
                    "maxChange is the maximum change between iterations (e.g. 0.001).");
            }

            maxChange = value;
            applied.Add("maxChange");
        }

        if (props.TryGetPropertyValue("fullPrecision", out var fpNode) && fpNode is not null)
        {
            fullPrecision = Bool(fpNode, "fullPrecision", index);
            applied.Add("fullPrecision");
        }

        return (new Spec(mode, iterate, maxIterations, maxChange, fullPrecision), applied);
    }

    /// <summary>Applies a queued calc spec to the file ClosedXML just saved (calcPr is authored raw).</summary>
    public static void Apply(string file, Spec spec)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is not { } workbook)
        {
            return;
        }

        var calcPr = workbook.CalculationProperties;
        if (calcPr is null)
        {
            calcPr = new S.CalculationProperties { CalculationId = 0 };
            InsertCalcPr(workbook, calcPr);
        }

        if (spec.Mode is { } mode)
        {
            calcPr.CalculationMode = mode;
        }

        if (spec.Iterate is { } iterate)
        {
            calcPr.Iterate = iterate;
        }

        if (spec.MaxIterations is { } maxIterations)
        {
            calcPr.IterateCount = maxIterations;
        }

        if (spec.MaxChange is { } maxChange)
        {
            calcPr.IterateDelta = maxChange;
        }

        if (spec.FullPrecision is { } fullPrecision)
        {
            calcPr.FullPrecision = fullPrecision;
        }

        workbook.Save();
    }

    /// <summary>
    /// Inserts a fresh calcPr in workbook document order. CT_Workbook orders
    /// calcPr after definedNames and before oleSize/customWorkbookViews/…; placing
    /// it after the last of (definedNames, sheets, bookViews) keeps it valid.
    /// </summary>
    private static void InsertCalcPr(S.Workbook workbook, S.CalculationProperties calcPr)
    {
        var after = (DocumentFormat.OpenXml.OpenXmlElement?)workbook.Elements<S.DefinedNames>().FirstOrDefault()
            ?? (DocumentFormat.OpenXml.OpenXmlElement?)workbook.Elements<S.Sheets>().FirstOrDefault()
            ?? workbook.Elements<S.BookViews>().FirstOrDefault();
        if (after is not null)
        {
            workbook.InsertAfter(calcPr, after);
        }
        else
        {
            workbook.Append(calcPr);
        }
    }

    /// <summary>The calc-settings block for <c>get /</c>; null when no calcPr is present.</summary>
    public static object? Read(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var calcPr = document.WorkbookPart?.Workbook?.CalculationProperties;
        if (calcPr is null)
        {
            return null;
        }

        var mode = calcPr.CalculationMode?.Value;
        return new
        {
            calculationMode = ModeName(mode),
            iterativeCalc = calcPr.Iterate?.Value is { } it ? it : (bool?)null,
            maxIterations = calcPr.IterateCount?.Value is { } ic ? (long)ic : (long?)null,
            maxChange = calcPr.IterateDelta?.Value,
            fullPrecision = calcPr.FullPrecision?.Value is { } fp ? fp : (bool?)null,
        };
    }

    private static string ModeName(S.CalculateModeValues? mode)
    {
        // OOXML default is "auto" when the attribute is absent.
        if (mode is null || mode == S.CalculateModeValues.Auto)
        {
            return "auto";
        }

        if (mode == S.CalculateModeValues.Manual)
        {
            return "manual";
        }

        return "autoExceptTables"; // AutoNoTable
    }

    private static AiofficeException Invalid(int index, string message, string suggestion) =>
        new(ErrorCodes.InvalidArgs, $"ops[{index}]: {message}", suggestion);

    private static bool Bool(JsonNode node, string prop, int index) =>
        node is JsonValue v && v.GetValueKind() is JsonValueKind.True or JsonValueKind.False
            ? v.GetValue<bool>()
            : node is JsonValue sv && sv.GetValueKind() == JsonValueKind.String && bool.TryParse(sv.GetValue<string>(), out var parsed)
                ? parsed
                : throw Invalid(index, $"{prop} must be true or false.", $"Pass e.g. {{\"{prop}\":true}}.");

    private static int Int(JsonNode node, string prop, int index)
    {
        if (node is JsonValue value)
        {
            if (value.GetValueKind() == JsonValueKind.Number && value.TryGetValue<int>(out var number))
            {
                return number;
            }

            if (value.GetValueKind() == JsonValueKind.String &&
                int.TryParse(value.GetValue<string>(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }
        }

        throw Invalid(index, $"{prop} must be a whole number.", $"Pass e.g. {{\"{prop}\":100}}.");
    }

    private static double Double(JsonNode node, string prop, int index)
    {
        if (node is JsonValue value)
        {
            if (value.GetValueKind() == JsonValueKind.Number && value.TryGetValue<double>(out var number))
            {
                return number;
            }

            if (value.GetValueKind() == JsonValueKind.String &&
                double.TryParse(value.GetValue<string>(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }
        }

        throw Invalid(index, $"{prop} must be a number.", $"Pass e.g. {{\"{prop}\":0.001}}.");
    }
}
