using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The M3 sheet-level set props, all ClosedXML-native:
/// <list type="bullet">
/// <item>freeze panes — <c>{op:set, path:/Sheet1, props:{freezeRows:1, freezeCols:2}}</c>, 0 clears an axis;</item>
/// <item>autofilter — <c>{op:set, path:/Sheet1/A1:D20, props:{autoFilter:true}}</c>, false clears,
/// one filter per sheet (a second range is invalid_args);</item>
/// <item>page setup — <c>{op:set, path:/Sheet1, props:{orientation:"landscape", paperSize:"A4",
/// fitToWidth:1, printArea:"A1:F40"}}</c>, get reflects everything.</item>
/// </list>
/// </summary>
public sealed partial class ExcelHandler
{
    private const int MaxSheetRows = 1048576;
    private const int MaxSheetColumns = 16384;

    private static readonly IReadOnlyList<string> Orientations = ["portrait", "landscape"];

    /// <summary>Wire name → ClosedXML paper size. Reflection back uses the same table.</summary>
    private static readonly IReadOnlyDictionary<string, XLPaperSize> PaperSizes =
        new Dictionary<string, XLPaperSize>(StringComparer.OrdinalIgnoreCase)
        {
            ["A3"] = XLPaperSize.A3Paper,
            ["A4"] = XLPaperSize.A4Paper,
            ["A5"] = XLPaperSize.A5Paper,
            ["Letter"] = XLPaperSize.LetterPaper,
            ["Legal"] = XLPaperSize.LegalPaper,
            ["Tabloid"] = XLPaperSize.TabloidPaper,
        };

    [GeneratedRegex("^[A-Z]{1,3}[0-9]{1,7}(:[A-Z]{1,3}[0-9]{1,7})?$")]
    private static partial Regex PrintAreaPattern();

    // ----- freeze panes -------------------------------------------------------

    private static void ApplyFreeze(ExcelTarget target, EditOp op, JsonNode node, string prop, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, prop, index);
        var value = IntPropValue(node, prop, index);
        var limit = prop == "freezeRows" ? MaxSheetRows - 1 : MaxSheetColumns - 1;
        if (value < 0 || value > limit)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: {prop} must be between 0 and {limit}; got {value}.",
                $"Pass the number of leading {(prop == "freezeRows" ? "rows" : "columns")} to freeze; 0 clears the freeze.");
        }

        if (prop == "freezeRows")
        {
            target.Sheet.SheetView.SplitRow = value;
        }
        else
        {
            target.Sheet.SheetView.SplitColumn = value;
        }

        applied.Add(prop);
    }

    // ----- autofilter ----------------------------------------------------------

    private static void ApplyAutoFilter(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        if (target.Kind != ExcelTargetKind.Range)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: autoFilter targets a range, not {target.Kind.ToString().ToLowerInvariant()} '{op.Path}'.",
                "Address the data including its header row, e.g. {op:set, path:/Sheet1/A1:D20, props:{autoFilter:true}}.");
        }

        var enable = BoolPropValue(node, "autoFilter", index);
        var sheet = target.Sheet;
        if (!enable)
        {
            sheet.AutoFilter.Clear();
            applied.Add("autoFilterCleared");
            return;
        }

        var requested = target.Range!.RangeAddress.ToString();
        if (sheet.AutoFilter.IsEnabled &&
            sheet.AutoFilter.Range?.RangeAddress.ToString() is { } existing &&
            !string.Equals(existing, requested, StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: sheet '{sheet.Name}' already has an autofilter on {existing}; Excel allows one per sheet.",
                "Clear it first with {op:set, path:" + ExcelPaths.SheetPath(sheet) + "/" + existing +
                ", props:{autoFilter:false}}, then set the new range.");
        }

        target.Range!.SetAutoFilter();
        applied.Add("autoFilter");
    }

    // ----- page setup -----------------------------------------------------------

    private static void ApplyOrientation(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "orientation", index);
        var text = StringPropValue(node, "orientation", index);
        if (!Orientations.Contains(text, StringComparer.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: unknown orientation '{text}'.",
                "Use \"portrait\" or \"landscape\".",
                candidates: Orientations);
        }

        target.Sheet.PageSetup.PageOrientation =
            string.Equals(text, "landscape", StringComparison.OrdinalIgnoreCase)
                ? XLPageOrientation.Landscape
                : XLPageOrientation.Portrait;
        applied.Add("orientation");
    }

    private static void ApplyPaperSize(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "paperSize", index);
        var text = StringPropValue(node, "paperSize", index);
        if (!PaperSizes.TryGetValue(text, out var size))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: unknown paperSize '{text}'.",
                "Supported paper sizes: " + string.Join(", ", PaperSizes.Keys.Order(StringComparer.Ordinal)) + ".",
                candidates: [.. PaperSizes.Keys.Order(StringComparer.Ordinal)]);
        }

        target.Sheet.PageSetup.PaperSize = size;
        applied.Add("paperSize");
    }

    private static void ApplyFitToWidth(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "fitToWidth", index);
        var value = IntPropValue(node, "fitToWidth", index);
        if (value < 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: fitToWidth must be 0 or more; got {value}.",
                "Pass the number of pages the sheet width must fit on (height stays automatic); 0 clears the fit.");
        }

        target.Sheet.PageSetup.FitToPages(value, 0); // height 0 = automatic
        applied.Add("fitToWidth");
    }

    private static void ApplyPrintArea(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "printArea", index);
        var text = StringPropValue(node, "printArea", index);
        var setup = target.Sheet.PageSetup;
        if (text.Length == 0)
        {
            setup.PrintAreas.Clear();
            applied.Add("printAreaCleared");
            return;
        }

        var normalized = text.ToUpperInvariant();
        if (!PrintAreaPattern().IsMatch(normalized))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{text}' is not a usable printArea.",
                "printArea is a plain range on the sheet itself, e.g. A1:F40 (no sheet prefix); \"\" clears it.");
        }

        setup.PrintAreas.Clear();
        setup.PrintAreas.Add(normalized);
        applied.Add("printArea");
    }

    // ----- get reflection --------------------------------------------------------

    /// <summary>The page-setup block for sheet get (null members are omitted on the wire).</summary>
    private static object PageSetupInfo(IXLWorksheet sheet)
    {
        var setup = sheet.PageSetup;
        var printAreas = setup.PrintAreas.Cast<IXLRange>()
            .Select(r => RelativeRangeText(r.RangeAddress))
            .ToList();
        return new
        {
            orientation = setup.PageOrientation.ToString().ToLowerInvariant(),
            paperSize = PaperSizes.FirstOrDefault(p => p.Value == setup.PaperSize).Key ?? setup.PaperSize.ToString(),
            fitToWidth = setup.PagesWide > 0 ? setup.PagesWide : (int?)null,
            printArea = printAreas.Count > 0 ? string.Join(",", printAreas) : null,
        };
    }

    /// <summary>Print areas stringify as absolute refs ($A$1:$F$40); the wire uses plain A1:F40.</summary>
    private static string RelativeRangeText(IXLRangeAddress address) => string.Create(
        CultureInfo.InvariantCulture,
        $"{address.FirstAddress.ColumnLetter}{address.FirstAddress.RowNumber}:" +
        $"{address.LastAddress.ColumnLetter}{address.LastAddress.RowNumber}");

    // ----- prop plumbing -----------------------------------------------------------

    private static void RequireSheetTarget(ExcelTarget target, EditOp op, string prop, int index)
    {
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{prop}' targets a sheet path like /Sheet1, not '{op.Path}'.",
                "Use {op:set, path:/Sheet1, props:{" + prop + ":…}}.");
        }
    }

    private static int IntPropValue(JsonNode node, string prop, int index)
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

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: '{prop}' must be a whole number.",
            $"Pass e.g. {{\"{prop}\":1}}.");
    }

    private static bool BoolPropValue(JsonNode node, string prop, int index)
    {
        if (node is JsonValue value)
        {
            switch (value.GetValueKind())
            {
                case JsonValueKind.True:
                    return true;
                case JsonValueKind.False:
                    return false;
                case JsonValueKind.String when bool.TryParse(value.GetValue<string>(), out var parsed):
                    return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: '{prop}' must be true or false.",
            $"Pass e.g. {{\"{prop}\":true}}.");
    }

    private static string StringPropValue(JsonNode node, string prop, int index)
    {
        if (node is JsonValue value && value.GetValueKind() == JsonValueKind.String)
        {
            return value.GetValue<string>();
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: '{prop}' must be a string.",
            $"Pass e.g. {{\"{prop}\":\"A4\"}}.");
    }
}
