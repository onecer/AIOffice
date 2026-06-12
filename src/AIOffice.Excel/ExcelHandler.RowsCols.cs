using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The M4 row/column structure ops, all ClosedXML-native (formula references
/// shift automatically on insert/delete — asserted by tests):
/// <list type="bullet">
/// <item>insert — <c>{op:add, type:row, path:/Sheet1/row[3], position:"before"|"after"}</c>
/// and <c>{op:add, type:col, path:/Sheet1/col[C], position:…}</c> (default before);</item>
/// <item>delete — <c>{op:remove, path:/Sheet1/row[3]}</c> / <c>/Sheet1/col[C]</c>;</item>
/// <item>sizing — <c>{op:set, path:/Sheet1/row[3], props:{height:24}}</c> (points, 0–409)
/// and <c>{op:set, path:/Sheet1/col[C], props:{width:18}}</c> (characters, 0–255);</item>
/// <item>hide — <c>{op:set, props:{hidden:true|false}}</c> on a row or col path.</item>
/// </list>
/// </summary>
public sealed partial class ExcelHandler
{
    private const double MaxRowHeightPoints = 409;
    private const double MaxColumnWidthChars = 255;

    private static readonly IReadOnlyList<string> InsertPositions = ["before", "after"];

    /// <summary>The validated insert position of an add op: "before" (default) or "after".</summary>
    private static string InsertPosition(EditOp op, int index)
    {
        var position = op.Position ?? "before";
        if (!InsertPositions.Contains(position, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: unknown position '{op.Position}'.",
                "Use position \"before\" (default) or \"after\".",
                candidates: InsertPositions);
        }

        return position;
    }

    private static void AddColumn(XLWorkbook workbook, EditOp op, int index, List<object> details)
    {
        var target = ExcelPaths.Resolve(workbook, op.Path);
        if (target.Kind != ExcelTargetKind.Column)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add col targets a column path like /Sheet1/col[C].",
                "Use {op:add, type:col, path:/Sheet1/col[C], position:\"before\"|\"after\"}; " +
                "column letters are uppercase.");
        }

        var position = InsertPosition(op, index);
        var anchor = target.ColumnNumber!.Value;
        int columnNumber;
        if (position == "after")
        {
            target.Sheet.Column(anchor).InsertColumnsAfter(1);
            columnNumber = anchor + 1;
        }
        else
        {
            target.Sheet.Column(anchor).InsertColumnsBefore(1);
            columnNumber = anchor;
        }

        details.Add(new
        {
            op = "add",
            type = "col",
            path = ExcelPaths.ColumnPath(target.Sheet, columnNumber),
            column = ExcelCharts.ColumnLetters(columnNumber),
        });
    }

    // ----- sizing & visibility -------------------------------------------------

    private static void ApplyRowHeight(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        if (target.Kind != ExcelTargetKind.Row)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: 'height' targets a row path like /Sheet1/row[3], not '{op.Path}'.",
                "Use {op:set, path:/Sheet1/row[3], props:{height:24}} (points). Column size is 'width'.");
        }

        var value = DoublePropValue(node, "height", index);
        if (value is < 0 or > MaxRowHeightPoints)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: height must be between 0 and {MaxRowHeightPoints} points; got {value}.",
                "Pass the row height in points, e.g. {\"height\":24}.");
        }

        target.Sheet.Row(target.RowNumber!.Value).Height = value;
        applied.Add("height");
    }

    private static void ApplyColumnWidth(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        if (target.Kind != ExcelTargetKind.Column)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: 'width' targets a column path like /Sheet1/col[C], not '{op.Path}'.",
                "Use {op:set, path:/Sheet1/col[C], props:{width:18}} (characters). Row size is 'height'.");
        }

        var value = DoublePropValue(node, "width", index);
        if (value is < 0 or > MaxColumnWidthChars)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: width must be between 0 and {MaxColumnWidthChars} characters; got {value}.",
                "Pass the column width in characters (Excel's unit), e.g. {\"width\":18}.");
        }

        target.Sheet.Column(target.ColumnNumber!.Value).Width = value;
        applied.Add("width");
    }

    private static void ApplyHidden(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        var hide = BoolPropValue(node, "hidden", index);
        switch (target.Kind)
        {
            case ExcelTargetKind.Row:
            {
                var row = target.Sheet.Row(target.RowNumber!.Value);
                if (hide)
                {
                    row.Hide();
                }
                else
                {
                    row.Unhide();
                }

                break;
            }

            case ExcelTargetKind.Column:
            {
                var column = target.Sheet.Column(target.ColumnNumber!.Value);
                if (hide)
                {
                    column.Hide();
                }
                else
                {
                    column.Unhide();
                }

                break;
            }

            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: 'hidden' targets a row or column path, not '{op.Path}'.",
                    "Use {op:set, path:/Sheet1/row[3], props:{hidden:true}} or /Sheet1/col[C]; false unhides.");
        }

        applied.Add(hide ? "hidden" : "unhidden");
    }

    private static double DoublePropValue(JsonNode node, string prop, int index)
    {
        if (node is JsonValue value)
        {
            if (value.GetValueKind() == JsonValueKind.Number)
            {
                // In-memory nodes can be int/long/decimal-backed; bridge them all.
                return value.TryGetValue<double>(out var number)
                    ? number
                    : Convert.ToDouble(value.GetValue<object>(), CultureInfo.InvariantCulture);
            }

            if (value.GetValueKind() == JsonValueKind.String &&
                double.TryParse(value.GetValue<string>(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: '{prop}' must be a number.",
            $"Pass e.g. {{\"{prop}\":18}}.");
    }
}
