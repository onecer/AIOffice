using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// Outline grouping (M6) — ClosedXML-native row/column groups
/// (<c>IXLRows.Group</c> / <c>Collapse</c> / <c>Ungroup</c>).
///
/// <para>Add a group: <c>{op:add, type:group, path:/Sheet1/row[2]:row[6]}</c>
/// groups rows 2–6; <c>/Sheet1/col[B]:col[E]</c> groups columns B–E.
/// <c>{collapsed:true}</c> collapses the new group. Remove ungroups (one outline
/// level) over the same span.</para>
///
/// <para>The <c>row[a]:row[b]</c> / <c>col[a]:col[b]</c> span form is parsed
/// here, not by the shared DocPath grammar (which has no element-range syntax),
/// exactly like the pivot/table id-forms.</para>
///
/// <para>Outline levels are reflected in <c>read --view structure</c> and a row/
/// column <c>get</c> (<c>outlineLevel</c>, <c>collapsed</c>); ClosedXML writes
/// the <c>outlineLevelRow</c>/<c>outlineLevelCol</c> sheet-format attributes and
/// the per-row/col <c>outlineLevel</c>/<c>collapsed</c> flags so Excel renders
/// the outline symbols on reopen.</para>
/// </summary>
internal static partial class ExcelGroups
{
    [GeneratedRegex(@"^(?<sheet>/.+)/(?i:row)\[(?<from>[0-9]+)\]:(?i:row)\[(?<to>[0-9]+)\]$")]
    private static partial Regex RowSpan();

    [GeneratedRegex(@"^(?<sheet>/.+)/(?i:col)\[(?<from>[A-Z]{1,3})\]:(?i:col)\[(?<to>[A-Z]{1,3})\]$")]
    private static partial Regex ColumnSpan();

    /// <summary>What a group path resolved to: a row span or a column span on one sheet.</summary>
    internal sealed record GroupSpan(IXLWorksheet Sheet, bool IsRow, int First, int Last);

    /// <summary>
    /// True when <paramref name="pathText"/> is a group span form, with the span
    /// resolved against <paramref name="workbook"/>. Returns false (not a throw)
    /// for non-span paths so the caller can fall through to the normal resolver.
    /// </summary>
    public static bool TryResolveSpan(XLWorkbook workbook, string pathText, out GroupSpan? span)
    {
        span = null;

        var rowMatch = RowSpan().Match(pathText);
        if (rowMatch.Success)
        {
            var sheet = SheetOf(workbook, rowMatch.Groups["sheet"].Value, pathText);
            var first = int.Parse(rowMatch.Groups["from"].Value, CultureInfo.InvariantCulture);
            var last = int.Parse(rowMatch.Groups["to"].Value, CultureInfo.InvariantCulture);
            span = new GroupSpan(sheet, IsRow: true, OrderLow(first, last), OrderHigh(first, last));
            return true;
        }

        var columnMatch = ColumnSpan().Match(pathText);
        if (columnMatch.Success)
        {
            var sheet = SheetOf(workbook, columnMatch.Groups["sheet"].Value, pathText);
            var first = new CellRef(columnMatch.Groups["from"].Value, 1).ColumnNumber;
            var last = new CellRef(columnMatch.Groups["to"].Value, 1).ColumnNumber;
            span = new GroupSpan(sheet, IsRow: false, OrderLow(first, last), OrderHigh(first, last));
            return true;
        }

        return false;
    }

    private static int OrderLow(int a, int b) => Math.Min(a, b);

    private static int OrderHigh(int a, int b) => Math.Max(a, b);

    private static IXLWorksheet SheetOf(XLWorkbook workbook, string sheetPath, string pathText)
    {
        var target = ExcelPaths.Resolve(workbook, sheetPath);
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"A group span must follow a sheet name: {pathText}",
                "Use /SheetName/row[2]:row[6] or /SheetName/col[B]:col[E].");
        }

        return target.Sheet;
    }

    // ----- add (group) --------------------------------------------------------

    public static object Add(XLWorkbook workbook, EditOp op, int index)
    {
        if (!TryResolveSpan(workbook, op.Path, out var span) || span is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add group needs a row or column span path like /Sheet1/row[2]:row[6] or /Sheet1/col[B]:col[E].",
                "Group an outline span; {collapsed:true} collapses it. Remove ungroups the span.");
        }

        var collapsed = op.Props?.TryGetPropertyValue("collapsed", out var node) == true &&
                        node is JsonValue value &&
                        value.GetValueKind() == System.Text.Json.JsonValueKind.True;

        if (span.IsRow)
        {
            var rows = span.Sheet.Rows(span.First, span.Last);
            rows.Group();
            if (collapsed)
            {
                rows.Collapse();
            }
        }
        else
        {
            var columns = span.Sheet.Columns(span.First, span.Last);
            columns.Group();
            if (collapsed)
            {
                columns.Collapse();
            }
        }

        return new
        {
            op = "add",
            type = "group",
            path = SpanPath(span),
            axis = span.IsRow ? "row" : "col",
            from = Endpoint(span, span.First),
            to = Endpoint(span, span.Last),
            collapsed,
            outlineLevel = OutlineLevelOf(span),
        };
    }

    // ----- remove (ungroup) ---------------------------------------------------

    public static object Remove(XLWorkbook workbook, EditOp op, int index)
    {
        if (!TryResolveSpan(workbook, op.Path, out var span) || span is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: remove group needs a row or column span path like /Sheet1/row[2]:row[6].",
                "Ungroup the same span that was grouped (one outline level).");
        }

        if (span.IsRow)
        {
            span.Sheet.Rows(span.First, span.Last).Ungroup();
        }
        else
        {
            span.Sheet.Columns(span.First, span.Last).Ungroup();
        }

        return new
        {
            op = "remove",
            path = SpanPath(span),
            removed = "group",
            axis = span.IsRow ? "row" : "col",
            from = Endpoint(span, span.First),
            to = Endpoint(span, span.Last),
        };
    }

    // ----- structure reflection ----------------------------------------------

    private static int OutlineLevelOf(GroupSpan span) => span.IsRow
        ? span.Sheet.Row(span.First).OutlineLevel
        : span.Sheet.Column(span.First).OutlineLevel;

    private static string SpanPath(GroupSpan span) => span.IsRow
        ? string.Create(
            CultureInfo.InvariantCulture,
            $"{ExcelPaths.SheetPath(span.Sheet)}/row[{span.First}]:row[{span.Last}]")
        : string.Create(
            CultureInfo.InvariantCulture,
            $"{ExcelPaths.SheetPath(span.Sheet)}/col[{ExcelCharts.ColumnLetters(span.First)}]:col[{ExcelCharts.ColumnLetters(span.Last)}]");

    private static string Endpoint(GroupSpan span, int n) => span.IsRow
        ? n.ToString(CultureInfo.InvariantCulture)
        : ExcelCharts.ColumnLetters(n);
}
