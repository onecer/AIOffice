using System.Text;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

public sealed partial class ExcelHandler
{
    private const int DefaultMaxTextBytes = 65536;

    private static readonly IReadOnlyList<string> ReadViews = ["outline", "text", "stats", "structure"];

    public Envelope Read(CommandContext ctx) => Run(ctx, sw =>
    {
        var file = RequireFile(ctx, mustExist: true);
        var view = ArgString(ctx, "view") ?? "stats";
        using var workbook = OpenWorkbook(file);

        return view switch
        {
            "stats" => Envelope.Ok(ReadStats(workbook), MetaFor(file, sw)),
            "outline" => Envelope.Ok(ReadOutline(workbook), MetaFor(file, sw)),
            "structure" => Envelope.Ok(ReadStructure(workbook), MetaFor(file, sw)),
            "text" => ReadText(ctx, workbook, file, sw),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown view '{view}'.",
                "Use one of: outline, text, stats, structure.",
                candidates: ReadViews),
        };
    });

    private static object ReadStats(XLWorkbook workbook)
    {
        var sheets = workbook.Worksheets
            .OrderBy(ws => ws.Position)
            .Select(ws =>
            {
                var used = ws.RangeUsed();
                var cells = ws.CellsUsed().ToList();
                return new
                {
                    name = ws.Name,
                    position = ws.Position,
                    usedRange = used?.RangeAddress.ToString(),
                    cellCount = cells.Count,
                    formulaCount = cells.Count(c => c.HasFormula),
                };
            })
            .ToList();

        return new
        {
            kind = "xlsx",
            sheets,
            totals = new
            {
                sheets = sheets.Count,
                cells = sheets.Sum(s => s.cellCount),
                formulas = sheets.Sum(s => s.formulaCount),
            },
        };
    }

    private static object ReadOutline(XLWorkbook workbook) => new
    {
        kind = "xlsx",
        sheets = workbook.Worksheets
            .OrderBy(ws => ws.Position)
            .Select(ws => new
            {
                position = ws.Position,
                name = ws.Name,
                path = ExcelPaths.SheetPath(ws),
                visible = ws.Visibility == XLWorksheetVisibility.Visible,
                usedRange = ws.RangeUsed()?.RangeAddress.ToString(),
            })
            .ToList(),
    };

    private static object ReadStructure(XLWorkbook workbook) => new
    {
        kind = "xlsx",
        sheets = workbook.Worksheets
            .OrderBy(ws => ws.Position)
            .Select(ws => new
            {
                name = ws.Name,
                path = ExcelPaths.SheetPath(ws),
                position = ws.Position,
                visible = ws.Visibility == XLWorksheetVisibility.Visible,
                usedRange = ws.RangeUsed()?.RangeAddress.ToString(),
                tables = ws.Tables
                    .Select(t => new
                    {
                        name = t.Name,
                        range = t.RangeAddress.ToString(),
                        columns = t.Fields.Select(f => f.Name).ToList(),
                    })
                    .ToList(),
                mergedRanges = ws.MergedRanges.Select(r => r.RangeAddress.ToString()).ToList(),
            })
            .ToList(),
        definedNames = workbook.DefinedNames
            .Select(n => new { name = n.Name, refersTo = n.RefersTo })
            .ToList(),
    };

    private Envelope ReadText(CommandContext ctx, XLWorkbook workbook, string file, System.Diagnostics.Stopwatch sw)
    {
        var maxBytes = ArgInt(ctx, "maxBytes") ?? DefaultMaxTextBytes;
        var (content, truncated) = BuildText(workbook, ArgString(ctx, "range"), maxBytes);

        List<Warning>? warnings = truncated
            ? [new Warning(
                "result_truncated",
                $"Text output exceeded {maxBytes} bytes and was cut off. Narrow it with --range (e.g. --range A1:F50) or raise --max-bytes.")]
            : null;

        return Envelope.Ok(
            new { view = "text", content, truncated },
            MetaFor(file, sw, warnings));
    }

    /// <summary>
    /// CSV-ish text of the workbook: one section per sheet, each headed by
    /// <c># Sheet!Range</c>. A range argument (either a full path like
    /// <c>/Sheet1/A1:C10</c> or a bare <c>A1:C10</c> against the first sheet)
    /// narrows the window.
    /// </summary>
    private static (string Content, bool Truncated) BuildText(XLWorkbook workbook, string? rangeArg, int maxBytes)
    {
        var sections = ResolveTextSections(workbook, rangeArg);
        var sb = new StringBuilder();
        foreach (var (sheet, range) in sections)
        {
            if (sb.Length > 0)
            {
                sb.Append('\n');
            }

            if (range is null)
            {
                sb.Append("# ").Append(sheet.Name).Append(" (empty)\n");
                continue;
            }

            sb.Append("# ").Append(sheet.Name).Append('!').Append(range.RangeAddress.ToString()).Append('\n');
            foreach (var row in range.Rows())
            {
                var first = true;
                foreach (var cell in row.Cells(true))
                {
                    if (!first)
                    {
                        sb.Append(',');
                    }

                    sb.Append(ExcelValues.CsvEscape(ExcelValues.SafeFormatted(cell)));
                    first = false;
                }

                sb.Append('\n');
                if (sb.Length > maxBytes)
                {
                    return (sb.ToString(0, Math.Min(sb.Length, maxBytes)), true);
                }
            }
        }

        return (sb.ToString(), false);
    }

    private static List<(IXLWorksheet Sheet, IXLRange? Range)> ResolveTextSections(XLWorkbook workbook, string? rangeArg)
    {
        if (rangeArg is null)
        {
            return [.. workbook.Worksheets
                .OrderBy(ws => ws.Position)
                .Select(ws => (ws, (IXLRange?)ws.RangeUsed()))];
        }

        if (rangeArg.StartsWith('/'))
        {
            var target = ExcelPaths.Resolve(workbook, rangeArg);
            return [(target.Sheet, TargetAsRange(target))];
        }

        var firstSheet = workbook.Worksheets.OrderBy(ws => ws.Position).First();
        var onFirstSheet = ExcelPaths.Resolve(workbook, ExcelPaths.SheetPath(firstSheet) + "/" + rangeArg.ToUpperInvariant());
        return [(onFirstSheet.Sheet, TargetAsRange(onFirstSheet))];
    }

    /// <summary>Widens any resolved target to a concrete range (used range for sheets, used cells for rows).</summary>
    private static IXLRange? TargetAsRange(ExcelTarget target) => target.Kind switch
    {
        ExcelTargetKind.Cell => target.Cell!.AsRange(),
        ExcelTargetKind.Range => target.Range,
        ExcelTargetKind.Row => RowUsedRange(target.Sheet, target.RowNumber!.Value),
        _ => target.Sheet.RangeUsed(),
    };

    private static IXLRange? RowUsedRange(IXLWorksheet sheet, int rowNumber)
    {
        var lastColumn = sheet.Row(rowNumber).LastCellUsed()?.Address.ColumnNumber;
        return lastColumn is null ? null : sheet.Range(rowNumber, 1, rowNumber, lastColumn.Value);
    }
}
