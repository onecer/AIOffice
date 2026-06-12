using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

public sealed partial class ExcelHandler
{
    private const int DefaultMaxRangeCells = 1000;

    public Envelope Get(CommandContext ctx) => Run(ctx, sw =>
    {
        var file = RequireFile(ctx, mustExist: true);
        var pathArg = ArgString(ctx, "path") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "get needs a path.",
            "Pass an address like /Sheet1/A1 or /Sheet1/A1:C10; run 'aioffice query' to discover paths.");

        using var workbook = OpenWorkbook(file);
        var target = ExcelPaths.Resolve(workbook, pathArg);

        return target.Kind switch
        {
            ExcelTargetKind.Sheet => Envelope.Ok(SheetInfo(target.Sheet), MetaFor(file, sw)),
            ExcelTargetKind.Cell => Envelope.Ok(CellInfo(target.Sheet, target.Cell!), MetaFor(file, sw)),
            ExcelTargetKind.Row => Envelope.Ok(RowInfo(target.Sheet, target.RowNumber!.Value), MetaFor(file, sw)),
            ExcelTargetKind.Chart => Envelope.Ok(ChartTargetInfo(file, target), MetaFor(file, sw)),
            ExcelTargetKind.Pivot => Envelope.Ok(PivotTargetInfo(file, target), MetaFor(file, sw)),
            ExcelTargetKind.ConditionalFormat => Envelope.Ok(
                ExcelConditionalFormats.Describe(
                    target.Sheet, ExcelConditionalFormats.Find(target), target.ConditionalFormatIndex!.Value),
                MetaFor(file, sw)),
            ExcelTargetKind.Image => Envelope.Ok(
                ExcelImages.Describe(target.Sheet, ExcelImages.Find(target), target.ImageIndex!.Value),
                MetaFor(file, sw)),
            _ => RangeInfo(ctx, target, file, sw),
        };
    });

    /// <summary>One pivot table; the source range comes from the raw cache part.</summary>
    private static object PivotTargetInfo(string file, ExcelTarget target)
    {
        var (pivot, _) = ExcelPivots.Find(target);
        Dictionary<(string, string), (string?, string?)> sources;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(file, isEditable: false))
        {
            sources = ExcelPivots.ReadSources(document);
        }

        return ExcelPivots.Describe(target.Sheet, pivot, sources);
    }

    /// <summary>One chart, read back from the raw package (ClosedXML cannot see charts).</summary>
    private static object ChartTargetInfo(string file, ExcelTarget target)
    {
        List<ChartInfo> charts;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(file, isEditable: false))
        {
            charts = ExcelCharts.Read(document)
                .Where(c => string.Equals(c.SheetName, target.Sheet.Name, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        var wanted = target.ChartIndex!.Value;
        var info = charts.FirstOrDefault(c => c.Index == wanted);
        if (info is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No chart[{wanted}] on sheet '{target.Sheet.Name}' ({charts.Count} chart(s) exist).",
                charts.Count > 0
                    ? "Chart indices are 1-based per sheet; pick one of the candidates."
                    : "This sheet has no charts; add one with edit op {op:add, type:chart, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + ", props:{kind:\"bar\", dataRange:\"A1:B5\", anchor:\"D2\"}}.",
                candidates: charts.Count > 0
                    ? [.. charts.Select(c => c.Path)]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return new
        {
            path = info.Path,
            kind = "chart",
            sheet = target.Sheet.Name,
            chartKind = info.Kind,
            title = info.Title,
            dataRange = info.DataRange,
            anchor = info.Anchor,
            series = info.Series,
        };
    }

    private static object SheetInfo(IXLWorksheet sheet)
    {
        var used = sheet.RangeUsed();
        return new
        {
            path = ExcelPaths.SheetPath(sheet),
            kind = "sheet",
            name = sheet.Name,
            position = sheet.Position,
            visible = sheet.Visibility == XLWorksheetVisibility.Visible,
            usedRange = used?.RangeAddress.ToString(),
            lastRow = sheet.LastRowUsed()?.RowNumber(),
            lastColumn = sheet.LastColumnUsed()?.ColumnNumber(),
            tables = sheet.Tables.Select(t => t.Name).ToList(),
            mergedRanges = sheet.MergedRanges.Select(r => r.RangeAddress.ToString()).ToList(),
        };
    }

    /// <summary>One cell with its typed value and the properties an agent needs to edit it.</summary>
    private static object CellInfo(IXLWorksheet sheet, IXLCell cell)
    {
        XLCellValue value;
        try
        {
            value = cell.Value; // evaluates a dirty formula in memory; nothing is written
        }
        catch (Exception)
        {
            value = cell.CachedValue;
        }

        var numberFormat = cell.Style.NumberFormat.Format;
        return new
        {
            path = ExcelPaths.CellPath(sheet, cell.Address),
            kind = "cell",
            sheet = sheet.Name,
            address = cell.Address.ToString(),
            value = ExcelValues.ToJson(value),
            type = ExcelValues.TypeName(value.Type),
            formula = cell.HasFormula ? "=" + cell.FormulaA1 : null,
            cachedValue = cell.HasFormula ? ExcelValues.ToJson(cell.CachedValue) : null,
            text = ExcelValues.SafeFormatted(cell),
            numberFormat = string.IsNullOrEmpty(numberFormat) ? null : numberFormat,
            bold = cell.Style.Font.Bold ? true : (bool?)null,
            italic = cell.Style.Font.Italic ? true : (bool?)null,
            merged = MergedRangeOf(sheet, cell),
        };
    }

    private static string? MergedRangeOf(IXLWorksheet sheet, IXLCell cell)
    {
        foreach (var range in sheet.MergedRanges)
        {
            var a = range.RangeAddress;
            if (cell.Address.RowNumber >= a.FirstAddress.RowNumber &&
                cell.Address.RowNumber <= a.LastAddress.RowNumber &&
                cell.Address.ColumnNumber >= a.FirstAddress.ColumnNumber &&
                cell.Address.ColumnNumber <= a.LastAddress.ColumnNumber)
            {
                return a.ToString();
            }
        }

        return null;
    }

    private static object RowInfo(IXLWorksheet sheet, int rowNumber)
    {
        var lastColumn = sheet.Row(rowNumber).LastCellUsed()?.Address.ColumnNumber ?? 0;
        var values = new List<object?>(lastColumn);
        for (var column = 1; column <= lastColumn; column++)
        {
            values.Add(ExcelValues.ToJson(sheet.Cell(rowNumber, column).Value));
        }

        return new
        {
            path = $"/{ExcelPaths.QuoteSheet(sheet.Name)}/row[{rowNumber}]",
            kind = "row",
            sheet = sheet.Name,
            row = rowNumber,
            values,
        };
    }

    /// <summary>A range as a compact 2D value array, truncated row-wise beyond the cell cap.</summary>
    private static Envelope RangeInfo(CommandContext ctx, ExcelTarget target, string file, System.Diagnostics.Stopwatch sw)
    {
        var range = target.Range!;
        var maxCells = ArgInt(ctx, "maxCells") ?? DefaultMaxRangeCells;
        var address = range.RangeAddress;
        var columns = address.LastAddress.ColumnNumber - address.FirstAddress.ColumnNumber + 1;
        var totalRows = address.LastAddress.RowNumber - address.FirstAddress.RowNumber + 1;
        var maxRows = Math.Max(1, maxCells / Math.Max(1, columns));

        var values = new List<List<object?>>();
        var emitted = 0;
        foreach (var row in range.Rows())
        {
            if (emitted >= maxRows)
            {
                break;
            }

            var rowValues = new List<object?>(columns);
            for (var column = 1; column <= columns; column++)
            {
                rowValues.Add(ExcelValues.ToJson(row.Cell(column).Value));
            }

            values.Add(rowValues);
            emitted++;
        }

        var truncated = totalRows > emitted;
        List<Warning>? warnings = truncated
            ? [new Warning(
                "result_truncated",
                $"Range has {totalRows} rows; returning the first {emitted}. Request a smaller range or raise maxCells.")]
            : null;

        return Envelope.Ok(
            new
            {
                path = ExcelPaths.RangePath(target.Sheet, address),
                kind = "range",
                sheet = target.Sheet.Name,
                range = address.ToString(),
                rows = totalRows,
                columns,
                values,
                truncated,
            },
            MetaFor(file, sw, warnings));
    }
}
