using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

public sealed partial class ExcelHandler
{
    /// <summary>
    /// The M5 import hook the integrator wires to <c>create --from data.csv</c>:
    /// builds a NEW workbook whose first sheet holds the csv content, typed by
    /// the documented heuristic (<see cref="ExcelValues.ParseCsvField"/>).
    /// Values land at A1 through the existing bulk-write path — imports over
    /// 50k cells into the (necessarily bare) sheet stream through the SAX
    /// writer instead of the ClosedXML DOM, exactly like a big
    /// <c>set values</c> op.
    /// </summary>
    /// <param name="ctx">Target workbook context (<c>title</c> names the sheet, <c>delimiter</c> overrides sniffing).</param>
    /// <param name="sourcePath">The csv file, sandbox-resolved against the workspace.</param>
    public Envelope CreateFrom(CommandContext ctx, string sourcePath) => Run(ctx, sw =>
    {
        var file = RequireFile(ctx, mustExist: false);
        if (File.Exists(file))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"File already exists: {file}",
                "Pick a new file name, or use 'aioffice edit' to change the existing workbook.");
        }

        var source = ctx.Workspace.Resolve(sourcePath, mustExist: true);
        var extension = Path.GetExtension(source);
        if (!extension.Equals(".csv", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".tsv", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".txt", StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"xlsx import understands csv (and tsv/txt) sources; '{extension}' is not one.",
                "Export the data as .csv first, or write it with edit ops " +
                "({op:set, path:/Sheet1/A1, props:{values:[[…]]}}).",
                candidates: [".csv", ".tsv", ".txt"]);
        }

        FileSizeGuard.Ensure(source); // same guard as workbooks: no surprise gigabyte reads
        var text = File.ReadAllText(source); // StreamReader BOM detection strips a UTF-8 BOM
        var delimiter = ExcelCsv.ParseDelimiterArg(ArgString(ctx, "delimiter")) ?? ExcelCsv.Sniff(text);
        var grid = ExcelCsv.Parse(text, delimiter);

        var columns = grid.Count == 0 ? 0 : grid.Max(r => r.Count);
        if (grid.Count > MaxSheetRows || columns > MaxSheetColumns)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"The csv is {grid.Count} row(s) x {columns} column(s); a sheet ends at " +
                $"row {MaxSheetRows}, column {MaxSheetColumns}.",
                "Split the csv into smaller files and import each into its own workbook or sheet.");
        }

        var title = ArgString(ctx, "title") ?? "Sheet1";
        var warnings = new List<Warning>();
        if (grid.Count == 0)
        {
            warnings.Add(new Warning(
                "csv_empty",
                $"The csv has no data rows; created an empty workbook. Source: {source}"));
        }

        var directory = Path.GetDirectoryName(file);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var streamed = false;
        using (var workbook = new XLWorkbook())
        {
            var sheet = AddSheetOrThrowForImport(workbook, title);
            if (grid.Count > 0 &&
                ExcelBulkWrites.CellCount(grid) > ExcelBulkWrites.StreamingThresholdCells &&
                ExcelBulkWrites.IsStreamable(grid))
            {
                workbook.SaveAs(file); // bare save first; the SAX pass rewrites sheetData
                ExcelBulkWrites.ApplyAfterSave(file, [new ExcelBulkWrites.Pending(sheet, 1, 1, grid)]);
                streamed = true;
            }
            else
            {
                if (grid.Count > 0)
                {
                    WriteGridDom(sheet, 1, 1, grid);
                }

                if (SaveWithCachedValues(workbook, file) is { } saveWarnings)
                {
                    warnings.AddRange(saveWarnings);
                }
            }
        }

        // Streamed formulas need cached values: one normal pass over the bytes.
        if (streamed && grid.Any(row => row.Any(v => v.IsFormula)))
        {
            using var reopened = OpenWorkbook(file);
            if (SaveWithCachedValues(reopened, file) is { } formulaWarnings)
            {
                warnings.AddRange(formulaWarnings);
            }
        }

        // The bare-save SAX branch (no formulas) skips SaveWithCachedValues, so
        // its core properties are still in the legacy .psmdcp part; normalize so
        // every csv-imported workbook lands them at docProps/core.xml too.
        if (streamed)
        {
            ExcelCoreProperties.NormalizeAfterSave(file);
        }

        return Envelope.Ok(
            new
            {
                file,
                kind = "xlsx",
                sheet = title,
                source,
                delimiter = delimiter.ToString(),
                rows = grid.Count,
                columns,
                cells = ExcelBulkWrites.CellCount(grid),
                streamed,
            },
            MetaFor(file, sw, warnings.Count > 0 ? warnings : null));
    });

    private static IXLWorksheet AddSheetOrThrowForImport(XLWorkbook workbook, string name)
    {
        AddSheetOrThrow(workbook, name);
        return workbook.Worksheet(name);
    }

    // ----- read --view csv ------------------------------------------------------

    /// <summary>
    /// The M5 export view: ONE sheet window as RFC 4180 csv. <c>--sheet</c>
    /// picks the sheet (default: first), <c>--range</c> narrows the window
    /// (bare <c>A1:D100</c>, or a full path that names the sheet itself).
    /// </summary>
    private Envelope ReadCsv(CommandContext ctx, XLWorkbook workbook, string file, System.Diagnostics.Stopwatch sw)
    {
        var (sheet, range) = ResolveCsvWindow(workbook, ArgString(ctx, "sheet"), ArgString(ctx, "range"));
        var maxBytes = ArgInt(ctx, "maxBytes") ?? DefaultMaxTextBytes;
        var delimiter = ExcelCsv.ParseDelimiterArg(ArgString(ctx, "delimiter")) ?? ',';
        var (content, truncated) = ExcelCsv.Build(range, maxBytes, delimiter);

        List<Warning>? warnings = truncated
            ? [new Warning(
                "result_truncated",
                $"Csv output exceeded {maxBytes} bytes and was cut off. Narrow it with --range (e.g. --range A1:D100) or raise --max-bytes.")]
            : null;

        return Envelope.Ok(
            new
            {
                view = "csv",
                sheet = sheet.Name,
                range = range?.RangeAddress.ToString(),
                content,
                truncated,
            },
            MetaFor(file, sw, warnings));
    }

    private static (IXLWorksheet Sheet, IXLRange? Range) ResolveCsvWindow(
        XLWorkbook workbook, string? sheetArg, string? rangeArg)
    {
        if (rangeArg is not null && rangeArg.StartsWith('/'))
        {
            if (sheetArg is not null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "Pass the sheet either inside the range path or as --sheet, not both.",
                    "Use --range /Sheet1/A1:D100 alone, or --sheet Sheet1 --range A1:D100.");
            }

            var target = ExcelPaths.Resolve(workbook, rangeArg);
            return (target.Sheet, TargetAsRange(target));
        }

        IXLWorksheet sheet;
        if (sheetArg is null)
        {
            sheet = workbook.Worksheets.OrderBy(ws => ws.Position).First();
        }
        else if (!workbook.TryGetWorksheet(sheetArg, out sheet!))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No sheet named '{sheetArg}' exists in the workbook.",
                "Sheet names are matched case-insensitively; pick one of the candidates.",
                candidates: ExcelPaths.SheetCandidates(workbook, sheetArg));
        }

        if (rangeArg is null)
        {
            return (sheet, sheet.RangeUsed());
        }

        var onSheet = ExcelPaths.Resolve(workbook, ExcelPaths.SheetPath(sheet) + "/" + rangeArg.ToUpperInvariant());
        return (sheet, TargetAsRange(onSheet));
    }
}
