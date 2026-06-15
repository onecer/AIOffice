using System.Text;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

public sealed partial class ExcelHandler
{
    private const int DefaultMaxTextBytes = 65536;

    private static readonly IReadOnlyList<string> ReadViews =
        ["outline", "text", "stats", "structure", "csv", "comments", "properties", "styles", "embeds"];

    public Envelope Read(CommandContext ctx) => Run(ctx, sw =>
    {
        var file = RequireFile(ctx, mustExist: true);
        var view = ArgString(ctx, "view") ?? "stats";
        if (!ReadViews.Contains(view, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown view '{view}'.",
                "Use one of: outline, text, stats, structure, csv, comments, properties, styles, embeds.",
                candidates: ReadViews);
        }

        // Big files (or explicit stream=true) answer stats/text via the SAX
        // path without ever loading the workbook DOM.
        var wantStream = ArgBool(ctx, "stream") || ExcelStreaming.IsLarge(file);
        if (wantStream && view == "stats")
        {
            return Envelope.Ok(ExcelStreaming.ReadStats(file), MetaFor(file, sw));
        }

        if (wantStream && view == "text")
        {
            return ReadTextStreamed(ctx, file, sw);
        }

        // Embeds live in raw package parts ClosedXML cannot see; answer this view
        // straight from the package without ever loading the workbook DOM.
        if (view == "embeds")
        {
            return Envelope.Ok(ReadEmbeds(file), MetaFor(file, sw));
        }

        // outline/structure/csv/comments need the full model; on a big file
        // that is the honest slow path, flagged so agents know why.
        List<Warning>? fallback = wantStream
            ? [new Warning(
                "stream_fallback",
                $"Streaming covers --view stats|text; view '{view}' loads the whole workbook into memory.")]
            : null;

        using var workbook = OpenWorkbook(file);
        return view switch
        {
            "stats" => Envelope.Ok(ReadStats(workbook), MetaFor(file, sw, fallback)),
            "outline" => Envelope.Ok(ReadOutline(workbook), MetaFor(file, sw, fallback)),
            "structure" => Envelope.Ok(ReadStructure(workbook, file), MetaFor(file, sw, fallback)),
            "csv" => ReadCsv(ctx, workbook, file, sw),
            "comments" => ReadComments(workbook, file, sw),
            "properties" => Envelope.Ok(ExcelProperties.Describe(workbook), MetaFor(file, sw, fallback)),
            "styles" => Envelope.Ok(
                new { view = "styles", kind = "xlsx", styles = ExcelCellStyles.ListAll(workbook) }, MetaFor(file, sw, fallback)),
            _ => ReadText(ctx, workbook, file, sw),
        };
    });

    private Envelope ReadTextStreamed(CommandContext ctx, string file, System.Diagnostics.Stopwatch sw)
    {
        var maxBytes = ArgInt(ctx, "maxBytes") ?? DefaultMaxTextBytes;
        var (content, truncated) = ExcelStreaming.ReadText(file, ArgString(ctx, "range"), maxBytes);

        List<Warning>? warnings = truncated
            ? [new Warning(
                "result_truncated",
                $"Text output exceeded {maxBytes} bytes and was cut off. Narrow it with --range (e.g. --range A1:F50) or raise --max-bytes.")]
            : null;

        return Envelope.Ok(
            new { view = "text", content, truncated, streamed = true },
            MetaFor(file, sw, warnings));
    }

    /// <summary>
    /// The <c>--view embeds</c> projection: every embedded object in the
    /// workbook, grouped by sheet, in canonical-path order. Read straight from
    /// the raw package (ClosedXML cannot see these parts).
    /// </summary>
    private static object ReadEmbeds(string file)
    {
        var embeds = ExcelEmbeds.ReadAll(file);
        return new
        {
            view = "embeds",
            kind = "xlsx",
            embeds = embeds
                .Select(e => new
                {
                    path = e.Path,
                    sheet = e.SheetName,
                    name = e.Name,
                    mediaType = e.MediaType,
                    size = e.Size,
                    anchor = e.Anchor,
                    source = e.Source,
                })
                .ToList(),
            totals = new { count = embeds.Count },
        };
    }

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
            view = "stats",
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
        view = "outline",
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

    private static object ReadStructure(XLWorkbook workbook, string file)
    {
        // Charts and pivot source ranges live in parts ClosedXML cannot
        // see (or keeps internal); read them raw in one pass.
        List<ChartInfo> allCharts;
        List<SlicerInfo> allSlicers;
        List<ExcelEmbeds.Info> allEmbeds;
        List<FormControlInfo> allFormControls;
        Dictionary<(string, string), (string?, string?)> pivotSources;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(file, isEditable: false))
        {
            allCharts = ExcelCharts.Read(document);
            allSlicers = ExcelSlicers.Read(document);
            allEmbeds = ExcelEmbeds.ReadAll(document);
            allFormControls = ExcelFormControls.Read(document);
            pivotSources = ExcelPivots.ReadSources(document);
        }

        var chartsBySheet = allCharts.ToLookup(c => c.SheetName, StringComparer.OrdinalIgnoreCase);
        var slicersBySheet = allSlicers.ToLookup(s => s.SheetName, StringComparer.OrdinalIgnoreCase);
        var embedsBySheet = allEmbeds.ToLookup(e => e.SheetName, StringComparer.OrdinalIgnoreCase);
        var formControlsBySheet = allFormControls.ToLookup(f => f.SheetName, StringComparer.OrdinalIgnoreCase);
        return new
        {
            view = "structure",
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
                            path = ExcelPaths.TablePath(ws, t.Name),
                            name = t.Name,
                            range = t.RangeAddress.ToString(),
                            style = t.Theme == ClosedXML.Excel.XLTableTheme.None ? "none" : t.Theme.Name,
                            headerRow = t.ShowHeaderRow,
                            totalsRow = t.ShowTotalsRow,
                            bandedRows = t.ShowRowStripes,
                            columns = t.Fields.Select(f => f.Name).ToList(),
                        })
                        .ToList(),
                    charts = chartsBySheet[ws.Name]
                        .Select(c => new
                        {
                            path = c.Path,
                            kind = c.Kind,
                            title = c.Title,
                            dataRange = c.DataRange,
                            anchor = c.Anchor,
                            series = c.Series,
                        })
                        .ToList(),
                    pivots = ws.PivotTables
                        .Select(pt => ExcelPivots.Describe(ws, pt, pivotSources))
                        .ToList(),
                    slicers = slicersBySheet[ws.Name]
                        .Select(s => new
                        {
                            path = s.Path,
                            name = s.Name,
                            source = s.SourceKind,
                            sourceName = s.Source,
                            column = s.Column,
                            caption = s.Caption,
                        })
                        .ToList(),
                    conditionalFormats = ws.ConditionalFormats
                        .Select((cf, i) => ExcelConditionalFormats.Describe(ws, cf, i + 1))
                        .ToList(),
                    dataValidations = ExcelDataValidations.List(ws),
                    sparklineGroups = ExcelSparklines.ListGroups(ws),
                    images = ws.Pictures
                        .Select((pic, i) => ExcelImages.Describe(ws, pic, i + 1))
                        .ToList(),
                    embeds = embedsBySheet[ws.Name]
                        .Select(e => new
                        {
                            path = e.Path,
                            name = e.Name,
                            mediaType = e.MediaType,
                            size = e.Size,
                            anchor = e.Anchor,
                            source = e.Source,
                        })
                        .ToList(),
                    formControls = formControlsBySheet[ws.Name]
                        .Select(f => new
                        {
                            path = f.Path,
                            kind = f.Kind,
                            cell = f.Cell,
                            linkedCell = f.LinkedCell,
                            min = f.Min,
                            max = f.Max,
                            increment = f.Increment,
                        })
                        .ToList(),
                    notes = NoteList(ws),
                    mergedRanges = ws.MergedRanges.Select(r => r.RangeAddress.ToString()).ToList(),
                    autoFilter = ws.AutoFilter.IsEnabled ? ws.AutoFilter.Range?.RangeAddress.ToString() : null,
                    outline = OutlineInfo(ws),
                    // v1.2: protection state (null/omitted when the sheet is unprotected).
                    protection = ExcelProtection.SheetInfo(ws),
                })
                .ToList(),
            definedNames = ExcelNames.ListAll(workbook),
            // v1.2: workbook structure protection (null/omitted when not set).
            workbookProtection = ExcelProtection.WorkbookInfo(workbook),
        };
    }

    /// <summary>
    /// Outline groups on a sheet: the row/column spans that carry an outline
    /// level &gt; 0, contracted from consecutive same-level rows/columns. Null
    /// when the sheet has no grouping (the common case pays nothing extra).
    /// </summary>
    private static object? OutlineInfo(IXLWorksheet ws)
    {
        var rowGroups = GroupSpans(
            ws,
            ws.RowsUsed(XLCellsUsedOptions.All).Select(r => (r.RowNumber(), r.OutlineLevel)),
            (first, last) => string.Create(
                System.Globalization.CultureInfo.InvariantCulture,
                $"{ExcelPaths.SheetPath(ws)}/row[{first}]:row[{last}]"));
        var columnGroups = GroupSpans(
            ws,
            ws.ColumnsUsed(XLCellsUsedOptions.All).Select(c => (c.ColumnNumber(), c.OutlineLevel)),
            (first, last) => string.Create(
                System.Globalization.CultureInfo.InvariantCulture,
                $"{ExcelPaths.SheetPath(ws)}/col[{ExcelCharts.ColumnLetters(first)}]:col[{ExcelCharts.ColumnLetters(last)}]"));

        if (rowGroups.Count == 0 && columnGroups.Count == 0)
        {
            return null;
        }

        return new { rowGroups, columnGroups };
    }

    /// <summary>Contracts consecutive same-level indices into {path, level} spans.</summary>
    private static List<object> GroupSpans(
        IXLWorksheet ws, IEnumerable<(int Index, int Level)> levels, Func<int, int, string> path)
    {
        var spans = new List<object>();
        var ordered = levels.Where(l => l.Level > 0).OrderBy(l => l.Index).ToList();
        var i = 0;
        while (i < ordered.Count)
        {
            var level = ordered[i].Level;
            var first = ordered[i].Index;
            var last = first;
            var j = i + 1;
            while (j < ordered.Count && ordered[j].Level == level && ordered[j].Index == last + 1)
            {
                last = ordered[j].Index;
                j++;
            }

            spans.Add(new { path = path(first, last), outlineLevel = level });
            i = j;
        }

        return spans;
    }

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
        ExcelTargetKind.Column => ColumnUsedRange(target.Sheet, target.ColumnNumber!.Value),
        _ => target.Sheet.RangeUsed(),
    };

    private static IXLRange? RowUsedRange(IXLWorksheet sheet, int rowNumber)
    {
        var lastColumn = sheet.Row(rowNumber).LastCellUsed()?.Address.ColumnNumber;
        return lastColumn is null ? null : sheet.Range(rowNumber, 1, rowNumber, lastColumn.Value);
    }

    private static IXLRange? ColumnUsedRange(IXLWorksheet sheet, int columnNumber)
    {
        var lastRow = sheet.Column(columnNumber).LastCellUsed()?.Address.RowNumber;
        return lastRow is null ? null : sheet.Range(1, columnNumber, lastRow.Value, columnNumber);
    }
}
