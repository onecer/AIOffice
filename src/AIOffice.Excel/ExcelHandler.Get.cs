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

        // "/" (document root): the workbook-level calculation settings (v1.7) plus
        // workbook-structure protection state. (Pre-1.7 this was unsupported_feature;
        // returning the calc/protection block is additive.)
        if (pathArg == "/")
        {
            using var rootWorkbook = OpenWorkbook(file);
            return Envelope.Ok(
                new
                {
                    path = "/",
                    kind = "workbook",
                    calculation = ExcelCalculation.Read(file),
                    workbookProtection = ExcelProtection.WorkbookInfo(rootWorkbook),
                },
                MetaFor(file, sw));
        }

        // Workbook-level get targets the shared grammar cannot parse (no sheet
        // prefix), peeled off before path resolution like defined names.
        if (pathArg is "/properties" or "/styles")
        {
            using var propsWorkbook = OpenWorkbook(file);
            return pathArg == "/properties"
                ? Envelope.Ok(ExcelProperties.Describe(propsWorkbook), MetaFor(file, sw))
                : Envelope.Ok(
                    new { kind = "xlsx", styles = ExcelCellStyles.ListAll(propsWorkbook) }, MetaFor(file, sw));
        }

        if (ExcelCellStyles.TryParsePath(pathArg, out var styleName))
        {
            using var styleWorkbook = OpenWorkbook(file);
            return Envelope.Ok(ExcelCellStyles.Get(styleWorkbook, styleName), MetaFor(file, sw));
        }

        // Cell/range gets on big files (or with stream=true) are served by the
        // SAX path without loading the workbook DOM. Other targets (sheet,
        // chart, pivot, name…) still need the full model.
        if ((ArgBool(ctx, "stream") || ExcelStreaming.IsLarge(file)) &&
            ExcelStreaming.TryParseCellOrRange(pathArg, out var streamSheet, out var streamStart, out var streamEnd))
        {
            if (streamStart == streamEnd)
            {
                return Envelope.Ok(ExcelStreaming.GetCell(file, streamSheet, streamStart), MetaFor(file, sw));
            }

            var maxCells = ArgInt(ctx, "maxCells") ?? DefaultMaxRangeCells;
            var result = ExcelStreaming.GetRange(file, streamSheet, streamStart, streamEnd, maxCells);
            List<Warning>? streamWarnings = result.TotalRows > result.EmittedRows
                ? [new Warning(
                    "result_truncated",
                    $"Range has {result.TotalRows} rows; returning the first {result.EmittedRows}. Request a smaller range or raise maxCells.")]
                : null;
            return Envelope.Ok(result.Data, MetaFor(file, sw, streamWarnings));
        }

        using var workbook = OpenWorkbook(file);
        if (ExcelNames.TryParsePath(pathArg, out var nameSheetPath, out var definedName))
        {
            var (found, scopeSheet) = ExcelNames.Find(workbook, nameSheetPath, definedName);
            return Envelope.Ok(ExcelNames.Describe(found, scopeSheet), MetaFor(file, sw));
        }

        var target = ExcelPaths.Resolve(workbook, pathArg);

        return target.Kind switch
        {
            ExcelTargetKind.Sheet => Envelope.Ok(SheetInfo(target.Sheet, file), MetaFor(file, sw)),
            ExcelTargetKind.Cell => Envelope.Ok(
                CellInfo(target.Sheet, target.Cell!, CommentInfoFor(file, target)), MetaFor(file, sw)),
            ExcelTargetKind.Row => Envelope.Ok(RowInfo(target.Sheet, target.RowNumber!.Value), MetaFor(file, sw)),
            ExcelTargetKind.Column => Envelope.Ok(
                ColumnInfo(target.Sheet, target.ColumnNumber!.Value), MetaFor(file, sw)),
            ExcelTargetKind.Chart => Envelope.Ok(ChartTargetInfo(file, target), MetaFor(file, sw)),
            ExcelTargetKind.Pivot => Envelope.Ok(PivotTargetInfo(file, target), MetaFor(file, sw)),
            ExcelTargetKind.ConditionalFormat => Envelope.Ok(
                ConditionalFormatTargetInfo(file, target), MetaFor(file, sw)),
            ExcelTargetKind.Image => Envelope.Ok(
                ExcelImages.Describe(target.Sheet, ExcelImages.Find(target), target.ImageIndex!.Value),
                MetaFor(file, sw)),
            ExcelTargetKind.LinkedPicture => Envelope.Ok(
                ExcelLinkedPicture.Describe(workbook, target), MetaFor(file, sw)),
            ExcelTargetKind.DataValidation => Envelope.Ok(
                ExcelDataValidations.Describe(
                    target.Sheet, ExcelDataValidations.Find(target), target.DataValidationIndex!.Value),
                MetaFor(file, sw)),
            ExcelTargetKind.Sparkline => Envelope.Ok(
                ExcelSparklines.Describe(target.Sheet, ExcelSparklines.Find(target), target.SparklineIndex!.Value),
                MetaFor(file, sw)),
            ExcelTargetKind.Comment => Envelope.Ok(CommentTargetInfo(file, target), MetaFor(file, sw)),
            ExcelTargetKind.Table => Envelope.Ok(
                ExcelTables.Describe(target.Sheet, ExcelTables.Find(target)), MetaFor(file, sw)),
            ExcelTargetKind.Slicer => Envelope.Ok(SlicerTargetInfo(file, target), MetaFor(file, sw)),
            ExcelTargetKind.Embed => Envelope.Ok(
                ExcelEmbeds.Describe(ExcelEmbeds.Resolve(file, target)), MetaFor(file, sw)),
            ExcelTargetKind.FormControl => Envelope.Ok(FormControlTargetInfo(file, target), MetaFor(file, sw)),
            ExcelTargetKind.DataTable => Envelope.Ok(DataTableTargetInfo(file, target), MetaFor(file, sw)),
            ExcelTargetKind.Scenario => Envelope.Ok(ScenarioTargetInfo(file, target), MetaFor(file, sw)),
            _ => RangeInfo(ctx, target, file, sw),
        };
    });

    /// <summary>One comment thread, read from the raw threadedComments part.</summary>
    private static object CommentTargetInfo(string file, ExcelTarget target)
    {
        var model = ExcelComments.Load(file);
        var thread = model.FindById(target.Sheet.Name, target.CommentId!);
        if (thread is null)
        {
            var candidates = ExcelComments.CandidatesOn(target.Sheet, model);
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No comment thread with id '{target.CommentId}' on sheet '{target.Sheet.Name}'.",
                candidates.Count > 0
                    ? "Run 'aioffice read --view comments' to list thread paths; pick one of the candidates."
                    : "This sheet has no comment threads; start one with {op:add, type:comment, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + "/B2, props:{text:\"…\"}}.",
                candidates: candidates.Count > 0 ? candidates : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return ExcelComments.Describe(target.Sheet, thread, model);
    }

    /// <summary>
    /// The comment block for cell get (M5): present only when the cell carries
    /// a threaded comment. Detection rides on the cheap ClosedXML shadow check
    /// so plain cells never pay for a raw part read.
    /// </summary>
    private static object? CommentInfoFor(string file, ExcelTarget target)
    {
        if (!ExcelComments.HasShadow(target.Cell!))
        {
            return null;
        }

        var model = ExcelComments.Load(file);
        var thread = model.FindByCell(target.Sheet.Name, target.Cell!.Address.ToString()!);
        return thread is null ? null : ExcelComments.Describe(target.Sheet, thread, model);
    }

    /// <summary>One pivot table; the source range and calculated fields come from the raw cache part.</summary>
    private static object PivotTargetInfo(string file, ExcelTarget target)
    {
        var (pivot, _) = ExcelPivots.Find(target);
        Dictionary<(string, string), (string?, string?)> sources;
        Dictionary<(string, string), List<(string, string)>> calculatedFields;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(file, isEditable: false))
        {
            sources = ExcelPivots.ReadSources(document);
            calculatedFields = ExcelPivots.ReadCalculatedFields(document);
        }

        return ExcelPivots.Describe(target.Sheet, pivot, sources, calculatedFields);
    }

    /// <summary>One slicer, read back from the raw package (ClosedXML cannot see slicers).</summary>
    private static object SlicerTargetInfo(string file, ExcelTarget target)
    {
        List<SlicerInfo> slicers;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(file, isEditable: false))
        {
            slicers = ExcelSlicers.Read(document)
                .Where(s => string.Equals(s.SheetName, target.Sheet.Name, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        var info = target.SlicerName is { } wantedName
            ? slicers.FirstOrDefault(s => string.Equals(s.Name, wantedName, StringComparison.OrdinalIgnoreCase))
            : slicers.FirstOrDefault(s => s.Index == target.SlicerIndex!.Value);
        if (info is null)
        {
            var what = target.SlicerName is { } n ? $"slicer[@name={n}]" : $"slicer[{target.SlicerIndex}]";
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No {what} on sheet '{target.Sheet.Name}' ({slicers.Count} slicer(s) exist).",
                slicers.Count > 0
                    ? "Slicer indices are 1-based per sheet; pick one of the candidates."
                    : "This sheet has no slicers; add one with edit op {op:add, type:slicer, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + "/table[@name=Sales], props:{column:\"Region\"}}.",
                candidates: slicers.Count > 0
                    ? [.. slicers.Select(s => s.Path)]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return ExcelSlicers.Describe(info);
    }

    /// <summary>One what-if data table, read back raw (the t="dataTable" anchor; ClosedXML cannot see it).</summary>
    private static object DataTableTargetInfo(string file, ExcelTarget target)
    {
        var tables = ExcelDataTables.ReadOnSheet(file, target.Sheet.Name);
        var wanted = target.DataTableIndex!.Value;
        var info = tables.FirstOrDefault(t => t.Index == wanted);
        if (info is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No dataTable[{wanted}] on sheet '{target.Sheet.Name}' ({tables.Count} data table(s) exist).",
                tables.Count > 0
                    ? "Data-table indices are 1-based per sheet; pick one of the candidates."
                    : "This sheet has no data tables; add one with {op:add, type:dataTable, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + "/A1:C10, props:{rowInput:\"B1\", colInput:\"B2\"}}.",
                candidates: tables.Count > 0
                    ? [.. tables.Select(t => ExcelPaths.DataTablePath(target.Sheet, t.Index))]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return ExcelDataTables.Describe(info);
    }

    /// <summary>One scenario, read back raw from the worksheet's scenarios part (ClosedXML cannot see it).</summary>
    private static object ScenarioTargetInfo(string file, ExcelTarget target)
    {
        var scenarios = ExcelScenarios.ReadOnSheet(file, target.Sheet.Name);
        var wanted = target.ScenarioName!;
        var info = scenarios.FirstOrDefault(s => string.Equals(s.Name, wanted, StringComparison.OrdinalIgnoreCase));
        if (info is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No scenario named '{wanted}' on sheet '{target.Sheet.Name}' ({scenarios.Count} scenario(s) exist).",
                scenarios.Count > 0
                    ? "Pick one of the candidates; scenario names are case-insensitive."
                    : "This sheet has no scenarios; add one with {op:add, type:scenario, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + ", props:{name:\"Best Case\", cells:{\"B1\":120}}}.",
                candidates: scenarios.Count > 0
                    ? [.. scenarios.Select(s => ExcelScenarios.ScenarioPath(target.Sheet, s.Name))]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return ExcelScenarios.Describe(target.Sheet, info);
    }

    /// <summary>One form control, read back from the raw package (ClosedXML cannot see controls).</summary>
    private static object FormControlTargetInfo(string file, ExcelTarget target)
    {
        List<FormControlInfo> controls;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(file, isEditable: false))
        {
            controls = ExcelFormControls.Read(document)
                .Where(c => string.Equals(c.SheetName, target.Sheet.Name, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        var wanted = target.FormControlIndex!.Value;
        var info = controls.FirstOrDefault(c => c.Index == wanted);
        if (info is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No formControl[{wanted}] on sheet '{target.Sheet.Name}' ({controls.Count} control(s) exist).",
                controls.Count > 0
                    ? "Form-control indices are 1-based per sheet; pick one of the candidates."
                    : "This sheet has no form controls; add one with edit op {op:add, type:formControl, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + ", props:{kind:\"checkbox\", cell:\"E2\", linkedCell:\"F2\"}}.",
                candidates: controls.Count > 0
                    ? [.. controls.Select(c => c.Path)]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return new
        {
            path = info.Path,
            kind = "formControl",
            sheet = target.Sheet.Name,
            controlKind = info.Kind,
            cell = info.Cell,
            linkedCell = info.LinkedCell,
            min = info.Min,
            max = info.Max,
            increment = info.Increment,
        };
    }

    /// <summary>
    /// One conditional-format rule. The aboveBelowAverage kind's mode/stdDev are
    /// read raw (ClosedXML cannot see those rule attributes); every other kind is
    /// fully described from the ClosedXML model.
    /// </summary>
    private static object ConditionalFormatTargetInfo(string file, ExcelTarget target)
    {
        var format = ExcelConditionalFormats.Find(target);
        var index = target.ConditionalFormatIndex!.Value;
        if (format.ConditionalFormatType != ClosedXML.Excel.XLConditionalFormatType.AboveAverage)
        {
            return ExcelConditionalFormats.Describe(target.Sheet, format, index);
        }

        IReadOnlyDictionary<int, ExcelConditionalFormats.AverageRuleDetail> details;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(file, isEditable: false))
        {
            details = ExcelConditionalFormats.ReadAverageDetails(document, target.Sheet.Name);
        }

        details.TryGetValue(index, out var detail);
        return ExcelConditionalFormats.Describe(target.Sheet, format, index, detail);
    }

    /// <summary>One chart, read back from the raw package (ClosedXML cannot see charts).</summary>
    private static object ChartTargetInfo(string file, ExcelTarget target)
    {
        var wanted = target.ChartIndex!.Value;
        ChartInfo? info;
        object? polish = null;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(file, isEditable: false))
        {
            var charts = ExcelCharts.Read(document)
                .Where(c => string.Equals(c.SheetName, target.Sheet.Name, StringComparison.OrdinalIgnoreCase))
                .ToList();
            info = charts.FirstOrDefault(c => c.Index == wanted);
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

            // Read the chart's v1.3 polish settings (data labels, legend, …) from
            // the same raw part.
            if (ExcelChartPolish.ChartPartFor(document, target.Sheet.Name, wanted)
                    ?.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>() is { } chart)
            {
                polish = ExcelChartPolish.DescribePolish(chart);
            }
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
            polish,
        };
    }

    /// <summary>The number of linked pictures registered on a sheet (0 when none/absent).</summary>
    private static int? LinkedPictureCount(IXLWorksheet sheet)
    {
        var count = ExcelLinkedPicture.CountOnSheet(sheet.Workbook, sheet);
        return count > 0 ? count : (int?)null;
    }

    private static object SheetInfo(IXLWorksheet sheet, string file)
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
            freezeRows = sheet.SheetView.SplitRow > 0 ? sheet.SheetView.SplitRow : (int?)null,
            freezeCols = sheet.SheetView.SplitColumn > 0 ? sheet.SheetView.SplitColumn : (int?)null,
            autoFilter = sheet.AutoFilter.IsEnabled ? sheet.AutoFilter.Range?.RangeAddress.ToString() : null,
            // Embedded objects live in raw package parts ClosedXML cannot see;
            // surface a count so 'get' on a sheet hints at them (polish, M10).
            embeds = ExcelEmbeds.ListOnSheet(file, sheet.Name).Count,
            // (1.7) Linked pictures (camera tool) count, so 'get' on a sheet hints
            // at them (they also appear under image[i] as plain pictures).
            linkedPictures = LinkedPictureCount(sheet),
            pageSetup = PageSetupInfo(sheet, file),
            // v1.2: protection state (null/omitted when the sheet is unprotected),
            // plus the workbook-structure protection (omitted when not set).
            protection = ExcelProtection.SheetInfo(sheet),
            workbookProtection = ExcelProtection.WorkbookInfo(sheet.Workbook),
        };
    }

    /// <summary>One cell with its typed value and the properties an agent needs to edit it.</summary>
    private static object CellInfo(IXLWorksheet sheet, IXLCell cell, object? comment = null)
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
        var hyperlink = cell.HasHyperlink ? cell.GetHyperlink() : null;
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
            // (1.4) When the cell is a dynamic-array anchor, surface the rectangle
            // the result spills into (read off the array formula's ref); a plain
            // single-cell formula reports null.
            spillRange = SpillRangeOf(cell),
            text = ExcelValues.SafeFormatted(cell),
            numberFormat = string.IsNullOrEmpty(numberFormat) ? null : numberFormat,
            bold = cell.Style.Font.Bold ? true : (bool?)null,
            italic = cell.Style.Font.Italic ? true : (bool?)null,
            // v1.2: only surface 'locked' when it deviates from the OOXML default
            // (locked) — i.e. report an explicitly UNLOCKED cell so an editable
            // window under sheet protection is visible.
            locked = ExcelProtection.IsLocked(cell) ? (bool?)null : false,
            merged = MergedRangeOf(sheet, cell),
            hyperlink = hyperlink is null
                ? null
                : hyperlink.IsExternal
                    ? hyperlink.ExternalAddress.ToString()
                    : "#" + hyperlink.InternalAddress,
            hyperlinkTooltip = string.IsNullOrEmpty(hyperlink?.Tooltip) ? null : hyperlink.Tooltip,
            note = NoteInfo(cell),
            comment,
        };
    }

    /// <summary>
    /// (1.4) The spill rectangle of a dynamic-array anchor (FILTER/UNIQUE/SORT/…):
    /// the array formula's <c>ref</c> when it spans more than the anchor cell,
    /// otherwise null (a plain single-cell formula does not spill).
    /// </summary>
    private static string? SpillRangeOf(IXLCell cell)
    {
        if (!cell.HasFormula || cell.FormulaReference is not { } reference)
        {
            return null;
        }

        var single = reference.FirstAddress.RowNumber == reference.LastAddress.RowNumber &&
                     reference.FirstAddress.ColumnNumber == reference.LastAddress.ColumnNumber;
        return single ? null : reference.ToString();
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
        var row = sheet.Row(rowNumber);
        var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? 0;
        var values = new List<object?>(lastColumn);
        for (var column = 1; column <= lastColumn; column++)
        {
            values.Add(ExcelValues.ToJson(sheet.Cell(rowNumber, column).Value));
        }

        return new
        {
            path = ExcelPaths.RowPath(sheet, rowNumber),
            kind = "row",
            sheet = sheet.Name,
            row = rowNumber,
            values,
            height = row.Height,
            hidden = row.IsHidden ? true : (bool?)null,
            outlineLevel = row.OutlineLevel > 0 ? row.OutlineLevel : (int?)null,
        };
    }

    private static object ColumnInfo(IXLWorksheet sheet, int columnNumber)
    {
        var column = sheet.Column(columnNumber);
        var lastRow = column.LastCellUsed()?.Address.RowNumber ?? 0;
        var values = new List<object?>(lastRow);
        for (var row = 1; row <= lastRow; row++)
        {
            values.Add(ExcelValues.ToJson(sheet.Cell(row, columnNumber).Value));
        }

        return new
        {
            path = ExcelPaths.ColumnPath(sheet, columnNumber),
            kind = "col",
            sheet = sheet.Name,
            column = ExcelCharts.ColumnLetters(columnNumber),
            values,
            width = column.Width,
            hidden = column.IsHidden ? true : (bool?)null,
            outlineLevel = column.OutlineLevel > 0 ? column.OutlineLevel : (int?)null,
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
