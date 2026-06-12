using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

public sealed partial class ExcelHandler
{
    private static readonly IReadOnlyList<string> SetProps =
    [
        "value", "valueType", "values", "numberFormat", "bold", "italic", "fill", "merge", "name",
        "freezeRows", "freezeCols", "autoFilter", "orientation", "paperSize", "fitToWidth", "printArea",
    ];

    private static readonly IReadOnlyList<string> AddTypes =
        ["sheet", "table", "row", "chart", "pivot", "conditionalFormat", "image", "name"];

    public Envelope Edit(CommandContext ctx, IReadOnlyList<EditOp> ops) => Run(ctx, sw =>
    {
        var file = RequireFile(ctx, mustExist: true);

        var expectRev = ArgString(ctx, "expectRev");
        if (expectRev is not null)
        {
            var current = Core.Rev.OfFile(file);
            if (!string.Equals(current, expectRev, StringComparison.OrdinalIgnoreCase))
            {
                throw new AiofficeException(
                    ErrorCodes.StaleAddress,
                    $"The file changed since it was read: expected rev {expectRev}, but it is now {current}.",
                    "Re-run 'aioffice read' or 'aioffice query' to refresh paths, then retry with the new rev.");
            }
        }

        using var workbook = OpenWorkbook(file);

        // Apply every op in memory first; any failure aborts before a byte is
        // written. Chart ops are validated here too, but their raw OpenXml
        // write is deferred to a post-save pass (ClosedXML cannot author chart
        // parts; measured: it preserves them byte-identical once present).
        var charts = new ChartOpBatch(file);
        var post = new PostSaveWork();
        var details = new List<object>(ops.Count);
        for (var i = 0; i < ops.Count; i++)
        {
            ApplyOp(ctx, workbook, ops[i], i, details, charts, post);
        }

        if (ArgBool(ctx, "dryRun"))
        {
            return Envelope.Ok(
                new { dryRun = true, wouldApply = details.Count, ops = details },
                MetaFor(file, sw));
        }

        var snapshot = _snapshots.Save(file); // pre-image: every successful edit is undoable
        var warnings = SaveWithCachedValues(workbook, file);
        try
        {
            if (!charts.IsEmpty)
            {
                ExcelCharts.Apply(file, charts.Ops);
            }

            if (post.SyncPivotParts)
            {
                // ClosedXML's save leaves removed pivots' parts behind; sync
                // the package to the in-memory model (the source of truth).
                ExcelPivots.SyncPartsAfterSave(file, ExcelPivots.AliveNames(workbook));
            }

            if (post.RemovedImages.Count > 0)
            {
                ExcelImages.RemoveAfterSave(file, post.RemovedImages);
            }

            // Correct ClosedXML's data-bar GUID/orphan defects on the saved
            // bytes (no-op scan when there is nothing to fix).
            ExcelConditionalFormats.FixUpAfterSave(file);
        }
        catch (Exception)
        {
            // Keep the batch atomic: the ClosedXML half already hit disk,
            // so roll the file back to its pre-edit bytes before failing.
            File.Copy(snapshot.Path, file, overwrite: true);
            throw;
        }

        return Envelope.Ok(
            new { applied = details.Count, ops = details },
            MetaFor(file, sw, warnings));
    });

    /// <summary>Raw clean-up passes an edit batch queues for after the ClosedXML save.</summary>
    private sealed class PostSaveWork
    {
        /// <summary>Set when a pivot was removed (its parts outlive ClosedXML's save).</summary>
        public bool SyncPivotParts { get; set; }

        /// <summary>
        /// Pictures queued for raw removal, in batch order. The ClosedXML model
        /// is left untouched (its own picture deletion mangles the drawing
        /// relationships on save), so these are deleted from the raw package.
        /// </summary>
        public List<(string Sheet, string Name)> RemovedImages { get; } = [];

        /// <summary>Names already queued for removal on one sheet (batch projection).</summary>
        public IReadOnlySet<string> RemovedImageNames(string sheetName) =>
            RemovedImages
                .Where(r => string.Equals(r.Sheet, sheetName, StringComparison.OrdinalIgnoreCase))
                .Select(r => r.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private static void ApplyOp(
        CommandContext ctx, XLWorkbook workbook, EditOp op, int index, List<object> details,
        ChartOpBatch charts, PostSaveWork post)
    {
        switch (op.Op)
        {
            case "set":
                ApplySet(workbook, op, index, details);
                break;
            case "add":
                ApplyAdd(ctx, workbook, op, index, details, charts, post);
                break;
            case "remove":
                ApplyRemove(workbook, op, index, details, charts, post);
                break;
            default: // "move" — ParseBatch already rejected anything else
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"ops[{index}]: move is not supported for xlsx yet.",
                    "Copy the content to the destination with get + set, then remove the source.");
        }
    }

    // ----- set --------------------------------------------------------------

    private static void ApplySet(XLWorkbook workbook, EditOp op, int index, List<object> details)
    {
        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: set needs props.",
                "Pass props like {\"value\":42}, {\"value\":\"=SUM(A1:A2)\"} or {\"bold\":true}.");
        }

        foreach (var (key, _) in props)
        {
            if (!SetProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: unknown set prop '{key}'.",
                    "Supported props: " + string.Join(", ", SetProps) + ".",
                    candidates: SetProps);
            }
        }

        if (ExcelNames.TryParsePath(op.Path, out _, out _))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{index}]: set on a defined name is not supported yet.",
                "Remove the name and add it again pointing at the new range: " +
                "{op:remove, path:" + op.Path + "} then {op:add, type:name, path:/Sheet1/A1:B5, props:{name:…}}.");
        }

        var target = ExcelPaths.Resolve(workbook, op.Path);
        if (target.Kind is ExcelTargetKind.Chart or ExcelTargetKind.Pivot
            or ExcelTargetKind.ConditionalFormat or ExcelTargetKind.Image)
        {
            var kindName = target.Kind == ExcelTargetKind.ConditionalFormat
                ? "conditionalFormat"
                : target.Kind.ToString().ToLowerInvariant();
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{index}]: set on a {kindName} is not supported yet.",
                $"Remove the {kindName} and add it again with the new props: " +
                "{op:remove, path:" + op.Path + "} then {op:add, type:" + kindName + ", ...}.");
        }

        var applied = new List<string>();

        if (props.TryGetPropertyValue("value", out var valueNode))
        {
            SetValue(target, valueNode, ValueTypeOf(props), applied);
        }

        if (props.TryGetPropertyValue("values", out var valuesNode))
        {
            SetValues(target, valuesNode, index, applied);
        }

        if (props.TryGetPropertyValue("numberFormat", out var formatNode) && formatNode is not null)
        {
            StyleOf(target).NumberFormat.Format = formatNode.GetValue<string>();
            applied.Add("numberFormat");
        }

        if (props.TryGetPropertyValue("bold", out var boldNode) && boldNode is not null)
        {
            StyleOf(target).Font.Bold = boldNode.GetValue<bool>();
            applied.Add("bold");
        }

        if (props.TryGetPropertyValue("italic", out var italicNode) && italicNode is not null)
        {
            StyleOf(target).Font.Italic = italicNode.GetValue<bool>();
            applied.Add("italic");
        }

        if (props.TryGetPropertyValue("fill", out var fillNode) && fillNode is not null)
        {
            StyleOf(target).Fill.BackgroundColor = ParseColor(fillNode.GetValue<string>());
            applied.Add("fill");
        }

        if (props.TryGetPropertyValue("merge", out var mergeNode) && mergeNode is not null)
        {
            if (target.Kind != ExcelTargetKind.Range)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: merge applies to a range, not {target.Kind.ToString().ToLowerInvariant()} '{op.Path}'.",
                    "Target a range like /Sheet1/A1:B2 to merge or unmerge cells.");
            }

            if (mergeNode.GetValue<bool>())
            {
                target.Range!.Merge();
                applied.Add("merge");
            }
            else
            {
                target.Range!.Unmerge();
                applied.Add("unmerge");
            }
        }

        if (props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null)
        {
            if (target.Kind != ExcelTargetKind.Sheet)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: 'name' renames a sheet; target a sheet path like /Sheet1.",
                    "Use {op:set, path:/OldName, props:{name:\"NewName\"}} to rename.");
            }

            RenameSheet(target.Sheet, nameNode.GetValue<string>());
            applied.Add("name");
        }

        if (props.TryGetPropertyValue("freezeRows", out var freezeRowsNode) && freezeRowsNode is not null)
        {
            ApplyFreeze(target, op, freezeRowsNode, "freezeRows", index, applied);
        }

        if (props.TryGetPropertyValue("freezeCols", out var freezeColsNode) && freezeColsNode is not null)
        {
            ApplyFreeze(target, op, freezeColsNode, "freezeCols", index, applied);
        }

        if (props.TryGetPropertyValue("autoFilter", out var autoFilterNode) && autoFilterNode is not null)
        {
            ApplyAutoFilter(target, op, autoFilterNode, index, applied);
        }

        if (props.TryGetPropertyValue("orientation", out var orientationNode) && orientationNode is not null)
        {
            ApplyOrientation(target, op, orientationNode, index, applied);
        }

        if (props.TryGetPropertyValue("paperSize", out var paperSizeNode) && paperSizeNode is not null)
        {
            ApplyPaperSize(target, op, paperSizeNode, index, applied);
        }

        if (props.TryGetPropertyValue("fitToWidth", out var fitToWidthNode) && fitToWidthNode is not null)
        {
            ApplyFitToWidth(target, op, fitToWidthNode, index, applied);
        }

        if (props.TryGetPropertyValue("printArea", out var printAreaNode) && printAreaNode is not null)
        {
            ApplyPrintArea(target, op, printAreaNode, index, applied);
        }

        details.Add(new { op = "set", path = DocPath.Parse(op.Path).ToCanonicalString(), applied });
    }

    private static string? ValueTypeOf(JsonObject props) =>
        props.TryGetPropertyValue("valueType", out var node) && node is not null
            ? node.GetValue<string>()
            : null;

    private static void SetValue(ExcelTarget target, JsonNode? valueNode, string? valueType, List<string> applied)
    {
        var parsed = ExcelValues.Parse(valueNode, valueType);
        IEnumerable<IXLCell> cells = target.Kind switch
        {
            ExcelTargetKind.Cell => [target.Cell!],
            ExcelTargetKind.Range => target.Range!.Cells(),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "'value' targets a cell or range.",
                "For rows use the 'values' array prop; for sheets there is nothing to set a value on."),
        };

        foreach (var cell in cells)
        {
            if (parsed.IsFormula)
            {
                cell.FormulaA1 = parsed.Formula!;
            }
            else
            {
                cell.Value = parsed.Value;
            }
        }

        applied.Add(parsed.IsFormula ? "formula" : "value");
    }

    private static void SetValues(ExcelTarget target, JsonNode? valuesNode, int index, List<string> applied)
    {
        if (valuesNode is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: 'values' must be a JSON array.",
                "Rows take a flat array ([1,\"a\",true]); ranges take a 2D array ([[1,2],[3,4]]).");
        }

        switch (target.Kind)
        {
            case ExcelTargetKind.Row:
                WriteRowValues(target.Sheet, target.RowNumber!.Value, array);
                break;

            case ExcelTargetKind.Range:
            {
                var address = target.Range!.RangeAddress;
                var rangeRows = address.LastAddress.RowNumber - address.FirstAddress.RowNumber + 1;
                var rangeColumns = address.LastAddress.ColumnNumber - address.FirstAddress.ColumnNumber + 1;
                if (array.Count > rangeRows)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{index}]: 'values' has {array.Count} rows but the range has only {rangeRows}.",
                        "Match the range size, or target a larger range.");
                }

                for (var r = 0; r < array.Count; r++)
                {
                    if (array[r] is not JsonArray rowArray)
                    {
                        throw new AiofficeException(
                            ErrorCodes.InvalidArgs,
                            $"ops[{index}]: 'values' on a range must be a 2D array; row {r + 1} is not an array.",
                            "Pass [[1,2],[3,4]] — one inner array per row.");
                    }

                    if (rowArray.Count > rangeColumns)
                    {
                        throw new AiofficeException(
                            ErrorCodes.InvalidArgs,
                            $"ops[{index}]: row {r + 1} has {rowArray.Count} values but the range has only {rangeColumns} columns.",
                            "Match the range size, or target a larger range.");
                    }

                    for (var c = 0; c < rowArray.Count; c++)
                    {
                        WriteParsed(
                            target.Sheet.Cell(address.FirstAddress.RowNumber + r, address.FirstAddress.ColumnNumber + c),
                            ExcelValues.Parse(rowArray[c]));
                    }
                }

                break;
            }

            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: 'values' targets a row or range.",
                    "For a single cell use the 'value' prop instead.");
        }

        applied.Add("values");
    }

    private static void WriteRowValues(IXLWorksheet sheet, int rowNumber, JsonArray values)
    {
        for (var i = 0; i < values.Count; i++)
        {
            WriteParsed(sheet.Cell(rowNumber, i + 1), ExcelValues.Parse(values[i]));
        }
    }

    private static void WriteParsed(IXLCell cell, ExcelValues.ParsedValue parsed)
    {
        if (parsed.IsFormula)
        {
            cell.FormulaA1 = parsed.Formula!;
        }
        else
        {
            cell.Value = parsed.Value;
        }
    }

    private static IXLStyle StyleOf(ExcelTarget target) => target.Kind switch
    {
        ExcelTargetKind.Cell => target.Cell!.Style,
        ExcelTargetKind.Range => target.Range!.Style,
        ExcelTargetKind.Row => target.Sheet.Row(target.RowNumber!.Value).Style,
        _ => throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "Style props (numberFormat, bold, italic, fill) target a cell, range or row.",
            "Address a cell like /Sheet1/A1 or a range like /Sheet1/A1:C10."),
    };

    private static XLColor ParseColor(string html)
    {
        try
        {
            return XLColor.FromHtml(html);
        }
        catch (Exception exception)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{html}' is not a recognizable color.",
                "Use hex like #FFEE00 (or #AARRGGBB for alpha).",
                innerException: exception);
        }
    }

    private static void RenameSheet(IXLWorksheet sheet, string newName)
    {
        try
        {
            sheet.Name = newName;
        }
        catch (ArgumentException exception)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{newName}' is not a usable sheet name: {exception.Message}",
                @"Sheet names are 1-31 characters and cannot contain : \ / ? * [ ].",
                innerException: exception);
        }
    }

    // ----- add --------------------------------------------------------------

    private static void ApplyAdd(
        CommandContext ctx, XLWorkbook workbook, EditOp op, int index, List<object> details,
        ChartOpBatch charts, PostSaveWork post)
    {
        switch (op.Type)
        {
            case "sheet":
                AddSheet(workbook, op, index, details);
                break;
            case "table":
                AddTable(workbook, op, index, details);
                break;
            case "row":
                AddRow(workbook, op, index, details);
                break;
            case "chart":
                AddChart(workbook, op, index, details, charts);
                break;
            case "pivot":
                AddPivot(workbook, op, index, details);
                break;
            case "conditionalFormat":
                details.Add(ExcelConditionalFormats.Add(ExcelPaths.Resolve(workbook, op.Path), op, index));
                break;
            case "name":
                details.Add(ExcelNames.Add(workbook, ExcelPaths.Resolve(workbook, op.Path), op, index));
                break;
            case "image":
            {
                var target = ExcelPaths.Resolve(workbook, op.Path);
                var pending = target.Kind == ExcelTargetKind.Sheet
                    ? post.RemovedImageNames(target.Sheet.Name)
                    : new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                details.Add(ExcelImages.Add(ctx.Workspace, target, op, index, pending));
                break;
            }
            case null:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: add needs a type.",
                    "Supported xlsx adds: " + string.Join(", ", AddTypes) + ".",
                    candidates: AddTypes);
            default:
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"ops[{index}]: add type '{op.Type}' is not supported for xlsx yet.",
                    "Supported adds: " + string.Join(", ", AddTypes) + ".",
                    candidates: AddTypes);
        }
    }

    /// <summary>
    /// Validates and applies an <c>add pivot</c> op. The op's path names the
    /// SOURCE sheet (where sourceRange lives); props.targetSheet says where the
    /// pivot lands (created when absent).
    /// </summary>
    private static void AddPivot(XLWorkbook workbook, EditOp op, int index, List<object> details)
    {
        var target = ExcelPaths.Resolve(workbook, op.Path);
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add pivot targets the source sheet path like /Sheet1; " +
                "the data location goes in props.sourceRange.",
                "Use {op:add, type:pivot, path:/Sheet1, props:{sourceRange:\"A1:D20\", targetSheet:\"Pivot\", " +
                "rows:[\"Region\"], values:[{field:\"Sales\", agg:\"sum\"}]}}.");
        }

        details.Add(ExcelPivots.Add(workbook, target.Sheet, op, index));
    }

    private static void AddSheet(XLWorkbook workbook, EditOp op, int index, List<object> details)
    {
        var path = DocPath.Parse(op.Path);
        var first = path.Segments[0];
        var name = path.Segments.Count == 1 && first.Kind is PathSegmentKind.Name or PathSegmentKind.Element && first.Index is null
            ? first.Name
            : null;
        if (name is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add sheet takes the new sheet name as the path, e.g. /Summary.",
                "Quote names with spaces: /'Q3 Data'.");
        }

        if (workbook.TryGetWorksheet(name, out _))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: a sheet named '{name}' already exists.",
                "Pick a different name, or rename the existing sheet first with {op:set, props:{name:...}}.");
        }

        AddSheetOrThrow(workbook, name);
        details.Add(new { op = "add", type = "sheet", path = path.ToCanonicalString() });
    }

    private static void AddTable(XLWorkbook workbook, EditOp op, int index, List<object> details)
    {
        var target = ExcelPaths.Resolve(workbook, op.Path);
        if (target.Kind != ExcelTargetKind.Range)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add table needs a range path like /Sheet1/A1:C10 (first row = headers).",
                "Address the data including its header row.");
        }

        var name = op.Props?.TryGetPropertyValue("name", out var nameNode) == true && nameNode is not null
            ? nameNode.GetValue<string>()
            : null;
        try
        {
            var table = name is null ? target.Range!.CreateTable() : target.Range!.CreateTable(name);
            details.Add(new
            {
                op = "add",
                type = "table",
                path = DocPath.Parse(op.Path).ToCanonicalString(),
                name = table.Name,
                columns = table.Fields.Select(f => f.Name).ToList(),
            });
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: could not create the table: {exception.Message}",
                "Table names must be unique in the workbook and the range must not overlap an existing table.",
                innerException: exception);
        }
    }

    private static void AddRow(XLWorkbook workbook, EditOp op, int index, List<object> details)
    {
        var target = ExcelPaths.Resolve(workbook, op.Path);
        int rowNumber;
        switch (target.Kind)
        {
            case ExcelTargetKind.Sheet: // append after the used range
                rowNumber = (target.Sheet.LastRowUsed()?.RowNumber() ?? 0) + 1;
                break;
            case ExcelTargetKind.Row: // insert at this position, pushing rows down
                rowNumber = target.RowNumber!.Value;
                target.Sheet.Row(rowNumber).InsertRowsAbove(1);
                break;
            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: add row targets a sheet (/Sheet1 appends) or a row (/Sheet1/row[3] inserts).",
                    "Pass values via props, e.g. {\"values\":[1,\"a\",true]}.");
        }

        if (op.Props?.TryGetPropertyValue("values", out var valuesNode) == true && valuesNode is JsonArray values)
        {
            WriteRowValues(target.Sheet, rowNumber, values);
        }

        details.Add(new
        {
            op = "add",
            type = "row",
            path = string.Create(
                CultureInfo.InvariantCulture,
                $"{ExcelPaths.SheetPath(target.Sheet)}/row[{rowNumber}]"),
            row = rowNumber,
        });
    }

    /// <summary>
    /// Validates an <c>add chart</c> op against the live workbook and queues
    /// the raw OpenXml write for the post-save pass.
    /// </summary>
    private static void AddChart(XLWorkbook workbook, EditOp op, int index, List<object> details, ChartOpBatch charts)
    {
        var target = ExcelPaths.Resolve(workbook, op.Path);
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add chart targets a sheet path like /Sheet1; the data location goes in props.dataRange.",
                "Use {op:add, type:chart, path:/Sheet1, props:{kind:\"bar\", dataRange:\"A1:B5\", anchor:\"D2\"}}.");
        }

        var spec = ExcelCharts.ParseAdd(target.Sheet, op, index);
        var chartIndex = charts.Add(spec);
        details.Add(new
        {
            op = "add",
            type = "chart",
            path = string.Create(
                CultureInfo.InvariantCulture,
                $"{ExcelPaths.SheetPath(target.Sheet)}/chart[{chartIndex}]"),
            kind = spec.Kind,
            dataRange = spec.DataRange,
            anchor = spec.Anchor,
            series = spec.Series.Count,
        });
    }

    // ----- remove -----------------------------------------------------------

    private static void ApplyRemove(
        XLWorkbook workbook, EditOp op, int index, List<object> details, ChartOpBatch charts, PostSaveWork post)
    {
        // Defined-name paths use an id form the shared grammar cannot parse;
        // peel them off before path resolution (precedent: pivot[@name=…]).
        if (ExcelNames.TryParsePath(op.Path, out var nameSheetPath, out var definedName))
        {
            var (found, scopeSheet) = ExcelNames.Find(workbook, nameSheetPath, definedName);
            var namePath = ExcelNames.PathOf(scopeSheet, found.Name);
            var removedName = found.Name;
            found.Delete();
            details.Add(new { op = "remove", path = namePath, removed = "name", name = removedName });
            return;
        }

        var target = ExcelPaths.Resolve(workbook, op.Path);
        switch (target.Kind)
        {
            case ExcelTargetKind.Pivot:
            {
                var (pivot, _) = ExcelPivots.Find(target);
                var path = ExcelPaths.PivotPath(target.Sheet, pivot.Name);
                target.Sheet.PivotTables.Delete(pivot.Name);
                post.SyncPivotParts = true; // ClosedXML leaves the parts behind; sync after save
                details.Add(new { op = "remove", path, removed = "pivot", name = pivot.Name });
                return;
            }

            case ExcelTargetKind.ConditionalFormat:
            {
                var format = ExcelConditionalFormats.Find(target);
                target.Sheet.ConditionalFormats.Remove(f => ReferenceEquals(f, format));
                details.Add(new
                {
                    op = "remove",
                    path = ExcelPaths.ConditionalFormatPath(target.Sheet, target.ConditionalFormatIndex!.Value),
                    removed = "conditionalFormat",
                    cfKind = ExcelConditionalFormats.KindName(format.ConditionalFormatType),
                });
                return;
            }

            case ExcelTargetKind.Image:
            {
                // Resolve against the batch-projected state, but leave the
                // ClosedXML model alone — the raw post-save pass deletes the
                // anchor (ClosedXML's own deletion corrupts the drawing rels).
                var picture = ExcelImages.FindForEdit(target, post.RemovedImageNames(target.Sheet.Name));
                post.RemovedImages.Add((target.Sheet.Name, picture.Name));
                details.Add(new
                {
                    op = "remove",
                    path = ExcelPaths.ImagePath(target.Sheet, target.ImageIndex!.Value),
                    removed = "image",
                    name = picture.Name,
                });
                return;
            }
        }

        var canonical = DocPath.Parse(op.Path).ToCanonicalString();
        switch (target.Kind)
        {
            case ExcelTargetKind.Sheet:
                if (workbook.Worksheets.Count == 1)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        "A workbook must keep at least one sheet.",
                        "Add a replacement sheet first ({op:add, type:sheet, path:/New}), then remove this one.");
                }

                if (target.Sheet.PivotTables.Any())
                {
                    post.SyncPivotParts = true; // the sheet's pivot caches need pruning after save
                }

                target.Sheet.Delete();
                details.Add(new { op = "remove", path = canonical, removed = "sheet" });
                break;

            case ExcelTargetKind.Row:
                target.Sheet.Row(target.RowNumber!.Value).Delete();
                details.Add(new { op = "remove", path = canonical, removed = "row" });
                break;

            case ExcelTargetKind.Chart:
                charts.Remove(target.Sheet.Name, ExcelPaths.SheetPath(target.Sheet), target.ChartIndex!.Value, index);
                details.Add(new { op = "remove", path = canonical, removed = "chart" });
                break;

            case ExcelTargetKind.Cell:
                target.Cell!.Clear(XLClearOptions.Contents);
                details.Add(new { op = "remove", path = canonical, removed = "contents", note = "styles kept" });
                break;

            default:
                target.Range!.Clear(XLClearOptions.Contents);
                details.Add(new { op = "remove", path = canonical, removed = "contents", note = "styles kept" });
                break;
        }
    }

    // ----- save with cached formula values (the flagship) --------------------

    /// <summary>
    /// Saves the workbook so the file carries cached formula values in
    /// <c>&lt;v&gt;</c> elements (agents and viewers see results without Excel).
    /// Formulas the ClosedXML engine cannot evaluate (they come back
    /// <c>#NAME?</c>) are saved as formula-text-only — their stale error value
    /// is stripped so Excel recomputes them on open — and reported in a
    /// <c>formula_not_evaluated</c> warning. Never silent.
    /// </summary>
    private static List<Warning>? SaveWithCachedValues(XLWorkbook workbook, string file)
    {
        var unevaluated = new List<(string Sheet, string Address)>();
        foreach (var sheet in workbook.Worksheets)
        {
            foreach (var cell in sheet.CellsUsed().Where(c => c.HasFormula))
            {
                XLCellValue value;
                try
                {
                    value = cell.Value; // forces evaluation of dirty formulas
                }
                catch (Exception)
                {
                    unevaluated.Add((sheet.Name, cell.Address.ToString()!));
                    continue;
                }

                if (value.IsError && value.GetError() == XLError.NameNotRecognized)
                {
                    unevaluated.Add((sheet.Name, cell.Address.ToString()!));
                }
            }
        }

        try
        {
            workbook.SaveAs(file, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }
        catch (Exception)
        {
            // Evaluation blew up inside save: fall back to a plain save (no cached
            // values at all) rather than failing the edit. Excel computes on open.
            workbook.SaveAs(file);
            return
            [
                new Warning(
                    ErrorCodes.FormulaNotEvaluated,
                    "Formula evaluation failed during save; the file was saved without cached values. " +
                    "Excel will compute all formulas when the file is opened."),
            ];
        }

        StripStaleCachedValues(file, unevaluated);

        if (unevaluated.Count == 0)
        {
            return null;
        }

        var sample = string.Join(
            ", ",
            unevaluated.Take(10).Select(c => $"/{ExcelPaths.QuoteSheet(c.Sheet)}/{c.Address}"));
        var more = unevaluated.Count > 10 ? $" (+{unevaluated.Count - 10} more)" : string.Empty;
        return
        [
            new Warning(
                ErrorCodes.FormulaNotEvaluated,
                $"{unevaluated.Count} formula cell(s) use functions the built-in engine cannot evaluate: {sample}{more}. " +
                "The formula text is saved without a cached value; Excel computes it when the file opens."),
        ];
    }

    /// <summary>
    /// Removes the cached <c>#NAME?</c> error that ClosedXML wrote for formulas
    /// it cannot evaluate. A formula cell without a cached value forces Excel
    /// to recalculate it on open, which is the honest behavior.
    /// </summary>
    private static void StripStaleCachedValues(string file, List<(string Sheet, string Address)> cells)
    {
        if (cells.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return;
        }

        foreach (var group in cells.GroupBy(c => c.Sheet, StringComparer.OrdinalIgnoreCase))
        {
            var sheetElement = workbookPart.Workbook
                ?.Descendants<Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, group.Key, StringComparison.OrdinalIgnoreCase));
            if (sheetElement?.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart ||
                worksheetPart.Worksheet is not { } worksheetRoot)
            {
                continue;
            }

            var wanted = group.Select(g => g.Address).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var dirty = false;
            foreach (var cell in worksheetRoot.Descendants<Cell>())
            {
                if (cell.CellReference?.Value is { } reference && wanted.Contains(reference))
                {
                    cell.CellValue?.Remove();
                    cell.DataType = null;
                    dirty = true;
                }
            }

            if (dirty)
            {
                worksheetRoot.Save();
            }
        }
    }
}
