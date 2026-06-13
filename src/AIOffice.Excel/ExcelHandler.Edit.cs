using System.Globalization;
using System.Text.Json;
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
        "value", "valueType", "values", "numberFormat", "bold", "italic", "fill", "color", "merge", "name",
        "freezeRows", "freezeCols", "autoFilter", "orientation", "paperSize", "fitToWidth", "printArea",
        "height", "width", "hidden", "hyperlink", "hyperlinkTooltip", "cellStyle",
    ];

    private static readonly IReadOnlyList<string> AddTypes =
    [
        "sheet", "table", "row", "col", "chart", "pivot", "conditionalFormat", "image", "name", "note",
        "dataValidation", "sparkline", "comment", "reply", "group",
    ];

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

        // In-place streaming write (M6 flagship): a large file (or stream=true)
        // whose batch is ENTIRELY streamable cell/range writes is edited without
        // a DOM load. Any other batch (or any op the SAX path can't handle)
        // transparently falls through to the ClosedXML path below — same result.
        if (TryStreamingEdit(ctx, file, ops, sw) is { } streamedEnvelope)
        {
            return streamedEnvelope;
        }

        using var workbook = OpenWorkbook(file);

        // Apply every op in memory first; any failure aborts before a byte is
        // written. Chart ops are validated here too, but their raw OpenXml
        // write is deferred to a post-save pass (ClosedXML cannot author chart
        // parts; measured: it preserves them byte-identical once present).
        var charts = new ChartOpBatch(file);
        var post = new PostSaveWork();
        var opWarnings = new List<Warning>();
        var details = new List<object>(ops.Count);
        for (var i = 0; i < ops.Count; i++)
        {
            // Sequential-op semantics: a queued streamed bulk write must land
            // before any later op that addresses the same sheet (it would
            // otherwise replay after that op and clobber its cells).
            FlushStreamedWritesFor(post, ops[i].Path);
            ApplyOp(ctx, workbook, ops[i], i, details, charts, post, opWarnings);
        }

        if (ArgBool(ctx, "dryRun"))
        {
            return Envelope.Ok(
                new { dryRun = true, wouldApply = details.Count, ops = details },
                MetaFor(file, sw, opWarnings));
        }

        // Safety net: a queued streamed write only stays streamed while its
        // sheet is still bare (e.g. a pivot's targetSheet can land cells on it
        // without addressing it by path). Anything else falls back to the DOM.
        FlushStreamedWritesNoLongerBare(workbook, post);

        var snapshot = _snapshots.Save(file); // pre-image: every successful edit is undoable
        var warnings = new List<Warning>(opWarnings);
        if (SaveWithCachedValues(workbook, file) is { } saveWarnings)
        {
            warnings.AddRange(saveWarnings);
        }

        try
        {
            if (post.StreamedWrites.Count > 0)
            {
                ExcelBulkWrites.ApplyAfterSave(file, post.StreamedWrites);

                // Formulas written by the SAX path (and any formula elsewhere in
                // the workbook that references the streamed sheet) need cached
                // values: run the normal pipeline once over the streamed bytes.
                if (post.StreamedWrites.Any(w => w.HasFormulas) || HasAnyFormula(workbook))
                {
                    using var reopened = OpenWorkbook(file);
                    if (SaveWithCachedValues(reopened, file) is { } formulaWarnings)
                    {
                        warnings.AddRange(formulaWarnings);
                    }
                }
            }

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

            // Threaded comments (M5) live in parts ClosedXML cannot author;
            // write the batch's model back raw.
            if (post.Comments is { Dirty: true } comments)
            {
                ExcelComments.WriteAfterSave(file, comments);
            }

            if (post.MaterializeStyles)
            {
                // Surface the named cell styles in Excel's gallery (real
                // cellStyle entries); the registry property carries the truth.
                using var reopened = OpenWorkbook(file);
                ExcelCellStyles.MaterializeAfterSave(file, ExcelCellStyles.Load(reopened).Values.ToList());
            }

            // Correct ClosedXML's data-bar GUID/orphan defects on the saved
            // bytes (no-op scan when there is nothing to fix).
            ExcelConditionalFormats.FixUpAfterSave(file);

            // Strip the xr2:uid attribute ClosedXML stamps on sparkline groups
            // (Office2019 validator flags it); no-op when there are none.
            ExcelSparklines.FixUpAfterSave(file);
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

    /// <summary>
    /// Attempts the M6 in-place streaming write. Returns a non-null envelope
    /// only when the whole batch was handled by the SAX path; otherwise null,
    /// and the caller runs the normal ClosedXML DOM pipeline (transparent
    /// fallback — the user is never told which path ran, because both produce
    /// the same workbook). dryRun is honored without touching the file.
    /// </summary>
    private Envelope? TryStreamingEdit(
        CommandContext ctx, string file, IReadOnlyList<EditOp> ops, System.Diagnostics.Stopwatch sw)
    {
        // Explicit stream=false forces the DOM path even on a big file (the
        // documented escape hatch; also how the equality gate routes its twin).
        if (ArgIsExplicitlyFalse(ctx, "stream"))
        {
            return null;
        }

        if (!ExcelStreamingWrite.ShouldConsider(file, ArgBool(ctx, "stream")))
        {
            return null;
        }

        var plan = ExcelStreamingWrite.TryPlan(file, ops);
        if (plan is null)
        {
            return null; // not a fully-streamable batch → DOM fallback
        }

        var details = new List<object>(ops.Count);
        foreach (var op in ops)
        {
            details.Add(new
            {
                op = "set",
                path = DocPath.Parse(op.Path).ToCanonicalString(),
                applied = op.Props?.ContainsKey("values") == true ? new List<string> { "values" } : ["value"],
                streamed = true,
            });
        }

        if (ArgBool(ctx, "dryRun"))
        {
            return Envelope.Ok(
                new { dryRun = true, wouldApply = details.Count, ops = details, streamed = true },
                MetaFor(file, sw));
        }

        var snapshot = _snapshots.Save(file); // pre-image: streamed edits are undoable too
        List<Warning>? warnings;
        try
        {
            warnings = ExcelStreamingWrite.Apply(file, plan);
        }
        catch (Exception)
        {
            File.Copy(snapshot.Path, file, overwrite: true); // keep the file intact on any SAX failure
            throw;
        }

        return Envelope.Ok(
            new { applied = details.Count, ops = details, streamed = true },
            MetaFor(file, sw, warnings));
    }

    private static bool HasAnyFormula(XLWorkbook workbook) =>
        workbook.Worksheets.Any(ws => ws.CellsUsed().Any(c => c.HasFormula));

    /// <summary>Raw clean-up passes an edit batch queues for after the ClosedXML save.</summary>
    private sealed class PostSaveWork
    {
        /// <summary>Set when a pivot was removed (its parts outlive ClosedXML's save).</summary>
        public bool SyncPivotParts { get; set; }

        /// <summary>
        /// Set when a named cell style was added/removed: the registry lives in a
        /// custom property ClosedXML persists, but Excel's gallery needs real
        /// <c>cellStyle</c> entries authored raw after the save.
        /// </summary>
        public bool MaterializeStyles { get; set; }

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

        /// <summary>
        /// Bulk 2D writes queued for the SAX streaming path (&gt;50k cells into
        /// a bare sheet). The ClosedXML model is left untouched; the post-save
        /// pass rewrites the sheet's part directly.
        /// </summary>
        public List<ExcelBulkWrites.Pending> StreamedWrites { get; } = [];

        /// <summary>
        /// Threaded-comment state (M5), loaded from disk on the batch's first
        /// comment op and written back raw after the ClosedXML save (ClosedXML
        /// cannot see threadedComments parts).
        /// </summary>
        public ExcelComments.Model? Comments { get; set; }
    }

    /// <summary>The batch's comment model, loaded lazily from the on-disk file.</summary>
    private static ExcelComments.Model CommentModelFor(CommandContext ctx, PostSaveWork post) =>
        post.Comments ??= ExcelComments.Load(ctx.File!);

    /// <summary>
    /// Materializes queued streamed writes whose sheet the next op addresses,
    /// preserving sequential-op semantics (the queued write happened earlier).
    /// </summary>
    private static void FlushStreamedWritesFor(PostSaveWork post, string nextOpPath)
    {
        if (post.StreamedWrites.Count == 0)
        {
            return;
        }

        string? sheetName = null;
        if (ExcelNames.TryParsePath(nextOpPath, out var nameSheetPath, out _))
        {
            nextOpPath = nameSheetPath ?? string.Empty;
        }

        if (nextOpPath.Length > 0 && DocPath.TryParse(nextOpPath, out var path, out _) &&
            path!.Segments[0] is { } first && first.Kind is PathSegmentKind.Name or PathSegmentKind.Element)
        {
            sheetName = first.Name;
        }

        if (sheetName is null)
        {
            return;
        }

        for (var i = post.StreamedWrites.Count - 1; i >= 0; i--)
        {
            var pending = post.StreamedWrites[i];
            if (string.Equals(pending.Sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                WriteGridDom(pending.Sheet, pending.FirstRow, pending.FirstColumn, pending.Grid);
                post.StreamedWrites.RemoveAt(i);
            }
        }
    }

    /// <summary>
    /// Save-time safety net: drops queued writes whose sheet was deleted later
    /// in the batch, and falls back to the DOM for any sheet that is no longer
    /// bare (something else wrote to it without addressing it by path).
    /// </summary>
    private static void FlushStreamedWritesNoLongerBare(XLWorkbook workbook, PostSaveWork post)
    {
        for (var i = post.StreamedWrites.Count - 1; i >= 0; i--)
        {
            var pending = post.StreamedWrites[i];
            if (!workbook.Worksheets.Any(ws => ReferenceEquals(ws, pending.Sheet)))
            {
                post.StreamedWrites.RemoveAt(i); // the sheet (and its write) was removed later in the batch
                continue;
            }

            if (!SheetIsBare(pending.Sheet))
            {
                WriteGridDom(pending.Sheet, pending.FirstRow, pending.FirstColumn, pending.Grid);
                post.StreamedWrites.RemoveAt(i);
            }
        }
    }

    private static void ApplyOp(
        CommandContext ctx, XLWorkbook workbook, EditOp op, int index, List<object> details,
        ChartOpBatch charts, PostSaveWork post, List<Warning> warnings)
    {
        // "/" (document root) is meaningful only as a replace scope (expanded
        // upstream into per-sheet ops). Any other root-targeted op has no xlsx
        // surface — address a sheet/cell/range instead.
        if (op.Op != "replace" && op.Path == "/")
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{index}]: xlsx has no document-root edit target ('/').",
                "Address a sheet (/Sheet1), cell (/Sheet1/A1) or range; document-wide find/replace uses 'replace' with path '/'.");
        }

        // Workbook-level edit targets (document properties, named cell styles)
        // use paths the shared grammar cannot resolve to a sheet, so they are
        // routed before the cell/range dispatch (precedent: defined names).
        if (op.Path == "/properties")
        {
            if (op.Op != "set")
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: /properties only supports 'set'.",
                    "Use {op:set, path:/properties, props:{title:\"…\", custom:{…}}}.");
            }

            var appliedProps = ExcelProperties.ApplySet(workbook, op, index);
            details.Add(new { op = "set", path = "/properties", applied = appliedProps });
            return;
        }

        if (op.Path == "/styles")
        {
            if (op.Op != "add" || !string.Equals(op.Type, "cellStyle", StringComparison.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: /styles only supports adding a cellStyle.",
                    "Use {op:add, type:cellStyle, path:/styles, props:{name:\"Currency-Red\", numberFormat:\"$#,##0.00\"}}.");
            }

            details.Add(ExcelCellStyles.Add(workbook, op, index));
            post.MaterializeStyles = true;
            return;
        }

        if (ExcelCellStyles.TryParsePath(op.Path, out var styleName))
        {
            if (op.Op != "remove")
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: /style[@name=…] only supports 'remove'.",
                    "Define styles with {op:add, type:cellStyle, path:/styles, ...} and apply them with " +
                    "{op:set, path:/Sheet1/B2, props:{cellStyle:\"" + styleName + "\"}}.");
            }

            details.Add(ExcelCellStyles.Remove(workbook, styleName));
            post.MaterializeStyles = true;
            return;
        }

        switch (op.Op)
        {
            case "set":
                ApplySet(workbook, op, index, details, post);
                break;
            case "add":
                ApplyAdd(ctx, workbook, op, index, details, charts, post);
                break;
            case "remove":
                ApplyRemove(ctx, workbook, op, index, details, charts, post);
                break;
            case "replace":
                ApplyReplace(workbook, op, index, details, warnings);
                break;
            case "move":
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"ops[{index}]: move is not supported for xlsx yet.",
                    "Copy the content to the destination with get + set, then remove the source.");
            default: // "accept"/"reject" — ParseBatch already rejected anything else
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"ops[{index}]: {op.Op} is not supported for xlsx.",
                    "accept/reject resolve tracked docx revisions; xlsx has no tracked changes.");
        }
    }

    // ----- set --------------------------------------------------------------

    private static void ApplySet(XLWorkbook workbook, EditOp op, int index, List<object> details, PostSaveWork post)
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
            or ExcelTargetKind.ConditionalFormat or ExcelTargetKind.Image
            or ExcelTargetKind.DataValidation or ExcelTargetKind.Sparkline or ExcelTargetKind.Comment
            or ExcelTargetKind.Table)
        {
            var kindName = target.Kind switch
            {
                ExcelTargetKind.ConditionalFormat => "conditionalFormat",
                ExcelTargetKind.DataValidation => "dataValidation",
                _ => target.Kind.ToString().ToLowerInvariant(),
            };
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
            SetValues(target, valuesNode, index, applied, post);
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

        if (props.TryGetPropertyValue("color", out var colorNode) && colorNode is not null)
        {
            StyleOf(target).Font.FontColor = ParseColor(colorNode.GetValue<string>());
            applied.Add("color");
        }

        if (props.TryGetPropertyValue("cellStyle", out var cellStyleNode) && cellStyleNode is not null)
        {
            // Apply the named style's concrete formatting LAST so explicit props
            // in the same op act as the base and the style wins (the common
            // intent: "make this look like Currency-Red").
            ExcelCellStyles.Apply(workbook, StyleOf(target), cellStyleNode.GetValue<string>(), index);
            applied.Add("cellStyle");
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

        if (props.TryGetPropertyValue("height", out var heightNode) && heightNode is not null)
        {
            ApplyRowHeight(target, op, heightNode, index, applied);
        }

        if (props.TryGetPropertyValue("width", out var widthNode) && widthNode is not null)
        {
            ApplyColumnWidth(target, op, widthNode, index, applied);
        }

        if (props.TryGetPropertyValue("hidden", out var hiddenNode) && hiddenNode is not null)
        {
            ApplyHidden(target, op, hiddenNode, index, applied);
        }

        if (props.ContainsKey("hyperlink") || props.ContainsKey("hyperlinkTooltip"))
        {
            ApplyHyperlink(target, props, index, applied);
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

    /// <summary>
    /// Bulk 2D writes. Two forms (one op writes a whole table):
    /// <list type="bullet">
    /// <item>anchor — <c>{op:set, path:/Sheet1/A2, props:{values:[[…],[…]]}}</c>:
    /// the extent is inferred from the array (rows may be ragged);</item>
    /// <item>range — <c>{op:set, path:/Sheet1/A2:D101, props:{values:[[…]]}}</c>:
    /// the array shape must match the range exactly.</item>
    /// </list>
    /// Values are typed like single-cell set (numbers/booleans/strings/dates/
    /// formula strings). Writes over 50k cells into a bare sheet take the SAX
    /// streaming path (see <see cref="ExcelBulkWrites"/>); everything else goes
    /// through the ClosedXML DOM. Rows keep their flat-array form.
    /// </summary>
    private static void SetValues(ExcelTarget target, JsonNode? valuesNode, int index, List<string> applied, PostSaveWork post)
    {
        if (valuesNode is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: 'values' must be a JSON array.",
                "Rows take a flat array ([1,\"a\",true]); cells (anchor form) and ranges take a 2D array ([[1,2],[3,4]]).");
        }

        switch (target.Kind)
        {
            case ExcelTargetKind.Row:
                WriteRowValues(target.Sheet, target.RowNumber!.Value, array);
                break;

            case ExcelTargetKind.Cell: // anchor form: extent inferred from the array
            {
                var grid = ParseGrid(array, index);
                WriteGrid(
                    target.Sheet, target.Cell!.Address.RowNumber, target.Cell!.Address.ColumnNumber,
                    grid, index, post);
                break;
            }

            case ExcelTargetKind.Range: // range form: the shapes must match exactly
            {
                var address = target.Range!.RangeAddress;
                var rangeRows = address.LastAddress.RowNumber - address.FirstAddress.RowNumber + 1;
                var rangeColumns = address.LastAddress.ColumnNumber - address.FirstAddress.ColumnNumber + 1;
                var grid = ParseGrid(array, index);
                RequireExactShape(grid, rangeRows, rangeColumns, address, index);
                WriteGrid(target.Sheet, address.FirstAddress.RowNumber, address.FirstAddress.ColumnNumber, grid, index, post);
                break;
            }

            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: 'values' targets a row, a cell (anchor form) or a range.",
                    "Anchor a 2D array at a cell ({op:set, path:/Sheet1/A2, props:{values:[[1,2],[3,4]]}}) " +
                    "or address the exact range (/Sheet1/A2:B3).");
        }

        applied.Add("values");
    }

    /// <summary>Parses a 2D JSON array into typed cell values (formula strings included).</summary>
    private static List<List<ExcelValues.ParsedValue>> ParseGrid(JsonArray array, int index)
    {
        if (array.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: 'values' must not be empty.",
                "Pass at least one row, e.g. [[1,2],[3,4]].");
        }

        var grid = new List<List<ExcelValues.ParsedValue>>(array.Count);
        for (var r = 0; r < array.Count; r++)
        {
            if (array[r] is not JsonArray rowArray)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: 'values' must be a 2D array; row {r + 1} is not an array.",
                    "Pass [[1,2],[3,4]] — one inner array per row. (Only row targets take a flat array.)");
            }

            var row = new List<ExcelValues.ParsedValue>(rowArray.Count);
            for (var c = 0; c < rowArray.Count; c++)
            {
                row.Add(ExcelValues.Parse(rowArray[c]));
            }

            grid.Add(row);
        }

        return grid;
    }

    /// <summary>Range form contract: dimension mismatches name BOTH shapes.</summary>
    private static void RequireExactShape(
        List<List<ExcelValues.ParsedValue>> grid, int rangeRows, int rangeColumns, IXLRangeAddress address, int index)
    {
        var gridColumns = grid.Max(r => r.Count);
        if (grid.Count != rangeRows)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: 'values' is {grid.Count} row(s) x {gridColumns} column(s) but the range " +
                $"{address} is {rangeRows} row(s) x {rangeColumns} column(s).",
                "Make the array shape match the range exactly, or anchor at the top-left cell " +
                "(path /Sheet/A2) to infer the extent from the array.");
        }

        for (var r = 0; r < grid.Count; r++)
        {
            if (grid[r].Count != rangeColumns)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: 'values' row {r + 1} has {grid[r].Count} column(s) but the range " +
                    $"{address} is {rangeRows} row(s) x {rangeColumns} column(s).",
                    "Make the array shape match the range exactly, or anchor at the top-left cell " +
                    "(path /Sheet/A2) to infer the extent from the array.");
            }
        }
    }

    /// <summary>
    /// Writes a parsed grid at an anchor: the SAX streaming path when it is
    /// large and the sheet is bare, otherwise cell-by-cell through the DOM.
    /// </summary>
    private static void WriteGrid(
        IXLWorksheet sheet, int firstRow, int firstColumn, List<List<ExcelValues.ParsedValue>> grid,
        int index, PostSaveWork post)
    {
        var lastRow = firstRow + grid.Count - 1;
        var lastColumn = firstColumn + grid.Max(r => r.Count) - 1;
        if (lastRow > MaxSheetRows || lastColumn > MaxSheetColumns)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: 'values' runs past the sheet edge (to row {lastRow}, column {lastColumn}; " +
                $"the sheet ends at row {MaxSheetRows}, column {MaxSheetColumns}).",
                "Move the anchor up/left or shrink the array.");
        }

        if (ExcelBulkWrites.CellCount(grid) > ExcelBulkWrites.StreamingThresholdCells &&
            SheetIsBare(sheet) && ExcelBulkWrites.IsStreamable(grid))
        {
            post.StreamedWrites.Add(new ExcelBulkWrites.Pending(sheet, firstRow, firstColumn, grid));
            return;
        }

        WriteGridDom(sheet, firstRow, firstColumn, grid);
    }

    private static void WriteGridDom(
        IXLWorksheet sheet, int firstRow, int firstColumn, List<List<ExcelValues.ParsedValue>> grid)
    {
        for (var r = 0; r < grid.Count; r++)
        {
            var row = grid[r];
            for (var c = 0; c < row.Count; c++)
            {
                WriteParsed(sheet.Cell(firstRow + r, firstColumn + c), row[c]);
            }
        }
    }

    /// <summary>
    /// True when nothing in the sheet's ClosedXML model would be lost by a raw
    /// sheetData rewrite. (Charts/drawings live in elements the rewrite
    /// preserves verbatim, so they do not gate streaming.)
    /// </summary>
    private static bool SheetIsBare(IXLWorksheet sheet) =>
        sheet.FirstCellUsed(XLCellsUsedOptions.All) is null &&
        !sheet.Tables.Any() &&
        !sheet.PivotTables.Any() &&
        !sheet.ConditionalFormats.Any() &&
        !sheet.Pictures.Any() &&
        !sheet.MergedRanges.Any();

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

    /// <summary>
    /// Cell hyperlinks (M5): <c>{"hyperlink":"https://…"}</c> sets an external
    /// link, <c>{"hyperlink":"#Sheet2!A1"}</c> an internal one (a defined name
    /// works too: <c>#Results</c>), <c>{"hyperlink":""}</c> clears.
    /// <c>hyperlinkTooltip</c> rides along (or retargets an existing link's
    /// tooltip when sent alone).
    /// </summary>
    private static void ApplyHyperlink(ExcelTarget target, JsonObject props, int index, List<string> applied)
    {
        if (target.Kind != ExcelTargetKind.Cell)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: 'hyperlink' targets a single cell like /Sheet1/A1.",
                "Set the link cell-by-cell; ranges are not supported for hyperlinks.");
        }

        var cell = target.Cell!;
        var hasAddress = props.TryGetPropertyValue("hyperlink", out var addressNode);
        var address = addressNode is JsonValue addressValue &&
                      addressValue.GetValueKind() == JsonValueKind.String
            ? addressValue.GetValue<string>()
            : null;
        var tooltip = props.TryGetPropertyValue("hyperlinkTooltip", out var tooltipNode) &&
                      tooltipNode is JsonValue tooltipValue &&
                      tooltipValue.GetValueKind() == JsonValueKind.String
            ? tooltipValue.GetValue<string>()
            : null;

        if (hasAddress && string.IsNullOrEmpty(address))
        {
            // {"hyperlink":""} (or null) clears; clearing a link-less cell is a no-op.
            if (cell.HasHyperlink)
            {
                cell.GetHyperlink().Delete();
            }

            applied.Add("hyperlinkCleared");
            return;
        }

        if (!hasAddress)
        {
            if (!cell.HasHyperlink)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: 'hyperlinkTooltip' needs a hyperlink on the cell.",
                    "Pass both props together: {\"hyperlink\":\"https://…\",\"hyperlinkTooltip\":\"…\"}.");
            }

            cell.GetHyperlink().Tooltip = tooltip;
            applied.Add("hyperlinkTooltip");
            return;
        }

        cell.SetHyperlink(BuildHyperlink(address!, tooltip, index));
        applied.Add("hyperlink");
        if (tooltip is not null)
        {
            applied.Add("hyperlinkTooltip");
        }
    }

    private static XLHyperlink BuildHyperlink(string address, string? tooltip, int index)
    {
        if (address.StartsWith('#'))
        {
            var internalAddress = address[1..];
            if (internalAddress.Length == 0)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: an internal hyperlink needs a target after '#'.",
                    "Use e.g. {\"hyperlink\":\"#Sheet2!A1\"} or a defined name like #Results.");
            }

            return new XLHyperlink(internalAddress, tooltip);
        }

        if (!Uri.TryCreate(address, UriKind.Absolute, out _))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{address}' is neither an absolute URL nor a #internal reference.",
                "Use a full URL ({\"hyperlink\":\"https://example.com\"}, mailto: works too) or " +
                "an internal target ({\"hyperlink\":\"#Sheet2!A1\"}).");
        }

        return new XLHyperlink(address, tooltip);
    }

    private static IXLStyle StyleOf(ExcelTarget target) => target.Kind switch
    {
        ExcelTargetKind.Cell => target.Cell!.Style,
        ExcelTargetKind.Range => target.Range!.Style,
        ExcelTargetKind.Row => target.Sheet.Row(target.RowNumber!.Value).Style,
        ExcelTargetKind.Column => target.Sheet.Column(target.ColumnNumber!.Value).Style,
        _ => throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "Style props (numberFormat, bold, italic, fill) target a cell, range, row or col.",
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
                details.Add(ExcelTables.Add(ExcelPaths.Resolve(workbook, op.Path), op, index));
                break;
            case "row":
                AddRow(workbook, op, index, details);
                break;
            case "col":
                AddColumn(workbook, op, index, details);
                break;
            case "group":
                details.Add(ExcelGroups.Add(workbook, op, index));
                break;
            case "note":
                details.Add(AddNote(ExcelPaths.Resolve(workbook, op.Path), op, index));
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
            case "dataValidation":
                details.Add(ExcelDataValidations.Add(workbook, ExcelPaths.Resolve(workbook, op.Path), op, index));
                break;
            case "sparkline":
                details.Add(ExcelSparklines.Add(workbook, ExcelPaths.Resolve(workbook, op.Path), op, index));
                break;
            case "comment":
                details.Add(AddComment(ctx, ExcelPaths.Resolve(workbook, op.Path), op, index, post));
                break;
            case "reply":
                details.Add(AddReply(ctx, ExcelPaths.Resolve(workbook, op.Path), op, index, post));
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

    private static void AddRow(XLWorkbook workbook, EditOp op, int index, List<object> details)
    {
        var target = ExcelPaths.Resolve(workbook, op.Path);
        var position = InsertPosition(op, index);
        int rowNumber;
        switch (target.Kind)
        {
            case ExcelTargetKind.Sheet: // append after the used range
                rowNumber = (target.Sheet.LastRowUsed()?.RowNumber() ?? 0) + 1;
                break;
            case ExcelTargetKind.Row when position == "after": // insert below, pushing rows down
                rowNumber = target.RowNumber!.Value + 1;
                target.Sheet.Row(target.RowNumber!.Value).InsertRowsBelow(1);
                break;
            case ExcelTargetKind.Row: // default "before": insert at this position, pushing rows down
                rowNumber = target.RowNumber!.Value;
                target.Sheet.Row(rowNumber).InsertRowsAbove(1);
                break;
            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: add row targets a sheet (/Sheet1 appends) or a row (/Sheet1/row[3] inserts).",
                    "Pass values via props, e.g. {\"values\":[1,\"a\",true]}; position is \"before\" (default) or \"after\".");
        }

        if (op.Props?.TryGetPropertyValue("values", out var valuesNode) == true && valuesNode is JsonArray values)
        {
            WriteRowValues(target.Sheet, rowNumber, values);
        }

        details.Add(new
        {
            op = "add",
            type = "row",
            path = ExcelPaths.RowPath(target.Sheet, rowNumber),
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
        CommandContext ctx, XLWorkbook workbook, EditOp op, int index, List<object> details,
        ChartOpBatch charts, PostSaveWork post)
    {
        // Group spans (row[a]:row[b] / col[a]:col[b]) are an element-range form
        // the shared grammar cannot parse; peel them off first (precedent:
        // pivot[@name=…]).
        if (ExcelGroups.TryResolveSpan(workbook, op.Path, out _))
        {
            details.Add(ExcelGroups.Remove(workbook, op, index));
            return;
        }

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

        // remove with props {target:"note"} deletes a cell's note, not its contents.
        if (op.Props?.TryGetPropertyValue("target", out var removeTargetNode) == true && removeTargetNode is not null)
        {
            details.Add(RemoveCellTarget(target, op, removeTargetNode, index));
            return;
        }

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

            case ExcelTargetKind.DataValidation:
            {
                var validation = ExcelDataValidations.Find(target);
                var path = ExcelPaths.DataValidationPath(target.Sheet, target.DataValidationIndex!.Value);
                target.Sheet.DataValidations.Delete(v => ReferenceEquals(v, validation));
                details.Add(new { op = "remove", path, removed = "dataValidation" });
                return;
            }

            case ExcelTargetKind.Sparkline:
                details.Add(ExcelSparklines.Remove(target));
                return;

            case ExcelTargetKind.Comment:
                details.Add(RemoveComment(ctx, target, index, post));
                return;

            case ExcelTargetKind.Table:
                details.Add(ExcelTables.Remove(target));
                return;
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

            case ExcelTargetKind.Column:
                target.Sheet.Column(target.ColumnNumber!.Value).Delete();
                details.Add(new
                {
                    op = "remove",
                    path = ExcelPaths.ColumnPath(target.Sheet, target.ColumnNumber!.Value),
                    removed = "col",
                });
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
