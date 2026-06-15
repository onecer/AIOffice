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
        // v1.2 (additive): per-cell lock + sheet/workbook protection.
        "locked", "protected", "password", "protectStructure", "protectWindows",
        "allowFormatCells", "allowFormatColumns", "allowFormatRows", "allowInsertColumns", "allowInsertRows",
        "allowDeleteColumns", "allowDeleteRows", "allowSort", "allowAutoFilter", "allowPivotTables",
        "allowSelectLockedCells", "allowSelectUnlockedCells",
        // v1.5 (additive): compute actions that persist their result on a cell/sheet.
        "applyScenario", "goalSeek",
    ];

    private static readonly IReadOnlyList<string> AddTypes =
    [
        "sheet", "table", "row", "col", "chart", "pivot", "conditionalFormat", "image", "name", "note",
        "dataValidation", "sparkline", "comment", "reply", "group", "slicer", "embed",
        // v1.2 (additive): interactive form controls.
        "formControl",
        // v1.4 (additive): what-if data tables.
        "dataTable",
        // v1.5 (additive): scenario manager.
        "scenario",
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
        var slicers = new ExcelSlicers.Batch(file);
        var embeds = new ExcelEmbeds.Batch(file);
        var formControls = new ExcelFormControls.Batch(file);
        var post = new PostSaveWork();

        // aboveBelowAverage conditional rules (v1.3) are authored raw: ClosedXML
        // 0.105 reads them but throws when re-serializing. If the file already
        // carries any, capture their specs and strip them from the ClosedXML
        // model so the save succeeds, then re-author them in the post-save pass.
        if (ExcelConditionalFormats.FileHasAverageRules(file))
        {
            post.AverageRules.AddRange(ExcelConditionalFormats.CaptureAverageRules(file));
            ExcelConditionalFormats.StripAverageRulesFromModel(workbook);
        }
        var opWarnings = new List<Warning>();
        var details = new List<object>(ops.Count);
        for (var i = 0; i < ops.Count; i++)
        {
            // Sequential-op semantics: a queued streamed bulk write must land
            // before any later op that addresses the same sheet (it would
            // otherwise replay after that op and clobber its cells).
            FlushStreamedWritesFor(post, ops[i].Path);
            ApplyOp(ctx, workbook, ops[i], i, details, charts, slicers, embeds, formControls, post, opWarnings);
        }

        if (ArgBool(ctx, "dryRun"))
        {
            return Envelope.Ok(
                new { dryRun = true, wouldApply = details.Count, ops = details },
                MetaFor(file, sw, opWarnings));
        }

        // A batch of nothing but extract ops produces dest files but never
        // touches the source: skip the save (and the snapshot) so the source
        // stays byte-identical, honoring "extract does not modify the document".
        if (ops.Count > 0 && ops.All(o => o.Op == "extract"))
        {
            return Envelope.Ok(
                new { applied = details.Count, ops = details },
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

            // (1.4) Dynamic-array spills and what-if data tables are authored raw
            // on the saved bytes: ClosedXML cannot model spill metadata or the
            // {=TABLE(...)} construct, and the cells were left untouched above so
            // the cached values land cleanly. Runs before charts (a chart can
            // reference a spilled range).
            if (post.Spills.Count > 0)
            {
                ExcelDynamicArrays.ApplyAfterSave(file, post.Spills);
            }

            if (post.DataTables.Count > 0)
            {
                ExcelDataTables.ApplyAfterSave(file, post.DataTables);
            }

            if (post.RemovedDataTables.Count > 0)
            {
                ExcelDataTables.RemoveAfterSave(file, post.RemovedDataTables);
            }

            // (1.5) Scenario manager: scenarios are authored/removed raw in the
            // worksheet's scenarios part (ClosedXML has no model). Applying a
            // scenario writes its values into the live model during op application,
            // so the normal save recomputes dependents — only the saved scenario
            // definitions ride this raw pass.
            if (post.Scenarios.Count > 0)
            {
                ExcelScenarios.ApplyAfterSave(file, post.Scenarios);
            }

            if (post.RemovedScenarios.Count > 0)
            {
                ExcelScenarios.RemoveAfterSave(file, post.RemovedScenarios);
            }

            if (!charts.IsEmpty)
            {
                ExcelCharts.Apply(file, charts.Ops);
            }

            if (!slicers.IsEmpty)
            {
                // Slicers ride the same post-save discipline as charts: ClosedXML
                // cannot author them and preserves them byte-identical once present.
                ExcelSlicers.Apply(file, slicers.Ops);
            }

            if (!formControls.IsEmpty)
            {
                // Form controls ride the same post-save discipline as charts and
                // slicers: ClosedXML cannot author the legacy VML/ctrlProp parts
                // and preserves them byte-identical once present (measured).
                ExcelFormControls.Apply(file, formControls.Ops);
            }

            if (post.HasProtection)
            {
                // Sheet/workbook protection is authored raw post-save: ClosedXML
                // cannot lift a password-protected sheet without the password,
                // and this is documented light protection aioffice always owns.
                ExcelProtection.Apply(file, post.SheetProtections, post.WorkbookProtection);
            }

            if (!embeds.IsEmpty)
            {
                // Embedded package parts and their registry property are authored
                // raw too; ClosedXML preserves both byte-identical once present.
                ExcelEmbeds.Apply(file, embeds.Ops);
            }

            if (post.SyncPivotParts)
            {
                // ClosedXML's save leaves removed pivots' parts behind; sync
                // the package to the in-memory model (the source of truth).
                ExcelPivots.SyncPartsAfterSave(file, ExcelPivots.AliveNames(workbook));
            }

            if (post.CalculatedFields.Count > 0)
            {
                // Author pivot calculated fields raw (v1.3): the cacheField formula,
                // its pivotField/dataField and per-record placeholders ClosedXML
                // cannot model. Runs after the pivot-part sync so the target parts
                // are settled. ClosedXML reopens such a file thanks to the
                // placeholder values (documented round-trip exception).
                ExcelPivots.ApplyCalculatedFields(file, post.CalculatedFields);
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

            // Author the aboveBelowAverage conditional-format rules (v1.3): the
            // rule type ClosedXML cannot save is written raw with its differential
            // style. Runs before the data-bar fix-up so the saved bytes are
            // already complete when that scan reads them.
            if (post.AverageRules.Count > 0)
            {
                ExcelConditionalFormats.ApplyAverageRules(file, post.AverageRules);
            }

            // Apply the chart-polish edits (v1.3): a set on a chart path adjusts
            // existing chart XML ClosedXML cannot see. Runs after chart adds so a
            // batch can add a chart and immediately polish it.
            if (post.ChartPolishEdits.Count > 0)
            {
                ExcelChartPolish.ApplyEdits(file, post.ChartPolishEdits);
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
            new { applied = details.Count, snapshot = snapshot.Number, ops = details },
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
            new { applied = details.Count, snapshot = snapshot.Number, ops = details, streamed = true },
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

        /// <summary>
        /// Sheet-protection changes queued for the post-save raw pass (v1.2).
        /// ClosedXML cannot unprotect a password-protected sheet without the
        /// password, so the sheetProtection bytes are owned directly.
        /// </summary>
        public List<SheetProtectionSpec> SheetProtections { get; } = [];

        /// <summary>Workbook-structure-protection change queued for the post-save raw pass (v1.2).</summary>
        public WorkbookProtectionSpec? WorkbookProtection { get; set; }

        /// <summary>True when any protection change must be written after the save.</summary>
        public bool HasProtection => SheetProtections.Count > 0 || WorkbookProtection is not null;

        /// <summary>
        /// aboveBelowAverage conditional-format rules queued for the post-save raw
        /// pass (v1.3). ClosedXML 0.105 throws on the <c>aboveAverage</c> rule
        /// type, so these rules (plus their differential fill) are authored
        /// directly on the saved bytes.
        /// </summary>
        public List<ExcelConditionalFormats.AverageRuleSpec> AverageRules { get; } = [];

        /// <summary>
        /// Chart-polish edits queued for the post-save raw pass (v1.3). A <c>set</c>
        /// on a <c>/Sheet/chart[i]</c> path adjusts existing chart XML
        /// (data labels, legend, axis titles, trendlines, error bars, gridlines,
        /// secondary axis); ClosedXML cannot see charts, so the writes ride the
        /// same discipline as chart adds.
        /// </summary>
        public List<ExcelChartPolish.EditSpec> ChartPolishEdits { get; } = [];

        /// <summary>
        /// Pivot calculated fields queued for the post-save raw pass (v1.3).
        /// ClosedXML 0.105 has no calculated-field model, so the cacheField
        /// (formula + databaseField=0), the matching pivotField/dataField and the
        /// per-record placeholder values are authored directly on the saved bytes.
        /// </summary>
        public List<ExcelPivots.CalculatedFieldSpec> CalculatedFields { get; } = [];

        /// <summary>
        /// (1.4) Dynamic-array spills queued for the post-save raw pass: the anchor
        /// CSE-array formula plus the computed cached values across the spill
        /// rectangle. ClosedXML 0.105 has no spill model, so these are authored
        /// directly on the saved bytes; the ClosedXML model leaves the spill cells
        /// untouched (the anchor formula is never written through ClosedXML — it
        /// would evaluate to #NAME?).
        /// </summary>
        public List<ExcelDynamicArrays.Pending> Spills { get; } = [];

        /// <summary>
        /// (1.4) What-if data tables queued for the post-save raw pass: the
        /// <c>{=TABLE(rowInput,colInput)}</c> array construct plus the computed
        /// cached results across the body. ClosedXML 0.105 has no data-table model.
        /// </summary>
        public List<ExcelDataTables.Pending> DataTables { get; } = [];

        /// <summary>(1.4) Data tables queued for raw removal post-save (sheet name + 1-based index).</summary>
        public List<(string Sheet, int Index)> RemovedDataTables { get; } = [];

        /// <summary>
        /// (1.5) Scenarios queued for the post-save raw pass: the changing cells and
        /// the values they take, authored into the worksheet's scenarios part.
        /// ClosedXML 0.105 has no scenario model.
        /// </summary>
        public List<ExcelScenarios.Pending> Scenarios { get; } = [];

        /// <summary>(1.5) Scenarios queued for raw removal post-save (sheet name + scenario name).</summary>
        public List<(string Sheet, string Name)> RemovedScenarios { get; } = [];

        /// <summary>True when any write-time formula evaluation must be authored after the save.</summary>
        public bool HasEvaluatedFormulas => Spills.Count > 0 || DataTables.Count > 0;

        /// <summary>
        /// (1.4) Cells already claimed by a queued spill/data-table in this batch,
        /// so a later op (and the next spill's clearance check) sees them as taken
        /// even though the ClosedXML model is empty there.
        /// </summary>
        public bool Claims(string sheetName, int row, int column)
        {
            foreach (var s in Spills)
            {
                if (string.Equals(s.Sheet, sheetName, StringComparison.OrdinalIgnoreCase) &&
                    row >= s.FirstRow && row < s.FirstRow + s.Grid.Count &&
                    column >= s.FirstColumn && column < s.FirstColumn + (s.Grid.Count > 0 ? s.Grid[0].Count : 0))
                {
                    return true;
                }
            }

            return false;
        }
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
        ChartOpBatch charts, ExcelSlicers.Batch slicers, ExcelEmbeds.Batch embeds,
        ExcelFormControls.Batch formControls, PostSaveWork post, List<Warning> warnings)
    {
        // "/" (document root) is a replace scope (expanded upstream into
        // per-sheet ops) AND, since v1.2, the workbook-structure protection
        // target ({op:set, path:/, props:{protectStructure:true}}). Any other
        // root-targeted op has no xlsx surface.
        if (op.Op == "set" && op.Path == "/")
        {
            ApplyWorkbookProtection(op, index, details, post);
            return;
        }

        if (op.Op != "replace" && op.Path == "/")
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{index}]: xlsx has no document-root edit target ('/').",
                "Address a sheet (/Sheet1), cell (/Sheet1/A1) or range; document-wide find/replace uses 'replace' with path '/'. " +
                "Workbook structure protection uses {op:set, path:/, props:{protectStructure:true}}.");
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
                ApplySet(ctx, workbook, op, index, details, post, warnings);
                break;
            case "add":
                ApplyAdd(ctx, workbook, op, index, details, charts, slicers, embeds, formControls, post);
                break;
            case "remove":
                ApplyRemove(ctx, workbook, op, index, details, charts, slicers, embeds, formControls, post);
                break;
            case "replace":
                ApplyReplace(workbook, op, index, details, warnings);
                break;
            case "extract":
                ApplyExtract(ctx, workbook, op, index, details);
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

    // ----- extract ----------------------------------------------------------

    /// <summary>
    /// The M10 <c>extract</c> op: writes an embedded object's payload to a
    /// sandbox-resolved destination. It is a producing op — it reads the embed
    /// from the on-disk file and writes the dest, never mutating the source
    /// document. Only embed paths are extractable today.
    /// </summary>
    private static void ApplyExtract(
        CommandContext ctx, XLWorkbook workbook, EditOp op, int index, List<object> details)
    {
        var target = ExcelPaths.Resolve(workbook, op.Path);
        if (target.Kind != ExcelTargetKind.Embed)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: extract targets an embedded object like /Sheet1/embed[1].",
                "Only embeds can be extracted; run 'aioffice read --view embeds' to list them.");
        }

        var to = op.Props is not null && op.Props.TryGetPropertyValue("to", out var toNode) &&
                 toNode is JsonValue toValue && toValue.GetValueKind() == JsonValueKind.String
            ? toValue.GetValue<string>()
            : null;
        if (string.IsNullOrWhiteSpace(to))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: extract needs props.to (the destination file inside the workspace).",
                "Use {op:extract, path:/Sheet1/embed[1], props:{to:\"out/report.xlsx\"}}.");
        }

        // Sandbox the dest BEFORE writing a byte (mustExist:false — it is created).
        var dest = ctx.Workspace.Resolve(to, mustExist: false);
        var canonical = ExcelEmbeds.Extract(ctx.File!, target, dest);
        details.Add(new { op = "extract", path = canonical, to, extracted = "embed" });
    }

    // ----- set --------------------------------------------------------------

    private static void ApplySet(
        CommandContext ctx, XLWorkbook workbook, EditOp op, int index, List<object> details, PostSaveWork post,
        List<Warning> warnings)
    {
        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: set needs props.",
                "Pass props like {\"value\":42}, {\"value\":\"=SUM(A1:A2)\"} or {\"bold\":true}.");
        }

        // A set on a chart path with chart-polish props is the v1.3 chart-polish
        // edit; it owns its own prop vocabulary, so route it before the cell-set
        // prop guard (which would reject dataLabels/legend/etc.).
        if (ExcelChartPolish.HasPolishProps(props) && TrySetChartPolish(ctx, workbook, op, index, details, post))
        {
            return;
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
            or ExcelTargetKind.Table or ExcelTargetKind.Embed or ExcelTargetKind.FormControl)
        {
            var kindName = target.Kind switch
            {
                ExcelTargetKind.ConditionalFormat => "conditionalFormat",
                ExcelTargetKind.DataValidation => "dataValidation",
                ExcelTargetKind.FormControl => "formControl",
                _ => target.Kind.ToString().ToLowerInvariant(),
            };
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{index}]: set on a {kindName} is not supported yet.",
                $"Remove the {kindName} and add it again with the new props: " +
                "{op:remove, path:" + op.Path + "} then {op:add, type:" + kindName + ", ...}.");
        }

        var applied = new List<string>();

        // v1.2 protection props, routed before the cell-style props. 'locked'
        // targets a cell/range/row/col; the sheet-protection props target a sheet.
        if (props.TryGetPropertyValue("locked", out var lockedNode) && lockedNode is not null)
        {
            ExcelProtection.ApplyLocked(target, BoolPropValue(lockedNode, "locked", index), index, applied);
        }

        if (props.Any(p => ExcelProtection.SheetProps.Contains(p.Key, StringComparer.Ordinal)))
        {
            ApplySheetProtection(target, props, index, applied, post);
        }

        // (1.5) applyScenario writes a saved scenario's values into its changing
        // cells (sheet target); goalSeek solves for a cell's value (cell target).
        // Both are compute actions that persist their result — additive behavior on
        // the frozen 'set' verb, no new verb.
        if (props.TryGetPropertyValue("applyScenario", out var applyNode) && applyNode is not null)
        {
            ApplyScenarioSet(ctx, workbook, target, applyNode, index, applied, post);
        }

        // goalSeek is a dedicated compute action that owns the whole op (it SETs the
        // changing cell to the solved value); it adds its own details entry and
        // returns, so it never combines with other set props.
        if (props.TryGetPropertyValue("goalSeek", out var goalSeekNode) && goalSeekNode is not null)
        {
            ApplyGoalSeek(workbook, target, goalSeekNode, index, details, warnings);
            return;
        }

        if (props.TryGetPropertyValue("value", out var valueNode))
        {
            SetValue(target, valueNode, ValueTypeOf(props), applied, index, post);
        }

        if (props.TryGetPropertyValue("values", out var valuesNode))
        {
            SetValues(target, valuesNode, index, applied, post);
        }

        if (props.TryGetPropertyValue("numberFormat", out var formatNode) && formatNode is not null)
        {
            // A named preset (v1.2, additive) resolves to its OOXML format code;
            // any other string is a literal custom format and passes through
            // unchanged (the v1.0 behavior — never broken).
            StyleOf(target).NumberFormat.Format = ExcelNumberFormats.Resolve(formatNode.GetValue<string>());
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

    /// <summary>
    /// Handles a v1.3 chart-polish set on a <c>/Sheet/chart[i]</c> path: validates
    /// the chart exists, every prop is a polish prop, and any secondaryAxis names
    /// match the chart's series; then queues the polish for the post-save raw
    /// pass. Returns false when the path is not a chart (the caller falls back to
    /// the normal cell-set flow, which reports the polish keys as unknown props).
    /// </summary>
    private static bool TrySetChartPolish(
        CommandContext ctx, XLWorkbook workbook, EditOp op, int index, List<object> details, PostSaveWork post)
    {
        // Defined-name and cell-style paths resolve through their own grammar;
        // they are never chart paths, so leave them to the normal flow.
        if (ExcelNames.TryParsePath(op.Path, out _, out _) || ExcelCellStyles.TryParsePath(op.Path, out _))
        {
            return false;
        }

        var target = ExcelPaths.Resolve(workbook, op.Path);
        if (target.Kind != ExcelTargetKind.Chart)
        {
            return false;
        }

        var props = op.Props!;
        foreach (var (key, _) in props)
        {
            if (!ExcelChartPolish.Props.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: '{key}' is not a chart-polish prop.",
                    "A set on a chart adjusts only its polish: " + string.Join(", ", ExcelChartPolish.Props) +
                    ". Remove and re-add the chart to change its data or kind.",
                    candidates: ExcelChartPolish.Props);
            }
        }

        var settings = ExcelChartPolish.Parse(props, index);

        // Resolve the chart's series names from the raw package so secondaryAxis
        // names can be validated (ClosedXML cannot see charts). This also confirms
        // the chart[index] exists on the sheet.
        var chartIndex = target.ChartIndex!.Value;
        IReadOnlyList<string> seriesNames;
        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(ctx.File!, isEditable: false))
        {
            var chartPart = ExcelChartPolish.ChartPartFor(document, target.Sheet.Name, chartIndex);
            if (chartPart?.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>() is not { } chart)
            {
                var charts = ExcelCharts.Read(document)
                    .Where(c => string.Equals(c.SheetName, target.Sheet.Name, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"No chart[{chartIndex}] on sheet '{target.Sheet.Name}' ({charts.Count} chart(s) exist).",
                    charts.Count > 0
                        ? "Chart indices are 1-based per sheet; pick one of the candidates."
                        : "This sheet has no charts; add one with {op:add, type:chart, path:" +
                          ExcelPaths.SheetPath(target.Sheet) + ", props:{kind:\"bar\", dataRange:\"A1:B5\", anchor:\"D2\"}}.",
                    candidates: charts.Count > 0
                        ? [.. charts.Select(c => c.Path)]
                        : [ExcelPaths.SheetPath(target.Sheet)]);
            }

            seriesNames = ExcelChartPolish.SeriesNames(chart);
        }

        if (settings.SecondaryAxisSeries is { Count: > 0 } secondaryNames)
        {
            foreach (var name in secondaryNames)
            {
                if (!seriesNames.Contains(name, StringComparer.OrdinalIgnoreCase))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{index}]: secondaryAxis names series '{name}', which chart[{chartIndex}] does not have.",
                        "secondaryAxis lists series names from the chart; pick one of the candidates.",
                        candidates: [.. seriesNames]);
                }
            }

            if (secondaryNames.Count >= seriesNames.Count)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: secondaryAxis would move every series off the primary axis.",
                    "Leave at least one series on the primary value axis.");
            }
        }

        post.ChartPolishEdits.Add(new ExcelChartPolish.EditSpec(target.Sheet.Name, chartIndex, settings));
        details.Add(new
        {
            op = "set",
            path = string.Create(
                CultureInfo.InvariantCulture,
                $"{ExcelPaths.SheetPath(target.Sheet)}/chart[{chartIndex}]"),
            applied = props.Select(p => p.Key).ToList(),
        });
        return true;
    }

    // ----- protection (v1.2) -----------------------------------------------------

    /// <summary>Routes the sheet-protection props ({protected, password?, allow*?}) to the protection layer.</summary>
    private static void ApplySheetProtection(
        ExcelTarget target, JsonObject props, int index, List<string> applied, PostSaveWork post)
    {
        bool? protect = props.TryGetPropertyValue("protected", out var protectedNode) && protectedNode is not null
            ? BoolPropValue(protectedNode, "protected", index)
            : null;
        var password = props.TryGetPropertyValue("password", out var passwordNode) &&
                       passwordNode is JsonValue passwordValue &&
                       passwordValue.GetValueKind() == JsonValueKind.String
            ? passwordValue.GetValue<string>()
            : null;

        if (protect is null)
        {
            // allow*/password without 'protected' is ambiguous — name the fix.
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: sheet-protection props need 'protected' to say whether to protect or unprotect.",
                "Pass {protected:true} (with optional password / allow* flags) or {protected:false} to lift protection.");
        }

        var flags = new Dictionary<string, bool>(StringComparer.Ordinal);
        foreach (var prop in ExcelProtection.SheetProps)
        {
            if (prop is "protected" or "password")
            {
                continue;
            }

            if (props.TryGetPropertyValue(prop, out var node) && node is not null)
            {
                flags[prop] = BoolPropValue(node, prop, index);
            }
        }

        // Queue the raw post-save write (one entry per sheet; a later op on the
        // same sheet replaces the earlier queued state).
        post.SheetProtections.RemoveAll(s =>
            string.Equals(s.SheetName, target.Sheet.Name, StringComparison.OrdinalIgnoreCase));
        post.SheetProtections.Add(ExcelProtection.BuildSheetSpec(target, protect.Value, password, flags, index));
        applied.Add(protect.Value ? "protected" : "unprotected");
    }

    /// <summary>
    /// (1.5) Handles {op:set, path:/Sheet1, props:{applyScenario:"Best Case"}} —
    /// queues writing the named scenario's stored values into its changing cells
    /// (and a recalc) for the post-save raw pass. The scenario must already exist
    /// on the saved file (added in an earlier op of this batch, or previously).
    /// </summary>
    private static void ApplyScenarioSet(
        CommandContext ctx, XLWorkbook workbook, ExcelTarget target, JsonNode applyNode, int index,
        List<string> applied, PostSaveWork post)
    {
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: applyScenario targets a sheet path like /Sheet1.",
                "Use {op:set, path:/Sheet1, props:{applyScenario:\"Best Case\"}} to apply a saved scenario.");
        }

        if (applyNode is not JsonValue value || value.GetValueKind() != JsonValueKind.String)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: applyScenario takes the scenario name as a string.",
                "Use {op:set, path:/Sheet1, props:{applyScenario:\"Best Case\"}}.");
        }

        var name = value.GetValue<string>();

        // Resolve the scenario's changing cells + values: a scenario added earlier in
        // THIS batch (not yet on disk) is read from the pending queue; otherwise from
        // the saved file. Apply the values into the LIVE ClosedXML model now, so the
        // normal save recomputes every dependent formula's cached value (headless
        // readers see the scenario's effect immediately — not just a recalc flag).
        var pending = post.Scenarios
            .Where(s => string.Equals(s.Sheet, target.Sheet.Name, StringComparison.OrdinalIgnoreCase))
            .LastOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));

        IReadOnlyList<(string Cell, XLCellValue Value)> cells;
        if (pending is not null)
        {
            cells = pending.Cells;
        }
        else
        {
            var info = ExcelScenarios.ReadOnSheet(ctx.File!, target.Sheet.Name)
                .FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
                ?? throw NoSuchScenario(ctx.File!, target.Sheet, name, index);
            cells = [.. info.Cells.Select(c => (c.Cell, ExcelScenarios.TypeStoredValue(c.Val)))];
        }

        foreach (var (cellRef, cellValue) in cells)
        {
            target.Sheet.Cell(cellRef).Value = cellValue;
        }

        // Recalculate so every formula that depends on a changing cell carries a
        // fresh cached value through the normal save (setting .Value alone does not
        // mark transitive dependents dirty in ClosedXML 0.105).
        workbook.RecalculateAllFormulas();
        applied.Add("applyScenario");
    }

    /// <summary>The invalid-args exception for an applyScenario that names a scenario that does not exist.</summary>
    private static AiofficeException NoSuchScenario(string file, IXLWorksheet sheet, string name, int index)
    {
        var existing = ExcelScenarios.ReadOnSheet(file, sheet.Name);
        return new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: no scenario named '{name}' on sheet '{sheet.Name}' ({existing.Count} scenario(s) exist).",
            existing.Count > 0
                ? "Pick one of the candidates, or add it first with {op:add, type:scenario, …}."
                : "Add a scenario first with {op:add, type:scenario, path:" + ExcelPaths.SheetPath(sheet) +
                  ", props:{name:\"" + name + "\", cells:{…}}}.",
            candidates: existing.Count > 0
                ? [.. existing.Select(s => ExcelScenarios.ScenarioPath(sheet, s.Name))]
                : [ExcelPaths.SheetPath(sheet)]);
    }

    /// <summary>
    /// (1.5) Handles {op:set, path:/Sheet1/B1, props:{goalSeek:{targetCell:"B5",
    /// targetValue:1000}}} — solves for the value of the changing cell B1 that makes
    /// the formula cell B5 equal targetValue (Newton + bisection), SETs B1 to the
    /// found value and recalculates. No convergence leaves B1 unchanged and adds a
    /// goal_seek_no_solution warning (soft outcome, never a hard error).
    /// </summary>
    private static void ApplyGoalSeek(
        XLWorkbook workbook, ExcelTarget target, JsonNode goalSeekNode, int index,
        List<object> details, List<Warning> warnings)
    {
        if (target.Kind != ExcelTargetKind.Cell)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: goalSeek targets the changing cell (a single cell) like /Sheet1/B1.",
                "Use {op:set, path:/Sheet1/B1, props:{goalSeek:{targetCell:\"B5\", targetValue:1000}}}.");
        }

        if (goalSeekNode is not JsonObject spec)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: goalSeek takes an object {{targetCell, targetValue}}.",
                "Use {op:set, path:/Sheet1/B1, props:{goalSeek:{targetCell:\"B5\", targetValue:1000}}}.");
        }

        var targetCellRef = spec.TryGetPropertyValue("targetCell", out var tcNode) && tcNode is JsonValue tcv &&
                            tcv.GetValueKind() == JsonValueKind.String
            ? tcv.GetValue<string>()
            : throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: goalSeek needs a targetCell (the formula cell to drive).",
                "Use goalSeek:{targetCell:\"B5\", targetValue:1000}.");

        if (!(spec.TryGetPropertyValue("targetValue", out var tvNode) && tvNode is JsonValue tvv &&
              tvv.GetValueKind() == JsonValueKind.Number))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: goalSeek needs a numeric targetValue.",
                "Use goalSeek:{targetCell:\"B5\", targetValue:1000}.");
        }

        var targetValue = tvv.GetValue<double>();

        IXLCell targetCell;
        try
        {
            targetCell = target.Sheet.Cell(targetCellRef);
        }
        catch (Exception)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: goalSeek targetCell '{targetCellRef}' is not a valid single-cell address.",
                "Pass one cell that holds a formula depending on the changing cell, e.g. targetCell:\"B5\".");
        }

        if (!targetCell.HasFormula)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: goalSeek targetCell {targetCell.Address} must hold a formula that depends on " +
                $"the changing cell {target.Cell!.Address}.",
                "Set the target cell to a formula (e.g. =B1*2) before goal-seeking the input.");
        }

        var result = ExcelGoalSeek.Solve(workbook, target.Cell!, targetCell, targetValue);
        var canonical = ExcelPaths.CellPath(target.Sheet, target.Cell!.Address);
        if (!result.Converged)
        {
            // Leave the changing cell at its original value (Solve restored the
            // model). A soft warning, not a hard error.
            warnings.Add(new Warning(
                WarningCodes.GoalSeekNoSolution,
                $"Goal seek on {canonical} could not find an input that makes {targetCell.Address} equal " +
                $"{ExcelGoalSeek.Format(targetValue)}; {target.Cell!.Address} was left unchanged."));
            details.Add(new
            {
                op = "set",
                path = canonical,
                applied = new List<string> { "goalSeekNoSolution" },
                goalSeek = new { targetCell = targetCell.Address.ToString(), targetValue, converged = false },
            });
            return;
        }

        // Persist the found input and recalc so the target cell reflects it.
        target.Cell!.Value = result.Input;
        workbook.RecalculateAllFormulas();
        var achieved = targetCell.TryGetValue<double>(out var a) ? a : double.NaN;

        details.Add(new
        {
            op = "set",
            path = canonical,
            applied = new List<string> { "goalSeek" },
            goalSeek = new
            {
                targetCell = targetCell.Address.ToString(),
                targetValue,
                converged = true,
                input = result.Input,
                achievedTarget = achieved,
            },
        });
    }

    /// <summary>Handles {op:set, path:/, props:{protectStructure, protectWindows?, password?}}.</summary>
    private static void ApplyWorkbookProtection(EditOp op, int index, List<object> details, PostSaveWork post)
    {
        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: workbook-root set needs props.",
                "Use {op:set, path:/, props:{protectStructure:true, password:\"secret\"}}.");
        }

        foreach (var (key, _) in props)
        {
            if (!ExcelProtection.WorkbookProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: unknown workbook-root prop '{key}'.",
                    "The workbook root ('/') supports only structure protection: " +
                    string.Join(", ", ExcelProtection.WorkbookProps) + ".",
                    candidates: ExcelProtection.WorkbookProps);
            }
        }

        bool? structure = props.TryGetPropertyValue("protectStructure", out var sNode) && sNode is not null
            ? BoolPropValue(sNode, "protectStructure", index)
            : null;
        bool? windows = props.TryGetPropertyValue("protectWindows", out var wNode) && wNode is not null
            ? BoolPropValue(wNode, "protectWindows", index)
            : null;
        var password = props.TryGetPropertyValue("password", out var pwNode) &&
                       pwNode is JsonValue pwValue && pwValue.GetValueKind() == JsonValueKind.String
            ? pwValue.GetValue<string>()
            : null;

        if (structure is null && windows is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: workbook protection needs protectStructure (and/or protectWindows).",
                "Use {op:set, path:/, props:{protectStructure:true}} to lock the sheet structure.");
        }

        post.WorkbookProtection = ExcelProtection.BuildWorkbookSpec(structure, windows, password, index);
        var applied = post.WorkbookProtection.Clear
            ? new List<string> { "structureUnprotected" }
            : [structure == true ? "structureProtected" : "windowsProtected"];
        details.Add(new { op = "set", path = "/", applied });
    }

    private static string? ValueTypeOf(JsonObject props) =>
        props.TryGetPropertyValue("valueType", out var node) && node is not null
            ? node.GetValue<string>()
            : null;

    private static void SetValue(
        ExcelTarget target, JsonNode? valueNode, string? valueType, List<string> applied, int index, PostSaveWork post)
    {
        var parsed = ExcelValues.Parse(valueNode, valueType);

        // (1.4) A dynamic-array function on a single cell is EVALUATED and spilled
        // at write time (FILTER/UNIQUE/SORT/…): queue the computed result for the
        // post-save raw pass and leave the ClosedXML model untouched (writing the
        // formula through ClosedXML would only evaluate to #NAME?).
        if (parsed.IsFormula && target.Kind == ExcelTargetKind.Cell &&
            ExcelDynamicArrays.IsDynamicArrayFormula(parsed.Formula!))
        {
            var spill = ExcelDynamicArrays.Evaluate(
                target.Sheet, target.Cell!, parsed.Formula!, index,
                (r, c) => post.Claims(target.Sheet.Name, r, c));
            post.Spills.Add(spill);
            applied.Add("spill");
            return;
        }

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
        ChartOpBatch charts, ExcelSlicers.Batch slicers, ExcelEmbeds.Batch embeds,
        ExcelFormControls.Batch formControls, PostSaveWork post)
    {
        switch (op.Type)
        {
            case "sheet":
                AddSheet(workbook, op, index, details);
                break;
            case "formControl":
                AddFormControl(workbook, op, index, details, formControls);
                break;
            case "table":
                details.Add(ExcelTables.Add(ExcelPaths.Resolve(workbook, op.Path), op, index));
                break;
            case "slicer":
                AddSlicer(workbook, op, index, details, slicers);
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
                AddPivot(workbook, op, index, details, post);
                break;
            case "conditionalFormat":
                details.Add(ExcelConditionalFormats.Add(ExcelPaths.Resolve(workbook, op.Path), op, index, post.AverageRules));
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
            case "embed":
                AddEmbed(ctx, workbook, op, index, details, embeds);
                break;
            case "dataTable":
                details.Add(ExcelDataTables.Add(workbook, ExcelPaths.Resolve(workbook, op.Path), op, index, post.DataTables));
                break;
            case "scenario":
                details.Add(ExcelScenarios.Add(ExcelPaths.Resolve(workbook, op.Path), op, index, post.Scenarios));
                break;
            case "equation":
                // xlsx has no math-object model: Excel carries cell formulas, not
                // OMML equations (docx and pptx do). This is N/A by design, not a gap.
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"ops[{index}]: Excel has no equation object — spreadsheets use cell formulas, not OMML math.",
                    "Put the math in a cell formula (e.g. set /Sheet1/A1 to \"=A2/2\"), or embed a rendered image " +
                    "of the equation; LaTeX-to-OMML equations are a docx/pptx feature (add type 'equation' there).");
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
    private static void AddPivot(XLWorkbook workbook, EditOp op, int index, List<object> details, PostSaveWork post)
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

        details.Add(ExcelPivots.Add(workbook, target.Sheet, op, index, post.CalculatedFields));
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

    /// <summary>
    /// Validates an <c>add slicer</c> op against the live workbook (the source
    /// table/pivot and its column/field must exist) and queues the raw OpenXml
    /// write for the post-save pass.
    /// </summary>
    private static void AddSlicer(
        XLWorkbook workbook, EditOp op, int index, List<object> details, ExcelSlicers.Batch slicers)
    {
        var target = ExcelPaths.Resolve(workbook, op.Path);
        var spec = ExcelSlicers.ParseAdd(workbook, target, op, index);
        var slicerIndex = slicers.Add(spec);
        details.Add(new
        {
            op = "add",
            type = "slicer",
            path = string.Create(
                CultureInfo.InvariantCulture,
                $"{ExcelPaths.SheetPath(target.Sheet)}/slicer[{slicerIndex}]"),
            source = spec.Source == SlicerSource.Table ? "table" : "pivot",
            column = spec.Column,
            caption = spec.Caption,
        });
    }

    /// <summary>
    /// Validates an <c>add formControl</c> op against the live workbook and queues
    /// the raw OpenXml write (VML drawing + ctrlProp + worksheet controls) for the
    /// post-save pass; ClosedXML cannot author these parts.
    /// </summary>
    private static void AddFormControl(
        XLWorkbook workbook, EditOp op, int index, List<object> details, ExcelFormControls.Batch formControls)
    {
        var target = ExcelPaths.Resolve(workbook, op.Path);
        var spec = ExcelFormControls.ParseAdd(target, op, index);
        var controlIndex = formControls.Add(spec);
        details.Add(new
        {
            op = "add",
            type = "formControl",
            path = ExcelPaths.FormControlPath(target.Sheet, controlIndex),
            kind = op.Props is not null && op.Props.TryGetPropertyValue("kind", out var kindNode) &&
                   kindNode is JsonValue kindValue && kindValue.GetValueKind() == JsonValueKind.String
                ? kindValue.GetValue<string>()
                : null,
            cell = spec.Cell,
            linkedCell = spec.LinkedCell,
        });
    }

    /// <summary>
    /// Validates an <c>add embed</c> op against the live workbook (the sheet must
    /// exist and the src must resolve inside the sandbox) and queues the raw
    /// EmbeddedPackagePart write for the post-save pass.
    /// </summary>
    private static void AddEmbed(
        CommandContext ctx, XLWorkbook workbook, EditOp op, int index, List<object> details, ExcelEmbeds.Batch embeds)
    {
        var target = ExcelPaths.Resolve(workbook, op.Path);
        var spec = ExcelEmbeds.ParseAdd(ctx.Workspace, target, op, index);
        var embedIndex = embeds.Add(spec);
        details.Add(new
        {
            op = "add",
            type = "embed",
            path = ExcelPaths.EmbedPath(target.Sheet, embedIndex),
            name = spec.Name,
            mediaType = spec.MediaType,
            size = (long)spec.Payload.Length,
            anchor = spec.Anchor,
        });
    }

    // ----- remove -----------------------------------------------------------

    private static void ApplyRemove(
        CommandContext ctx, XLWorkbook workbook, EditOp op, int index, List<object> details,
        ChartOpBatch charts, ExcelSlicers.Batch slicers, ExcelEmbeds.Batch embeds,
        ExcelFormControls.Batch formControls, PostSaveWork post)
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

            case ExcelTargetKind.Embed:
            {
                // Resolve the index against the batch-projected on-disk state and
                // queue the raw part deletion + registry update for the post-save pass.
                var removed = embeds.Remove(
                    ctx.File!, target.Sheet.Name, ExcelPaths.SheetPath(target.Sheet), target.EmbedIndex!.Value, index);
                details.Add(new
                {
                    op = "remove",
                    path = ExcelPaths.EmbedPath(target.Sheet, target.EmbedIndex!.Value),
                    removed = "embed",
                    name = removed.Name,
                });
                return;
            }
        }

        // Scenario uses a scenario[@name=…] form the shared DocPath grammar cannot
        // parse (like pivot[@name=…], which is peeled off earlier); build its
        // canonical path directly instead of through DocPath.
        var canonical = target.Kind == ExcelTargetKind.Scenario
            ? ExcelScenarios.ScenarioPath(target.Sheet, target.ScenarioName!)
            : DocPath.Parse(op.Path).ToCanonicalString();
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

            case ExcelTargetKind.Slicer:
                slicers.Remove(target.Sheet.Name, ExcelPaths.SheetPath(target.Sheet), target.SlicerIndex!.Value, index);
                details.Add(new { op = "remove", path = canonical, removed = "slicer" });
                break;

            case ExcelTargetKind.FormControl:
                formControls.Remove(
                    target.Sheet.Name, ExcelPaths.SheetPath(target.Sheet), target.FormControlIndex!.Value, index);
                details.Add(new { op = "remove", path = canonical, removed = "formControl" });
                break;

            case ExcelTargetKind.DataTable:
            {
                // The data table lives entirely in raw bytes (the t="dataTable"
                // anchor + cached body); validate the index against the on-disk
                // file now and queue the raw clearance for the post-save pass.
                var tablesOnSheet = ExcelDataTables.ReadOnSheet(ctx.File!, target.Sheet.Name);
                var wanted = target.DataTableIndex!.Value;
                if (tablesOnSheet.All(t => t.Index != wanted))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidPath,
                        $"ops[{index}]: no dataTable[{wanted}] on sheet '{target.Sheet.Name}' " +
                        $"({tablesOnSheet.Count} data table(s) exist).",
                        tablesOnSheet.Count > 0
                            ? "Data-table indices are 1-based per sheet; pick one of the candidates."
                            : "This sheet has no data tables to remove.",
                        candidates: tablesOnSheet.Count > 0
                            ? [.. tablesOnSheet.Select(t => ExcelPaths.DataTablePath(target.Sheet, t.Index))]
                            : [ExcelPaths.SheetPath(target.Sheet)]);
                }

                post.RemovedDataTables.Add((target.Sheet.Name, wanted));
                details.Add(new { op = "remove", path = canonical, removed = "dataTable" });
                break;
            }

            case ExcelTargetKind.Scenario:
            {
                // (1.5) Scenarios live in the raw scenarios part; validate the name
                // against the on-disk file and queue the raw removal post-save.
                var scenariosOnSheet = ExcelScenarios.ReadOnSheet(ctx.File!, target.Sheet.Name);
                var wantedName = target.ScenarioName!;
                if (scenariosOnSheet.All(s => !string.Equals(s.Name, wantedName, StringComparison.OrdinalIgnoreCase)))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidPath,
                        $"ops[{index}]: no scenario named '{wantedName}' on sheet '{target.Sheet.Name}' " +
                        $"({scenariosOnSheet.Count} scenario(s) exist).",
                        scenariosOnSheet.Count > 0
                            ? "Pick one of the candidates; scenario names are case-insensitive."
                            : "This sheet has no scenarios to remove.",
                        candidates: scenariosOnSheet.Count > 0
                            ? [.. scenariosOnSheet.Select(s => ExcelScenarios.ScenarioPath(target.Sheet, s.Name))]
                            : [ExcelPaths.SheetPath(target.Sheet)]);
                }

                post.RemovedScenarios.Add((target.Sheet.Name, wantedName));
                details.Add(new { op = "remove", path = canonical, removed = "scenario" });
                break;
            }

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

        // (1.4) Financial iterative cells (RATE/IRR/XIRR) and (1.5) the scalar
        // functions ClosedXML cannot evaluate (XLOOKUP/IFS/SWITCH/LET/MAXIFS/…) are
        // computed HERE and written back raw after the save, so they carry a real
        // cached value instead of riding the formula_not_evaluated warning.
        var computed = new List<(string Sheet, string Address, XLCellValue Value)>();

        foreach (var sheet in workbook.Worksheets)
        {
            // Materialize the formula cells before iterating: the scalar evaluator
            // (1.5) writes/clears a far-off scratch cell to delegate sub-expression
            // arithmetic to ClosedXML, which mutates the model mid-scan.
            foreach (var cell in sheet.CellsUsed().Where(c => c.HasFormula).ToList())
            {
                XLCellValue value;
                try
                {
                    value = cell.Value; // forces evaluation of dirty formulas
                }
                catch (Exception)
                {
                    if (TryEvaluateFinancial(sheet, cell, computed) ||
                        TryEvaluateScalar(sheet, cell, computed) || IsSpilledArrayAnchor(cell))
                    {
                        continue;
                    }

                    unevaluated.Add((sheet.Name, cell.Address.ToString()!));
                    continue;
                }

                if (value.IsError && value.GetError() == XLError.NameNotRecognized)
                {
                    // A dynamic-array anchor (FILTER/UNIQUE/…) round-tripping through
                    // a later edit keeps the cached spill values aioffice wrote raw;
                    // never strip them or warn (the result is already on disk).
                    if (TryEvaluateFinancial(sheet, cell, computed) ||
                        TryEvaluateScalar(sheet, cell, computed) || IsSpilledArrayAnchor(cell))
                    {
                        continue;
                    }

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
            ExcelCoreProperties.NormalizeAfterSave(file); // core props → docProps/core.xml
            return
            [
                new Warning(
                    ErrorCodes.FormulaNotEvaluated,
                    "Formula evaluation failed during save; the file was saved without cached values. " +
                    "Excel will compute all formulas when the file is opened."),
            ];
        }

        // ClosedXML writes core properties into a non-standard .psmdcp part; move
        // them to the conventional docProps/core.xml so unzip and Office see them.
        ExcelCoreProperties.NormalizeAfterSave(file);

        StripStaleCachedValues(file, unevaluated);
        WriteFinancialCachedValues(file, computed);

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

    /// <summary>
    /// (1.4) Computes one financial iterative formula (RATE/IRR/XIRR) ClosedXML
    /// could not evaluate and records its value for the raw write-back. Returns
    /// false when the cell is not one of those functions (so it keeps riding the
    /// normal unevaluated path).
    /// </summary>
    private static bool TryEvaluateFinancial(
        IXLWorksheet sheet, IXLCell cell, List<(string Sheet, string Address, XLCellValue Value)> sink)
    {
        if (!cell.HasFormula || !ExcelFinancialFunctions.IsFinancial(cell.FormulaA1))
        {
            return false;
        }

        var value = ExcelFinancialFunctions.Evaluate(sheet, cell.FormulaA1);
        sink.Add((sheet.Name, cell.Address.ToString()!, value));
        return true;
    }

    /// <summary>
    /// (1.5) Computes one scalar formula (XLOOKUP/IFS/SWITCH/LET/MAXIFS/MINIFS/
    /// AVERAGEIFS) ClosedXML returned #NAME? for and records its value for the raw
    /// write-back. Returns false when the cell is not one of those functions, OR
    /// when the evaluator could not evaluate it (a nested unevaluable function) —
    /// in which case the cell honestly keeps the formula_not_evaluated warning.
    /// </summary>
    private static bool TryEvaluateScalar(
        IXLWorksheet sheet, IXLCell cell, List<(string Sheet, string Address, XLCellValue Value)> sink)
    {
        if (!cell.HasFormula || !ExcelScalarFunctions.IsScalarFunction(cell.FormulaA1))
        {
            return false;
        }

        var value = ExcelScalarFunctions.Evaluate(sheet, cell.FormulaA1);
        if (value.IsError && value.GetError() == XLError.NameNotRecognized)
        {
            return false; // honest fallback: a sub-expression used an unevaluable function
        }

        sink.Add((sheet.Name, cell.Address.ToString()!, value));
        return true;
    }

    /// <summary>
    /// (1.4) True when the cell is a dynamic-array anchor whose spill aioffice
    /// already wrote raw (recognized by the array formula's multi-cell reference or
    /// the function name). Such a cell keeps its cached value across re-saves — it
    /// is neither stripped nor flagged formula_not_evaluated.
    /// </summary>
    private static bool IsSpilledArrayAnchor(IXLCell cell)
    {
        if (!cell.HasFormula)
        {
            return false;
        }

        if (cell.FormulaReference is { } reference &&
            (reference.FirstAddress.RowNumber != reference.LastAddress.RowNumber ||
             reference.FirstAddress.ColumnNumber != reference.LastAddress.ColumnNumber))
        {
            return true;
        }

        return ExcelDynamicArrays.IsDynamicArrayFormula(cell.FormulaA1);
    }

    /// <summary>
    /// Writes the computed cached values (financial 1.4 + scalar 1.5) into the
    /// saved cells, preserving each cell's formula text. Handles every value type —
    /// numbers, text (interned in the shared-string table), booleans, dates and
    /// error values — because the scalar evaluators (XLOOKUP/IFS/SWITCH/…) return
    /// non-numeric results too.
    /// </summary>
    private static void WriteFinancialCachedValues(
        string file, List<(string Sheet, string Address, XLCellValue Value)> cells)
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

            var byAddress = group.ToDictionary(g => g.Address, g => g.Value, StringComparer.OrdinalIgnoreCase);
            var dirty = false;
            foreach (var cell in worksheetRoot.Descendants<Cell>())
            {
                if (cell.CellReference?.Value is { } reference && byAddress.TryGetValue(reference, out var value))
                {
                    WriteCachedValue(cell, value);
                    dirty = true;
                }
            }

            if (dirty)
            {
                worksheetRoot.Save();
            }
        }
    }

    /// <summary>Writes one typed cached value onto a saved formula cell (formula text preserved).</summary>
    private static void WriteCachedValue(Cell cell, XLCellValue value)
    {
        var culture = System.Globalization.CultureInfo.InvariantCulture;
        switch (value.Type)
        {
            case XLDataType.Number:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(value.GetNumber().ToString("R", culture));
                break;
            case XLDataType.Boolean:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(value.GetBoolean() ? "1" : "0");
                break;
            case XLDataType.DateTime:
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(value.GetDateTime().ToOADate().ToString("R", culture));
                break;
            case XLDataType.Error:
                cell.DataType = CellValues.Error;
                cell.CellValue = new CellValue(value.ToString(culture));
                break;
            case XLDataType.Blank:
                cell.DataType = null;
                cell.CellValue = null;
                break;
            default:
                // A formula cell carrying a string result uses t="str" (an inline
                // formula-string), not a shared string — the validator-clean form.
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(value.GetText());
                break;
        }
    }
}
