using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// The xlsx pivot-table layer, built on ClosedXML's native XLPivotTable support
/// (validated: files save with a populated pivot cache, <c>refreshOnLoad</c> set,
/// and zero OpenXmlValidator errors, so Excel shows a working pivot on open).
/// The one thing ClosedXML does not expose publicly is the cache's source
/// range, so <see cref="ReadSources"/> reads it back from the raw
/// <c>pivotCacheDefinition</c> parts.
/// </summary>
internal static partial class ExcelPivots
{
    /// <summary>The aggregations an add op accepts (wire names).</summary>
    public static readonly IReadOnlyList<string> Aggs = ["sum", "count", "average", "min", "max"];

    /// <summary>
    /// The show-values-as modes a pivot value field accepts (wire names). Every mode
    /// here is APPLIED (the latent v1.0–1.12 bug — a <c>showAs</c> silently accepted
    /// and never written — is gone): an unknown name is <c>invalid_args</c> with these
    /// as candidates, and a mode ClosedXML/OOXML cannot express is rejected with
    /// <c>unsupported_feature</c> (see <see cref="UnsupportedShowAs"/>), never ignored.
    /// </summary>
    public static readonly IReadOnlyList<string> ShowAsModes =
    [
        "normal", "percentOfTotal", "percentOfColumn", "percentOfRow", "runningTotal",
        "differenceFrom", "percentDifferenceFrom", "percentOf", "index",
        // Accepted names, but rejected as unsupported_feature (OOXML showDataAs has no slot):
        "percentOfParentTotal", "rankAscending", "rankDescending",
    ];

    /// <summary>Modes a caller may name but OOXML's <c>showDataAs</c> cannot carry → <c>unsupported_feature</c>.</summary>
    private static readonly IReadOnlyDictionary<string, string> UnsupportedShowAs =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["percentOfParentTotal"] = "ClosedXML 0.105 / OOXML showDataAs only models a flat percentOfTotal, not the percentOfParent* family.",
            ["rankAscending"] = "OOXML showDataAs has no rank slot (it is an Excel-2010 dataField ext); ClosedXML cannot author it.",
            ["rankDescending"] = "OOXML showDataAs has no rank slot (it is an Excel-2010 dataField ext); ClosedXML cannot author it.",
        };

    /// <summary>Modes that require a <c>baseField</c> (and, for the difference pair, a <c>baseItem</c>).</summary>
    private static readonly IReadOnlySet<string> NeedsBaseField =
        new HashSet<string>(StringComparer.Ordinal) { "runningTotal", "differenceFrom", "percentDifferenceFrom", "percentOf" };

    private static readonly IReadOnlySet<string> NeedsBaseItem =
        new HashSet<string>(StringComparer.Ordinal) { "differenceFrom", "percentDifferenceFrom", "percentOf" };

    /// <summary>The OOXML <c>showDataAs</c> attribute value for each applied (non-normal) mode.</summary>
    private static readonly IReadOnlyDictionary<string, S.ShowDataAsValues> ShowDataAsOf =
        new Dictionary<string, S.ShowDataAsValues>(StringComparer.Ordinal)
        {
            ["percentOfTotal"] = S.ShowDataAsValues.PercentOfTotal,
            ["percentOfColumn"] = S.ShowDataAsValues.PercentOfColumn,
            ["percentOfRow"] = S.ShowDataAsValues.PercentOfRaw, // OOXML spells "% of row" percentOfRow
            ["runningTotal"] = S.ShowDataAsValues.RunTotal,
            ["differenceFrom"] = S.ShowDataAsValues.Difference,
            ["percentDifferenceFrom"] = S.ShowDataAsValues.PercentageDifference,
            ["percentOf"] = S.ShowDataAsValues.Percent,
            ["index"] = S.ShowDataAsValues.Index,
        };

    /// <summary>The OOXML <c>baseItem</c> sentinels (ECMA-376 ST_Index defaults): previous/next item.</summary>
    private const uint BaseItemPrevious = 1048828u;
    private const uint BaseItemNext = 1048829u;

    private static readonly IReadOnlyList<string> AddProps =
        ["sourceRange", "targetSheet", "targetAnchor", "rows", "columns", "values", "filters", "name", "calculatedFields", "grandTotals"];

    /// <summary>The string forms <c>grandTotals</c> accepts; the object form is <c>{rows,columns}</c>.</summary>
    private static readonly IReadOnlyList<string> GrandTotalModes = ["both", "rows", "columns", "none"];

    private static readonly IReadOnlyList<string> ValueProps = ["field", "agg", "showAs", "baseField", "baseItem"];

    [GeneratedRegex("^([A-Z]{1,3})([0-9]{1,7}):([A-Z]{1,3})([0-9]{1,7})$")]
    private static partial Regex RangePattern();

    [GeneratedRegex("^([A-Z]{1,3})([0-9]{1,7})$")]
    private static partial Regex CellPattern();

    // ----- add ---------------------------------------------------------------

    /// <summary>
    /// Validates and applies an <c>add pivot</c> op against the in-memory
    /// workbook (the whole batch aborts before any byte is written if this
    /// throws). Returns the details entry for the envelope. Any
    /// <c>calculatedFields</c> are validated against the source headers and
    /// queued on <paramref name="calculatedFields"/> for the post-save raw pass
    /// (ClosedXML has no calculated-field model).
    /// </summary>
    public static object Add(
        XLWorkbook workbook, IXLWorksheet sourceSheet, EditOp op, int opIndex,
        List<CalculatedFieldSpec> calculatedFields, List<ShowAsSpec> showValuesAs)
    {
        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add pivot needs props.",
                "Pass props like {\"sourceRange\":\"A1:D20\",\"targetSheet\":\"Pivot\"," +
                "\"rows\":[\"Region\"],\"values\":[{\"field\":\"Sales\",\"agg\":\"sum\"}]}.");
        }

        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown pivot prop '{key}'.",
                    "Supported pivot props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }

        var (firstColumn, firstRow, lastColumn, lastRow) =
            ParseSourceRange(RequiredString(props, "sourceRange", opIndex, "A1:D20"), opIndex);
        if (lastRow == firstRow)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: sourceRange needs a header row plus at least one data row.",
                "Extend the range to include the data under the headers, e.g. A1:D20.");
        }

        var headers = ReadHeaders(sourceSheet, firstColumn, firstRow, lastColumn, opIndex);

        var rows = StringList(props, "rows", opIndex);
        var columns = StringList(props, "columns", opIndex);
        var filters = StringList(props, "filters", opIndex);
        var values = ParseValues(props, opIndex);
        if (rows.Count == 0 && columns.Count == 0 && filters.Count == 0 && values.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: the pivot has no fields; pass at least one of rows, columns, values, filters.",
                "Example: {\"rows\":[\"Region\"],\"values\":[{\"field\":\"Sales\",\"agg\":\"sum\"}]}.");
        }

        // Resolve every requested field against the actual headers (the error
        // contract: unknown field -> invalid_args with the real headers).
        rows = [.. rows.Select(f => ResolveField(headers, f, opIndex))];
        columns = [.. columns.Select(f => ResolveField(headers, f, opIndex))];
        filters = [.. filters.Select(f => ResolveField(headers, f, opIndex))];
        values = [.. values.Select(v => v with
        {
            Field = ResolveField(headers, v.Field, opIndex),
            BaseField = v.BaseField is null ? null : ResolveField(headers, v.BaseField, opIndex),
        })];

        GuardAxisOverlap(rows, columns, filters, opIndex);
        GuardDuplicateValues(values, opIndex);

        // Calculated fields reference the source headers; validate names + formula
        // refs here so the whole op aborts before any byte is written on a bad ref.
        var calcFields = ParseCalculatedFields(props, headers, opIndex);

        // Grand-total visibility (omitted => null => today's always-both default,
        // left BYTE-IDENTICAL: nothing is set on the model below).
        var grandTotals = ParseGrandTotals(props, opIndex);

        var targetSheetName = RequiredString(props, "targetSheet", opIndex, "Pivot");
        if (!workbook.TryGetWorksheet(targetSheetName, out var targetSheet))
        {
            targetSheet = AddSheet(workbook, targetSheetName, opIndex);
        }

        var anchor = OptionalString(props, "targetAnchor") ?? "A1";
        if (!CellPattern().IsMatch(anchor.ToUpperInvariant()))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{anchor}' is not a usable targetAnchor.",
                "targetAnchor is the top-left cell of the pivot on the target sheet, e.g. A1.");
        }

        var name = OptionalString(props, "name") ?? DefaultName(targetSheet);
        if (targetSheet.PivotTables.Contains(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a pivot table named '{name}' already exists on sheet '{targetSheet.Name}'.",
                "Pick a different name, or remove the existing pivot first with " +
                "{op:remove, path:" + ExcelPaths.PivotPath(targetSheet, name) + "}.");
        }

        var source = sourceSheet.Range(firstRow, firstColumn, lastRow, lastColumn);
        IXLPivotTable pivot;
        try
        {
            pivot = targetSheet.PivotTables.Add(name, targetSheet.Cell(anchor.ToUpperInvariant()), source);
            foreach (var field in rows)
            {
                pivot.RowLabels.Add(field);
            }

            foreach (var field in columns)
            {
                pivot.ColumnLabels.Add(field);
            }

            foreach (var field in filters)
            {
                pivot.ReportFilters.Add(field);
            }

            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var value in values)
            {
                var pivotValue = pivot.Values
                    .Add(value.Field, UniqueValueName(value, usedNames))
                    .SetSummaryFormula(SummaryOf(value.Agg));

                // Set the in-memory show-values-as so anything reading the live model
                // (and ClosedXML's own cache compute, where it supports the mode) sees
                // it. ClosedXML 0.105's WRITER drops these attributes on save, so the
                // authoritative write is the raw post-save pass below — but the
                // in-memory call keeps the model self-consistent and round-trippable.
                ApplyInMemory(pivotValue, value);
            }

            // Grand-total visibility: only touch the model when the caller asked,
            // so an omitted grandTotals leaves ClosedXML's always-both write byte-stable.
            if (grandTotals is { } gt)
            {
                pivot.SetShowGrandTotalsRows(gt.Rows).SetShowGrandTotalsColumns(gt.Columns);
            }

            pivot.PivotCache.Refresh(); // populate cached field items + records from the live data
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: could not create the pivot table: {exception.Message}",
                "Check that the source headers are unique and the pivot name is unused on the target sheet.",
                innerException: exception);
        }

        // Queue the validated calculated fields for the post-save raw pass, keyed
        // by the pivot's (target sheet, name) so the writer can find its parts.
        foreach (var calc in calcFields)
        {
            calculatedFields.Add(calc with { TargetSheet = targetSheet.Name, PivotName = name });
        }

        // Queue every non-normal show-values-as for the post-save raw pass: ClosedXML
        // 0.105 sets the in-memory Calculation but drops showDataAs/baseField/baseItem
        // on save, so the dataField attributes are authored directly on the bytes.
        // The dataField order matches the value-add order (calculated fields append
        // AFTER), so the value index here is its dataField position.
        for (var i = 0; i < values.Count; i++)
        {
            if (!string.Equals(values[i].ShowAs, "normal", StringComparison.Ordinal))
            {
                showValuesAs.Add(new ShowAsSpec(
                    targetSheet.Name, name, i, values[i].ShowAs, values[i].BaseField, values[i].BaseItem));
            }
        }

        return new
        {
            op = "add",
            type = "pivot",
            path = ExcelPaths.PivotPath(targetSheet, name),
            name,
            sourceSheet = sourceSheet.Name,
            sourceRange = string.Create(
                CultureInfo.InvariantCulture,
                $"{ExcelCharts.ColumnLetters(firstColumn)}{firstRow}:{ExcelCharts.ColumnLetters(lastColumn)}{lastRow}"),
            targetSheet = targetSheet.Name,
            location = anchor.ToUpperInvariant(),
            rows,
            columns,
            values = values.Select(ValueDetail).ToList(),
            filters,
            calculatedFields = calcFields.Select(c => new { name = c.Name, formula = c.Formula }).ToList(),
            grandTotals = new { rows = pivot.ShowGrandTotalsRows, columns = pivot.ShowGrandTotalsColumns },
        };
    }

    /// <summary>The per-value details entry: field + agg, plus showAs (and its base) when not the default normal.</summary>
    private static object ValueDetail(ValueSpec v) =>
        string.Equals(v.ShowAs, "normal", StringComparison.Ordinal)
            ? new { field = v.Field, agg = v.Agg }
            : new { field = v.Field, agg = v.Agg, showAs = v.ShowAs, baseField = v.BaseField, baseItem = v.BaseItem };

    // ----- find / describe ---------------------------------------------------

    /// <summary>
    /// Finds the pivot a resolved <see cref="ExcelTarget"/> addresses, by 1-based
    /// index or by name (case-insensitive). Throws <c>invalid_path</c> with the
    /// sheet's actual pivot paths as candidates.
    /// </summary>
    public static (IXLPivotTable Pivot, int Index) Find(ExcelTarget target)
    {
        var pivots = target.Sheet.PivotTables.ToList();
        if (target.PivotName is { } wanted)
        {
            for (var i = 0; i < pivots.Count; i++)
            {
                if (string.Equals(pivots[i].Name, wanted, StringComparison.OrdinalIgnoreCase))
                {
                    return (pivots[i], i + 1);
                }
            }

            throw NotFound(target.Sheet, pivots, $"No pivot table named '{wanted}' on sheet '{target.Sheet.Name}'", wanted);
        }

        var index = target.PivotIndex!.Value;
        if (index > pivots.Count)
        {
            throw NotFound(
                target.Sheet, pivots,
                $"No pivot[{index}] on sheet '{target.Sheet.Name}' ({pivots.Count} pivot table(s) exist)",
                requested: null);
        }

        return (pivots[index - 1], index);
    }

    private static AiofficeException NotFound(
        IXLWorksheet sheet, List<IXLPivotTable> pivots, string message, string? requested)
    {
        var ordered = requested is null
            ? pivots
            : [.. pivots.OrderBy(p => ExcelPaths.Levenshtein(requested, p.Name))];
        return new AiofficeException(
            ErrorCodes.InvalidPath,
            message + ".",
            pivots.Count > 0
                ? "Pivot indices are 1-based per sheet; pick one of the candidates (the [@name=…] form is stable)."
                : "This sheet has no pivot tables; add one with {op:add, type:pivot, path:/SourceSheet, " +
                  "props:{sourceRange:\"A1:D20\", targetSheet:\"" + sheet.Name + "\", rows:[…], values:[…]}}.",
            candidates: pivots.Count > 0
                ? [.. ordered.Take(5).Select(p => ExcelPaths.PivotPath(sheet, p.Name))]
                : [ExcelPaths.SheetPath(sheet)]);
    }

    /// <summary>One pivot as agents see it (get and read --view structure).</summary>
    public static object Describe(
        IXLWorksheet sheet, IXLPivotTable pivot,
        IReadOnlyDictionary<(string Sheet, string Pivot), (string? SourceSheet, string? SourceRange)> sources,
        IReadOnlyDictionary<(string Sheet, string Pivot), List<(string Name, string Formula)>>? calculatedFields = null)
    {
        sources.TryGetValue((sheet.Name, pivot.Name), out var source);
        var calc = calculatedFields is not null && calculatedFields.TryGetValue((sheet.Name, pivot.Name), out var c)
            ? c
            : [];
        // The calculated fields surface in ClosedXML's Values too (they ARE data
        // fields); list them only under calculatedFields, not in values.
        var calcNames = new HashSet<string>(calc.Select(f => f.Name), StringComparer.OrdinalIgnoreCase);
        return new
        {
            path = ExcelPaths.PivotPath(sheet, pivot.Name),
            kind = "pivot",
            sheet = sheet.Name,
            name = pivot.Name,
            sourceSheet = source.SourceSheet,
            sourceRange = source.SourceRange,
            location = pivot.TargetCell.Address.ToString(),
            rows = pivot.RowLabels.Select(f => f.SourceName).ToList(),
            columns = pivot.ColumnLabels.Select(f => f.SourceName).ToList(),
            values = pivot.Values
                .Where(v => !calcNames.Contains(v.SourceName))
                .Select(DescribeValue)
                .ToList(),
            filters = pivot.ReportFilters.Select(f => f.SourceName).ToList(),
            calculatedFields = calc.Select(f => new { name = f.Name, formula = f.Formula }).ToList(),
            grandTotals = new { rows = pivot.ShowGrandTotalsRows, columns = pivot.ShowGrandTotalsColumns },
        };
    }

    /// <summary>
    /// One value field as <c>get</c> reports it: field + agg, plus showAs (and its
    /// base) when ClosedXML read a non-normal calculation back from the file's
    /// <c>showDataAs</c>. The (previous)/(next) base-item sentinels surface as their
    /// readable names.
    /// </summary>
    private static object DescribeValue(IXLPivotValue value)
    {
        var showAs = ShowAsName(value.Calculation);
        if (string.Equals(showAs, "normal", StringComparison.Ordinal))
        {
            return new { field = value.SourceName, agg = AggName(value.SummaryFormula) };
        }

        var baseField = string.IsNullOrEmpty(value.BaseFieldName) ? null : value.BaseFieldName;
        // baseItem is only meaningful for the difference family; runningTotal et al.
        // carry an unused index that ClosedXML surfaces as (previous) — suppress it.
        var baseItem = NeedsBaseItem.Contains(showAs)
            ? value.CalculationItem switch
            {
                XLPivotCalculationItem.Previous => "(previous)",
                XLPivotCalculationItem.Next => "(next)",
                _ => value.BaseItemValue.IsBlank ? null : value.BaseItemValue.ToString(CultureInfo.InvariantCulture),
            }
            : null;
        return new
        {
            field = value.SourceName,
            agg = AggName(value.SummaryFormula),
            showAs,
            baseField,
            baseItem,
        };
    }

    /// <summary>Wire show-values-as name for a calculation read back from a file (ClosedXML's enum).</summary>
    private static string ShowAsName(XLPivotCalculation calculation) => calculation switch
    {
        XLPivotCalculation.Normal => "normal",
        XLPivotCalculation.PercentageOfTotal => "percentOfTotal",
        XLPivotCalculation.PercentageOfColumn => "percentOfColumn",
        XLPivotCalculation.PercentageOfRow => "percentOfRow",
        XLPivotCalculation.RunningTotal => "runningTotal",
        XLPivotCalculation.DifferenceFrom => "differenceFrom",
        XLPivotCalculation.PercentageDifferenceFrom => "percentDifferenceFrom",
        XLPivotCalculation.PercentageOf => "percentOf",
        XLPivotCalculation.Index => "index",
        _ => "normal",
    };

    // ----- post-save part sync ------------------------------------------------

    /// <summary>
    /// The pivot names the in-memory workbook considers alive, per sheet —
    /// the input for <see cref="SyncPartsAfterSave"/>.
    /// </summary>
    public static Dictionary<string, HashSet<string>> AliveNames(XLWorkbook workbook)
    {
        var result = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in workbook.Worksheets)
        {
            result[sheet.Name] = sheet.PivotTables
                .Select(p => p.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        return result;
    }

    /// <summary>
    /// Deletes pivot parts the in-memory model no longer owns. Measured
    /// ClosedXML 0.105 defect: <c>PivotTables.Delete(name)</c> removes the
    /// pivot from the model, but its <c>pivotTable</c> part survives the save
    /// and resurrects the pivot on the next load. This raw pass removes those
    /// orphan parts, plus any cache definition (and its records) no remaining
    /// pivot references, and the workbook's dangling <c>pivotCache</c> entry.
    /// </summary>
    public static void SyncPartsAfterSave(string file, Dictionary<string, HashSet<string>> alive)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is not { } workbook)
        {
            return;
        }

        var keptCaches = new HashSet<PivotTableCacheDefinitionPart>();
        foreach (var sheet in workbook.Descendants<S.Sheet>())
        {
            if (sheet.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            alive.TryGetValue(sheet.Name?.Value ?? string.Empty, out var aliveNames);
            foreach (var pivotPart in worksheetPart.PivotTableParts.ToList())
            {
                var name = pivotPart.PivotTableDefinition?.Name?.Value;
                if (name is not null && aliveNames?.Contains(name) == true)
                {
                    if (pivotPart.PivotTableCacheDefinitionPart is { } usedCache)
                    {
                        keptCaches.Add(usedCache);
                    }

                    continue;
                }

                worksheetPart.DeletePart(pivotPart);
            }
        }

        var workbookDirty = false;
        foreach (var cachePart in workbookPart.PivotTableCacheDefinitionParts.ToList())
        {
            if (keptCaches.Contains(cachePart))
            {
                continue;
            }

            var cacheRelId = workbookPart.GetIdOfPart(cachePart);
            var entry = workbook.Elements<S.PivotCaches>()
                .SelectMany(c => c.Elements<S.PivotCache>())
                .FirstOrDefault(pc => pc.Id?.Value == cacheRelId);
            if (entry is not null)
            {
                entry.Remove();
                workbookDirty = true;
            }

            workbookPart.DeletePart(cachePart);
        }

        foreach (var caches in workbook.Elements<S.PivotCaches>().Where(c => !c.HasChildren).ToList())
        {
            caches.Remove();
            workbookDirty = true;
        }

        if (workbookDirty)
        {
            workbook.Save();
        }
    }

    // ----- calculated-field raw authoring (v1.3) ----------------------------------

    /// <summary>
    /// Authors the queued pivot calculated fields on the saved file. For each
    /// field: a <c>cacheField</c> (name + <c>formula</c> + <c>databaseField=0</c>)
    /// is appended to its pivot's cache definition, a placeholder <c>&lt;m/&gt;</c>
    /// value is appended to every cache record (so ClosedXML — which has no
    /// calculated-field model — can still reopen the file), and a matching
    /// <c>pivotField</c> / <c>dataField</c> are appended to the pivot table
    /// definition so the field shows as a value column. Validator-clean; real
    /// Excel recomputes the field from its formula on open.
    /// </summary>
    public static void ApplyCalculatedFields(string file, IReadOnlyList<CalculatedFieldSpec> fields)
    {
        if (fields.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return;
        }

        // Group by the pivot the field belongs to; each pivot's cache is mutated
        // once even when it gains several fields.
        foreach (var byPivot in fields.GroupBy(f => (f.TargetSheet, f.PivotName)))
        {
            var (pivotPart, cachePart) = LocatePivotParts(workbookPart, byPivot.Key.TargetSheet, byPivot.Key.PivotName);
            if (pivotPart?.PivotTableDefinition is not { } definition ||
                cachePart?.PivotCacheDefinition is not { } cacheDefinition)
            {
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    $"Pivot '{byPivot.Key.PivotName}' on '{byPivot.Key.TargetSheet}' disappeared before its calculated fields were written.",
                    "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
            }

            var cacheFields = cacheDefinition.CacheFields
                ?? cacheDefinition.AppendChild(new S.CacheFields());
            var pivotFields = definition.PivotFields
                ?? definition.InsertAfter(new S.PivotFields(), definition.Location);
            var dataFields = definition.DataFields;

            foreach (var field in byPivot)
            {
                // 1) cacheField (calculated): name + formula + databaseField=0.
                cacheFields.AppendChild(new S.CacheField(new S.SharedItems())
                {
                    Name = field.Name,
                    DatabaseField = false,
                    Formula = field.Formula,
                });
                var newFieldIndex = cacheFields.Elements<S.CacheField>().Count() - 1; // 0-based

                // 2) per-record placeholder value so record arity == field count
                //    (keeps ClosedXML's records reader — and Excel — happy).
                if (cachePart.PivotTableCacheRecordsPart?.PivotCacheRecords is { } records)
                {
                    foreach (var record in records.Elements<S.PivotCacheRecord>())
                    {
                        record.AppendChild(new S.MissingItem());
                    }
                }

                // 3) pivotField marking it a data field (count must match cacheFields).
                pivotFields!.AppendChild(new S.PivotField { DataField = true, ShowAll = false });

                // 4) dataField exposing the calculated field as a value column.
                dataFields ??= definition.InsertAfter(new S.DataFields(), LastBeforeDataFields(definition));
                dataFields.AppendChild(new S.DataField
                {
                    Name = field.Name,
                    Field = (uint)newFieldIndex,
                    BaseField = 0,
                    BaseItem = 0u,
                });
            }

            cacheFields.Count = (uint)cacheFields.Elements<S.CacheField>().Count();
            pivotFields!.Count = (uint)pivotFields.Elements<S.PivotField>().Count();
            if (dataFields is not null)
            {
                dataFields.Count = (uint)dataFields.Elements<S.DataField>().Count();
            }

            if (cachePart.PivotTableCacheRecordsPart?.PivotCacheRecords is { } recs)
            {
                recs.Save();
            }

            cacheDefinition.Save();
            definition.Save();
        }
    }

    // ----- show-values-as raw authoring (v1.13) ---------------------------------

    /// <summary>
    /// A validated show-values-as for the post-save raw pass. <c>ValueIndex</c> is the
    /// 0-based position of the value field in the pivot's <c>dataFields</c> (= its add
    /// order). <c>ShowAs</c> is the wire mode (never <c>normal</c> — those are not
    /// queued). <c>BaseField</c>/<c>BaseItem</c> carry the resolved base for modes that
    /// need one; both names are looked up against the saved cache definition.
    /// </summary>
    public sealed record ShowAsSpec(
        string TargetSheet, string PivotName, int ValueIndex, string ShowAs, string? BaseField, string? BaseItem);

    /// <summary>
    /// Authors the queued show-values-as settings on the saved file. For each value
    /// field it sets the <c>dataField</c>'s <c>showDataAs</c> attribute (and, for the
    /// base-field modes, <c>baseField</c> = the base's 0-based cacheField index and
    /// <c>baseItem</c> = the item's 0-based sharedItems index, or the
    /// previous/next sentinel). ClosedXML's writer drops these, so this raw pass is the
    /// authoritative write; Excel recomputes the displayed values on open and ClosedXML
    /// reads the attributes back (so <see cref="Describe"/> reports them). Runs after
    /// the calculated-fields pass so every dataField is present.
    /// </summary>
    public static void ApplyShowValuesAs(string file, IReadOnlyList<ShowAsSpec> specs)
    {
        if (specs.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return;
        }

        foreach (var byPivot in specs.GroupBy(s => (s.TargetSheet, s.PivotName)))
        {
            var (pivotPart, cachePart) = LocatePivotParts(workbookPart, byPivot.Key.TargetSheet, byPivot.Key.PivotName);
            if (pivotPart?.PivotTableDefinition is not { } definition ||
                definition.DataFields is not { } dataFields ||
                cachePart?.PivotCacheDefinition?.CacheFields is not { } cacheFields)
            {
                throw new AiofficeException(
                    ErrorCodes.InternalError,
                    $"Pivot '{byPivot.Key.PivotName}' on '{byPivot.Key.TargetSheet}' disappeared before its show-values-as was written.",
                    "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
            }

            var dataFieldList = dataFields.Elements<S.DataField>().ToList();
            var cacheFieldList = cacheFields.Elements<S.CacheField>().ToList();

            foreach (var spec in byPivot)
            {
                if (spec.ValueIndex >= dataFieldList.Count)
                {
                    throw new AiofficeException(
                        ErrorCodes.InternalError,
                        $"Pivot '{byPivot.Key.PivotName}' value #{spec.ValueIndex} vanished before its show-values-as was written.",
                        "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
                }

                var dataField = dataFieldList[spec.ValueIndex];
                dataField.ShowDataAs = ShowDataAsOf[spec.ShowAs];

                if (spec.BaseField is { } baseField)
                {
                    var baseFieldIndex = cacheFieldList.FindIndex(
                        cf => string.Equals(cf.Name?.Value, baseField, StringComparison.OrdinalIgnoreCase));
                    if (baseFieldIndex < 0)
                    {
                        throw new AiofficeException(
                            ErrorCodes.InternalError,
                            $"Base field '{baseField}' for pivot '{byPivot.Key.PivotName}' is missing from its cache.",
                            "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
                    }

                    dataField.BaseField = baseFieldIndex;

                    // baseItem only carries meaning for the difference family; for
                    // runningTotal it is unused, so leave the dataField default rather
                    // than writing a sentinel Excel would ignore.
                    if (spec.BaseItem is not null)
                    {
                        dataField.BaseItem = ResolveBaseItemIndex(cacheFieldList[baseFieldIndex], spec.BaseItem);
                    }
                }
            }

            definition.Save();
        }
    }

    /// <summary>
    /// The OOXML <c>baseItem</c> index for a difference's base item: the
    /// (previous)/(next) sentinel passes through, otherwise the item's 0-based
    /// position in the base field's <c>sharedItems</c> (matched on its formatted value).
    /// A literal with no match becomes the previous sentinel — Excel then recomputes
    /// against the prior item rather than erroring on open.
    /// </summary>
    private static uint ResolveBaseItemIndex(S.CacheField baseCacheField, string baseItem)
    {
        if (BaseItemSentinel(baseItem) is { } sentinel)
        {
            return sentinel;
        }

        var items = baseCacheField.SharedItems?.Elements().ToList() ?? [];
        for (var i = 0; i < items.Count; i++)
        {
            if (string.Equals(SharedItemText(items[i]), baseItem, StringComparison.OrdinalIgnoreCase))
            {
                return (uint)i;
            }
        }

        return BaseItemPrevious;
    }

    /// <summary>The text of a cache <c>sharedItems</c> entry (string/number/bool/date item), or null for a missing item.</summary>
    private static string? SharedItemText(OpenXmlElement item) => item switch
    {
        S.StringItem s => s.Val?.Value,
        S.NumberItem n => n.Val?.Value is { } d ? d.ToString(CultureInfo.InvariantCulture) : null,
        S.BooleanItem b => b.Val?.Value is { } flag ? (flag ? "TRUE" : "FALSE") : null,
        S.DateTimeItem t => t.Val is { } dt ? dt.Value.ToString("o", CultureInfo.InvariantCulture) : null,
        _ => null,
    };

    /// <summary>
    /// The element a new <c>dataFields</c> goes after in a pivotTableDefinition:
    /// rowItems / colItems / pageFields / formats sit before it, but pivotFields
    /// is always present, so anchor on the last of the known predecessors.
    /// </summary>
    private static OpenXmlElement LastBeforeDataFields(S.PivotTableDefinition definition)
    {
        return (OpenXmlElement?)definition.GetFirstChild<S.ColumnItems>()
            ?? (OpenXmlElement?)definition.GetFirstChild<S.RowItems>()
            ?? (OpenXmlElement?)definition.GetFirstChild<S.ColumnFields>()
            ?? (OpenXmlElement?)definition.GetFirstChild<S.RowFields>()
            ?? (OpenXmlElement?)definition.GetFirstChild<S.PageFields>()
            ?? definition.PivotFields!;
    }

    /// <summary>Locates a pivot's table + cache parts by (target sheet, name).</summary>
    private static (PivotTablePart? Table, PivotTableCacheDefinitionPart? Cache) LocatePivotParts(
        WorkbookPart workbookPart, string sheetName, string pivotName)
    {
        var sheet = workbookPart.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is not { } relationshipId ||
            workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
        {
            return (null, null);
        }

        foreach (var pivotPart in worksheetPart.PivotTableParts)
        {
            if (string.Equals(pivotPart.PivotTableDefinition?.Name?.Value, pivotName, StringComparison.OrdinalIgnoreCase))
            {
                return (pivotPart, pivotPart.PivotTableCacheDefinitionPart);
            }
        }

        return (null, null);
    }

    /// <summary>
    /// The calculated fields of every pivot, keyed by (sheet, pivot name), read
    /// raw from the cache definitions (a <c>cacheField</c> with
    /// <c>databaseField=0</c> and a <c>formula</c>). Used by <see cref="Describe"/>.
    /// </summary>
    public static Dictionary<(string Sheet, string Pivot), List<(string Name, string Formula)>> ReadCalculatedFields(
        SpreadsheetDocument document)
    {
        var result = new Dictionary<(string, string), List<(string, string)>>(SheetPivotComparer.Instance);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return result;
        }

        foreach (var sheet in workbookPart.Workbook.Descendants<S.Sheet>())
        {
            if (sheet.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var sheetName = sheet.Name?.Value ?? string.Empty;
            foreach (var pivotPart in worksheetPart.PivotTableParts)
            {
                if (pivotPart.PivotTableDefinition?.Name?.Value is not { } pivotName ||
                    pivotPart.PivotTableCacheDefinitionPart?.PivotCacheDefinition?.CacheFields is not { } cacheFields)
                {
                    continue;
                }

                var calc = cacheFields.Elements<S.CacheField>()
                    .Where(f => f.DatabaseField?.Value == false && f.Formula?.Value is not null)
                    .Select(f => (f.Name?.Value ?? string.Empty, f.Formula!.Value!))
                    .ToList();
                if (calc.Count > 0)
                {
                    result[(sheetName, pivotName)] = calc;
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Source ranges per (sheet, pivot name), read from the raw
    /// <c>pivotCacheDefinition</c> parts (ClosedXML keeps them internal).
    /// </summary>
    public static Dictionary<(string Sheet, string Pivot), (string? SourceSheet, string? SourceRange)> ReadSources(
        SpreadsheetDocument document)
    {
        var result = new Dictionary<(string, string), (string?, string?)>(SheetPivotComparer.Instance);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return result;
        }

        foreach (var sheet in workbookPart.Workbook.Descendants<S.Sheet>())
        {
            if (sheet.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var sheetName = sheet.Name?.Value ?? string.Empty;
            foreach (var pivotPart in worksheetPart.PivotTableParts)
            {
                if (pivotPart.PivotTableDefinition?.Name?.Value is not { } pivotName)
                {
                    continue;
                }

                var worksheetSource = pivotPart.PivotTableCacheDefinitionPart
                    ?.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                result[(sheetName, pivotName)] = (worksheetSource?.Sheet?.Value, worksheetSource?.Reference?.Value);
            }
        }

        return result;
    }

    private sealed class SheetPivotComparer : IEqualityComparer<(string Sheet, string Pivot)>
    {
        public static readonly SheetPivotComparer Instance = new();

        public bool Equals((string Sheet, string Pivot) x, (string Sheet, string Pivot) y) =>
            string.Equals(x.Sheet, y.Sheet, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(x.Pivot, y.Pivot, StringComparison.OrdinalIgnoreCase);

        public int GetHashCode((string Sheet, string Pivot) obj) =>
            HashCode.Combine(
                obj.Sheet.ToUpperInvariant(),
                obj.Pivot.ToUpperInvariant());
    }

    // ----- agg mapping ---------------------------------------------------------

    /// <summary>
    /// Mirrors a value's show-values-as onto the live ClosedXML model so the
    /// in-memory pivot is self-consistent (the authoritative file write is the raw
    /// post-save <see cref="ApplyShowValuesAs"/> pass, because ClosedXML 0.105's
    /// writer drops these attributes). The (previous)/(next) base-item sentinels map
    /// to ClosedXML's AndPrevious()/AndNext(); any other literal goes through And().
    /// </summary>
    private static void ApplyInMemory(IXLPivotValue pivotValue, ValueSpec value)
    {
        switch (value.ShowAs)
        {
            case "normal":
                pivotValue.ShowAsNormal();
                break;
            case "percentOfTotal":
                pivotValue.ShowAsPercentageOfTotal();
                break;
            case "percentOfColumn":
                pivotValue.ShowAsPercentageOfColumn();
                break;
            case "percentOfRow":
                pivotValue.ShowAsPercentageOfRow();
                break;
            case "index":
                pivotValue.ShowAsIndex();
                break;
            case "runningTotal":
                pivotValue.ShowAsRunningTotalIn(value.BaseField!);
                break;
            case "differenceFrom":
                ApplyCombination(pivotValue.ShowAsDifferenceFrom(value.BaseField!), value.BaseItem!);
                break;
            case "percentDifferenceFrom":
                ApplyCombination(pivotValue.ShowAsPercentageDifferenceFrom(value.BaseField!), value.BaseItem!);
                break;
            case "percentOf":
                ApplyCombination(pivotValue.ShowAsPercentageFrom(value.BaseField!), value.BaseItem!);
                break;
            default:
                throw new InvalidOperationException($"showAs '{value.ShowAs}' escaped validation");
        }
    }

    private static void ApplyCombination(IXLPivotValueCombination combination, string baseItem)
    {
        switch (BaseItemSentinel(baseItem))
        {
            case BaseItemPrevious:
                combination.AndPrevious();
                break;
            case BaseItemNext:
                combination.AndNext();
                break;
            default:
                combination.And(baseItem);
                break;
        }
    }

    /// <summary>Maps the (previous)/(next) base-item sentinels to their OOXML index; a literal returns null.</summary>
    private static uint? BaseItemSentinel(string baseItem) => baseItem switch
    {
        "(previous)" => BaseItemPrevious,
        "(next)" => BaseItemNext,
        _ => null,
    };

    private static XLPivotSummary SummaryOf(string agg) => agg switch
    {
        "sum" => XLPivotSummary.Sum,
        "count" => XLPivotSummary.Count,
        "average" => XLPivotSummary.Average,
        "min" => XLPivotSummary.Minimum,
        "max" => XLPivotSummary.Maximum,
        _ => throw new InvalidOperationException($"agg '{agg}' escaped validation"),
    };

    private static string AggLabel(string agg) => agg switch
    {
        "sum" => "Sum",
        "count" => "Count",
        "average" => "Average",
        "min" => "Min",
        _ => "Max",
    };

    /// <summary>
    /// A unique ClosedXML custom name for a value field. The base is
    /// "<c>{Agg} of {field}</c>"; a same-field-same-agg repeat (allowed when it varies
    /// by showAs) gets a "<c> (2)</c>" suffix so ClosedXML's unique-name constraint holds.
    /// </summary>
    private static string UniqueValueName(ValueSpec value, HashSet<string> used)
    {
        var baseName = $"{AggLabel(value.Agg)} of {value.Field}";
        var candidate = baseName;
        for (var n = 2; !used.Add(candidate); n++)
        {
            candidate = string.Create(CultureInfo.InvariantCulture, $"{baseName} ({n})");
        }

        return candidate;
    }

    /// <summary>Wire name of a summary read back from a file (foreign aggs keep their enum name, camelCased).</summary>
    public static string AggName(XLPivotSummary summary) => summary switch
    {
        XLPivotSummary.Sum => "sum",
        XLPivotSummary.Count => "count",
        XLPivotSummary.Average => "average",
        XLPivotSummary.Minimum => "min",
        XLPivotSummary.Maximum => "max",
        _ => char.ToLowerInvariant(summary.ToString()[0]) + summary.ToString()[1..],
    };

    // ----- calculated fields (v1.3) ----------------------------------------------

    /// <summary>
    /// A validated calculated field for the post-save raw pass. <c>Name</c> is
    /// the new field's display name; <c>Formula</c> is its Excel expression over
    /// source field names (no leading <c>=</c>). <c>TargetSheet</c> / <c>PivotName</c>
    /// are filled in once the pivot is created so the writer can locate its parts.
    /// </summary>
    public sealed record CalculatedFieldSpec(string Name, string Formula)
    {
        public string TargetSheet { get; init; } = string.Empty;

        public string PivotName { get; init; } = string.Empty;
    }

    [GeneratedRegex(@"'(?<q>(?:[^']|'')+)'|(?<bare>[A-Za-z_\\][A-Za-z0-9_.]*)")]
    private static partial Regex FormulaTokenPattern();

    /// <summary>
    /// Parses and validates <c>calculatedFields</c>. Each entry is
    /// <c>{name, formula}</c>; the name must be non-empty and not collide with a
    /// source header or another calculated field; the formula's field references
    /// (barewords and 'quoted names') must all be source headers. A bad reference
    /// is <c>invalid_args</c> with the source field names as candidates.
    /// </summary>
    private static List<CalculatedFieldSpec> ParseCalculatedFields(
        JsonObject props, List<string> headers, int opIndex)
    {
        if (!props.TryGetPropertyValue("calculatedFields", out var node) || node is null)
        {
            return [];
        }

        if (node is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'calculatedFields' must be an array like [{{\"name\":\"Margin\",\"formula\":\"=Revenue-Cost\"}}].",
                "Each entry names the calculated field and gives an Excel formula over the source field names.");
        }

        // Function names that may appear in a formula but are not field refs.
        var functionNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "IF", "AND", "OR", "NOT", "SUM", "AVERAGE", "MIN", "MAX", "ABS", "ROUND",
            "INT", "MOD", "SQRT", "POWER", "TRUE", "FALSE",
        };

        var result = new List<CalculatedFieldSpec>(array.Count);
        var seenNames = new HashSet<string>(headers, StringComparer.OrdinalIgnoreCase);
        foreach (var entry in array)
        {
            if (entry is not JsonObject obj)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: calculatedFields entries must be objects like {{\"name\":\"Margin\",\"formula\":\"=Revenue-Cost\"}}.",
                    "Pass a name and an Excel formula over the source fields.");
            }

            var name = OptionalString(obj, "name");
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: a calculatedFields entry is missing its 'name'.",
                    "Pass {\"name\":\"Margin\",\"formula\":\"=Revenue-Cost\"}.");
            }

            if (!seenNames.Add(name))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: calculated field name '{name}' collides with a source field or another calculated field.",
                    "Give each calculated field a unique name distinct from the source headers.");
            }

            var formula = OptionalString(obj, "formula");
            if (string.IsNullOrWhiteSpace(formula))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: calculated field '{name}' is missing its 'formula'.",
                    "Pass an Excel formula over the source field names, e.g. {\"formula\":\"=Revenue-Cost\"}.");
            }

            foreach (var (key, _) in obj)
            {
                if (key is not ("name" or "formula"))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{opIndex}]: unknown calculatedFields field '{key}'.",
                        "calculatedFields entries accept {name, formula}.",
                        candidates: ["name", "formula"]);
                }
            }

            // Excel stores the formula WITHOUT the leading '='.
            var expression = formula.StartsWith('=') ? formula[1..] : formula;
            ValidateFormulaReferences(name, expression, headers, functionNames, opIndex);
            result.Add(new CalculatedFieldSpec(name, expression));
        }

        return result;
    }

    /// <summary>
    /// Ensures every field reference in a calculated-field formula is a source
    /// header. Identifiers immediately followed by '(' are treated as function
    /// calls, not field refs. An unknown reference is invalid_args with the real
    /// source field names as candidates.
    /// </summary>
    private static void ValidateFormulaReferences(
        string fieldName, string expression, List<string> headers, HashSet<string> functionNames, int opIndex)
    {
        foreach (Match match in FormulaTokenPattern().Matches(expression))
        {
            string identifier;
            if (match.Groups["q"].Success)
            {
                identifier = match.Groups["q"].Value.Replace("''", "'", StringComparison.Ordinal);
            }
            else
            {
                identifier = match.Groups["bare"].Value;

                // A bareword directly followed by '(' is a function call; skip it.
                var after = match.Index + match.Length;
                if (after < expression.Length && expression[after] == '(')
                {
                    continue;
                }

                if (functionNames.Contains(identifier))
                {
                    continue;
                }
            }

            if (!headers.Contains(identifier, StringComparer.OrdinalIgnoreCase))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: calculated field '{fieldName}' references '{identifier}', which is not a source field.",
                    "Calculated-field formulas reference the source header names; pick one of the candidates " +
                    "(wrap names with spaces in single quotes, e.g. 'Unit Price').",
                    candidates: [.. headers.OrderBy(h => ExcelPaths.Levenshtein(identifier, h))]);
            }
        }
    }

    // ----- parsing helpers ------------------------------------------------------

    private sealed record ValueSpec(string Field, string Agg)
    {
        /// <summary>The wire show-values-as mode (default <c>normal</c>); see <see cref="ShowAsModes"/>.</summary>
        public string ShowAs { get; init; } = "normal";

        /// <summary>The base field name for modes that need one (running total, difference-from, …); else null.</summary>
        public string? BaseField { get; init; }

        /// <summary>The base item for difference-from modes: a literal value, or the <c>(previous)</c>/<c>(next)</c> sentinels.</summary>
        public string? BaseItem { get; init; }
    }

    private static (int FirstColumn, int FirstRow, int LastColumn, int LastRow) ParseSourceRange(
        string text, int opIndex)
    {
        var match = RangePattern().Match(text.ToUpperInvariant());
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{text}' is not a usable sourceRange.",
                "sourceRange is a plain range on the op's sheet whose first row holds the field headers, e.g. A1:D20.");
        }

        var start = new CellRef(match.Groups[1].Value, int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture));
        var end = new CellRef(match.Groups[3].Value, int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture));
        if (start.ColumnNumber > end.ColumnNumber || start.Row > end.Row)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: sourceRange start must not be past its end: {text}",
                "Write the range top-left to bottom-right, e.g. A1:D20.");
        }

        return (start.ColumnNumber, start.Row, end.ColumnNumber, end.Row);
    }

    private static List<string> ReadHeaders(
        IXLWorksheet sheet, int firstColumn, int headerRow, int lastColumn, int opIndex)
    {
        var headers = new List<string>(lastColumn - firstColumn + 1);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var column = firstColumn; column <= lastColumn; column++)
        {
            var cell = sheet.Cell(headerRow, column);
            var text = ExcelValues.SafeFormatted(cell);
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: header cell {ExcelPaths.CellPath(sheet, cell.Address)} is empty; " +
                    "every sourceRange column needs a header.",
                    "Fill the first row of sourceRange with unique field names, or narrow the range.");
            }

            if (!seen.Add(text))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: duplicate header '{text}' in the sourceRange header row.",
                    "Pivot field names must be unique; rename the duplicate column header first.");
            }

            headers.Add(text);
        }

        return headers;
    }

    private static string ResolveField(List<string> headers, string requested, int opIndex)
    {
        foreach (var header in headers)
        {
            if (string.Equals(header, requested, StringComparison.OrdinalIgnoreCase))
            {
                return header;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: '{requested}' is not a field of the sourceRange.",
            "Field names come from the header row of sourceRange; pick one of the candidates.",
            candidates: [.. headers.OrderBy(h => ExcelPaths.Levenshtein(requested, h))]);
    }

    private static void GuardAxisOverlap(
        List<string> rows, List<string> columns, List<string> filters, int opIndex)
    {
        var axisOf = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (axis, fields) in new[] { ("rows", rows), ("columns", columns), ("filters", filters) })
        {
            foreach (var field in fields)
            {
                if (!axisOf.TryAdd(field, axis))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{opIndex}]: field '{field}' appears in both {axisOf[field]} and {axis}.",
                        "A field can sit on one axis only (rows, columns or filters); values may reuse any field.");
                }
            }
        }
    }

    private static void GuardDuplicateValues(List<ValueSpec> values, int opIndex)
    {
        // A field may appear once per (agg, showAs) combination — Excel happily shows
        // the same field as both a raw sum AND a % of total, so showAs is part of the
        // identity (an exact (field, agg, showAs) repeat is the real duplicate).
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var value in values)
        {
            if (!seen.Add(value.Field + "|" + value.Agg + "|" + value.ShowAs))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: values lists '{value.Field}' with agg '{value.Agg}' and showAs '{value.ShowAs}' twice.",
                    "Each (field, agg, showAs) combination may appear once; drop the duplicate, use a different agg, or vary the showAs.");
            }
        }
    }

    /// <summary>Resolved grand-total visibility: which of the row / column grand totals show.</summary>
    private readonly record struct GrandTotalsSpec(bool Rows, bool Columns);

    /// <summary>
    /// Parses the <c>grandTotals</c> prop, controlling row/column grand-total visibility.
    /// Accepts EITHER a string <c>both</c>|<c>rows</c>|<c>columns</c>|<c>none</c> OR an
    /// object <c>{rows:bool, columns:bool}</c>. Returns null when the prop is absent (the
    /// caller then leaves ClosedXML's always-both default untouched, byte-for-byte). A bad
    /// string is <c>invalid_args</c> with <see cref="GrandTotalModes"/> as candidates; a
    /// non-boolean <c>rows</c>/<c>columns</c> (or an unknown sub-key) is also <c>invalid_args</c>.
    /// </summary>
    private static GrandTotalsSpec? ParseGrandTotals(JsonObject props, int opIndex)
    {
        if (!props.TryGetPropertyValue("grandTotals", out var node) || node is null)
        {
            return null;
        }

        switch (node)
        {
            case JsonValue value when value.GetValueKind() == JsonValueKind.String:
            {
                var mode = value.GetValue<string>();
                return mode switch
                {
                    "both" => new GrandTotalsSpec(true, true),
                    "rows" => new GrandTotalsSpec(true, false),
                    "columns" => new GrandTotalsSpec(false, true),
                    "none" => new GrandTotalsSpec(false, false),
                    _ => throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{opIndex}]: unknown grandTotals '{mode}'.",
                        "grandTotals is one of " + string.Join(", ", GrandTotalModes) +
                        ", or an object like {\"rows\":true,\"columns\":false}.",
                        candidates: GrandTotalModes),
                };
            }

            case JsonObject obj:
            {
                foreach (var (key, _) in obj)
                {
                    if (key is not ("rows" or "columns"))
                    {
                        throw new AiofficeException(
                            ErrorCodes.InvalidArgs,
                            $"ops[{opIndex}]: unknown grandTotals field '{key}'.",
                            "The grandTotals object accepts {rows, columns} booleans.",
                            candidates: ["rows", "columns"]);
                    }
                }

                // Omitting a sub-flag keeps that axis' grand total shown (today's default).
                return new GrandTotalsSpec(
                    ReadGrandTotalFlag(obj, "rows", opIndex),
                    ReadGrandTotalFlag(obj, "columns", opIndex));
            }

            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: 'grandTotals' must be a string ({string.Join("|", GrandTotalModes)}) " +
                    "or an object {rows:bool, columns:bool}.",
                    "E.g. {\"grandTotals\":\"none\"} or {\"grandTotals\":{\"rows\":false,\"columns\":true}}.");
        }
    }

    /// <summary>Reads a boolean grandTotals sub-flag; absent => true (axis stays shown); non-bool => invalid_args.</summary>
    private static bool ReadGrandTotalFlag(JsonObject obj, string key, int opIndex)
    {
        if (!obj.TryGetPropertyValue(key, out var node) || node is null)
        {
            return true;
        }

        if (node is not JsonValue value ||
            value.GetValueKind() is not (JsonValueKind.True or JsonValueKind.False))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: grandTotals.{key} must be a boolean (true or false).",
                "Pass e.g. {\"grandTotals\":{\"rows\":false,\"columns\":true}}.");
        }

        return value.GetValue<bool>();
    }

    private static List<ValueSpec> ParseValues(JsonObject props, int opIndex)
    {
        if (!props.TryGetPropertyValue("values", out var node) || node is null)
        {
            return [];
        }

        if (node is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'values' must be an array like [{{\"field\":\"Sales\",\"agg\":\"sum\"}}].",
                "Each entry names a field from the sourceRange header row plus an agg (sum|count|average|min|max).");
        }

        var result = new List<ValueSpec>(array.Count);
        foreach (var entry in array)
        {
            switch (entry)
            {
                case JsonValue value when value.GetValueKind() == JsonValueKind.String:
                    result.Add(new ValueSpec(value.GetValue<string>(), "sum"));
                    break;

                case JsonObject obj:
                {
                    foreach (var (key, _) in obj)
                    {
                        if (!ValueProps.Contains(key, StringComparer.Ordinal))
                        {
                            throw new AiofficeException(
                                ErrorCodes.InvalidArgs,
                                $"ops[{opIndex}]: unknown values prop '{key}'.",
                                "values entries accept " + string.Join(", ", ValueProps) + ".",
                                candidates: ValueProps);
                        }
                    }

                    var field = obj.TryGetPropertyValue("field", out var f) && f is JsonValue fv &&
                        fv.GetValueKind() == JsonValueKind.String
                            ? fv.GetValue<string>()
                            : throw new AiofficeException(
                                ErrorCodes.InvalidArgs,
                                $"ops[{opIndex}]: a values entry is missing its 'field'.",
                                "Pass {\"field\":\"Sales\",\"agg\":\"sum\"}; agg defaults to sum.");
                    var agg = obj.TryGetPropertyValue("agg", out var a) && a is JsonValue av &&
                        av.GetValueKind() == JsonValueKind.String
                            ? av.GetValue<string>()
                            : "sum";
                    if (!Aggs.Contains(agg, StringComparer.Ordinal))
                    {
                        throw new AiofficeException(
                            ErrorCodes.InvalidArgs,
                            $"ops[{opIndex}]: unknown agg '{agg}'.",
                            "Supported aggs: " + string.Join(", ", Aggs) + ".",
                            candidates: Aggs);
                    }

                    result.Add(ParseShowAs(obj, new ValueSpec(field, agg), opIndex));
                    break;
                }

                default:
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{opIndex}]: values entries must be objects like {{\"field\":\"Sales\",\"agg\":\"sum\"}}.",
                        "A bare field-name string is also accepted (agg defaults to sum).");
            }
        }

        return result;
    }

    /// <summary>
    /// Validates and folds the show-values-as props (<c>showAs</c>/<c>baseField</c>/
    /// <c>baseItem</c>) of one value entry onto its <see cref="ValueSpec"/>.
    /// An unknown <c>showAs</c> is <c>invalid_args</c> with <see cref="ShowAsModes"/>
    /// as candidates; a mode OOXML cannot carry is <c>unsupported_feature</c>; modes
    /// needing a base field/item that lack one are <c>invalid_args</c>. The base field
    /// name itself is resolved against the headers later (in <see cref="Add"/>) so the
    /// candidate list is the real headers.
    /// </summary>
    private static ValueSpec ParseShowAs(JsonObject obj, ValueSpec spec, int opIndex)
    {
        var showAs = OptionalString(obj, "showAs");
        var hasBaseField = obj.TryGetPropertyValue("baseField", out var bf) && bf is not null;
        var hasBaseItem = obj.TryGetPropertyValue("baseItem", out var bi) && bi is not null;

        if (string.IsNullOrWhiteSpace(showAs))
        {
            if (hasBaseField || hasBaseItem)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: '{spec.Field}' sets baseField/baseItem without a showAs.",
                    "baseField/baseItem only apply to a showAs like runningTotal or differenceFrom; add a showAs or drop them.");
            }

            return spec; // default normal
        }

        if (!ShowAsModes.Contains(showAs, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: unknown showAs '{showAs}' on value field '{spec.Field}'.",
                "Supported showAs modes: " + string.Join(", ", ShowAsModes) + ".",
                candidates: ShowAsModes);
        }

        if (UnsupportedShowAs.TryGetValue(showAs, out var why))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{opIndex}]: showAs '{showAs}' on '{spec.Field}' cannot be authored. {why}",
                "Use a supported mode: " + string.Join(", ", ShowAsModes.Where(m => !UnsupportedShowAs.ContainsKey(m))) +
                " — e.g. percentOfTotal for a flat share, or runningTotal with a baseField.",
                candidates: [.. ShowAsModes.Where(m => !UnsupportedShowAs.ContainsKey(m))]);
        }

        var baseField = OptionalString(obj, "baseField");
        var baseItem = ReadBaseItem(obj, spec.Field, opIndex);

        if (NeedsBaseField.Contains(showAs) && string.IsNullOrWhiteSpace(baseField))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: showAs '{showAs}' on '{spec.Field}' needs a baseField.",
                "Pass the field to run/compare against, e.g. {\"showAs\":\"" + showAs + "\",\"baseField\":\"Date\"" +
                (NeedsBaseItem.Contains(showAs) ? ",\"baseItem\":\"(previous)\"" : string.Empty) + "}.");
        }

        if (!NeedsBaseField.Contains(showAs) && !string.IsNullOrWhiteSpace(baseField))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: showAs '{showAs}' on '{spec.Field}' does not take a baseField.",
                "Only runningTotal / differenceFrom / percentDifferenceFrom / percentOf use a baseField; drop it.");
        }

        if (NeedsBaseItem.Contains(showAs) && string.IsNullOrWhiteSpace(baseItem))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: showAs '{showAs}' on '{spec.Field}' needs a baseItem.",
                "Pass the comparison item: a literal value from the baseField, or the \"(previous)\"/\"(next)\" sentinel.");
        }

        if (!NeedsBaseItem.Contains(showAs) && hasBaseItem)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: showAs '{showAs}' on '{spec.Field}' does not take a baseItem.",
                "Only differenceFrom / percentDifferenceFrom / percentOf use a baseItem; drop it.");
        }

        return spec with { ShowAs = showAs, BaseField = baseField, BaseItem = baseItem };
    }

    /// <summary>Reads a <c>baseItem</c> as a string (numbers and bools stringify; the (previous)/(next) sentinels pass through).</summary>
    private static string? ReadBaseItem(JsonObject obj, string field, int opIndex)
    {
        if (!obj.TryGetPropertyValue("baseItem", out var node) || node is null)
        {
            return null;
        }

        if (node is not JsonValue value)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: baseItem on '{field}' must be a value (a string, number, or \"(previous)\"/\"(next)\").",
                "Pass e.g. {\"baseItem\":\"2023\"} or {\"baseItem\":\"(previous)\"}.");
        }

        return value.GetValueKind() switch
        {
            JsonValueKind.String => value.GetValue<string>(),
            JsonValueKind.Number => value.GetValue<double>().ToString(CultureInfo.InvariantCulture),
            JsonValueKind.True => "TRUE",
            JsonValueKind.False => "FALSE",
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: baseItem on '{field}' must be a string, number, or boolean.",
                "Pass the literal item value from the baseField, or \"(previous)\"/\"(next)\"."),
        };
    }

    private static List<string> StringList(JsonObject props, string key, int opIndex)
    {
        if (!props.TryGetPropertyValue(key, out var node) || node is null)
        {
            return [];
        }

        if (node is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{key}' must be an array of field names.",
                $"Pass e.g. {{\"{key}\":[\"Region\"]}}; names come from the sourceRange header row.");
        }

        var result = new List<string>(array.Count);
        foreach (var entry in array)
        {
            if (entry is not JsonValue value || value.GetValueKind() != JsonValueKind.String)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: '{key}' entries must be strings (field names).",
                    $"Pass e.g. {{\"{key}\":[\"Region\",\"Product\"]}}.");
            }

            result.Add(value.GetValue<string>());
        }

        return result;
    }

    private static string DefaultName(IXLWorksheet targetSheet)
    {
        for (var i = 1; ; i++)
        {
            var candidate = string.Create(CultureInfo.InvariantCulture, $"PivotTable{i}");
            if (!targetSheet.PivotTables.Contains(candidate))
            {
                return candidate;
            }
        }
    }

    private static IXLWorksheet AddSheet(XLWorkbook workbook, string name, int opIndex)
    {
        try
        {
            return workbook.AddWorksheet(name);
        }
        catch (ArgumentException exception)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{name}' is not a usable targetSheet name: {exception.Message}",
                @"Sheet names are 1-31 characters and cannot contain : \ / ? * [ ].",
                innerException: exception);
        }
    }

    private static string RequiredString(JsonObject props, string key, int opIndex, string example)
    {
        if (props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
            value.GetValueKind() == JsonValueKind.String &&
            value.GetValue<string>() is { Length: > 0 } text)
        {
            return text;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add pivot needs the '{key}' prop.",
            $"Pass it as a string, e.g. {{\"{key}\":\"{example}\"}}.");
    }

    private static string? OptionalString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;
}
