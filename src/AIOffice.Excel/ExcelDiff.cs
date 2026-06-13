using System.Globalization;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The semantic workbook compare behind <c>aioffice diff</c> (M8). Two workbooks
/// are compared structure-first, then cell-by-cell over the union of each common
/// sheet's used range, plus defined-name and table (ListObject) add/remove. The
/// change list is returned through <see cref="DiffResult.FromChanges"/>, which
/// sorts it deterministically by (path, kind) so the output is byte-identical on
/// every platform.
///
/// <para>Semantics:</para>
/// <list type="bullet">
/// <item>A sheet present in only one workbook is one <c>added</c>/<c>removed</c>
/// change at the sheet path; its cells are NOT enumerated (the sheet itself is
/// the change).</item>
/// <item>For a common sheet, each cell in the union of the two used ranges is
/// compared: empty→set is <c>added</c>, set→empty is <c>removed</c>, and a
/// value/formula change is <c>modified</c> with concise Before/After. A change
/// in ONLY the number format is <c>modified</c> with Detail <c>"format"</c>.</item>
/// <item>Cell changes are capped (<see cref="MaxCellChanges"/>); when the cap is
/// hit the result carries a <c>diff_truncated</c> warning and stops emitting
/// further cell changes (sheet/name/table changes are always reported).</item>
/// </list>
/// </summary>
internal static class ExcelDiff
{
    /// <summary>The most per-pass cell changes reported before truncation kicks in.</summary>
    public const int MaxCellChanges = 500;

    public static DiffResult Compare(string baselineFile, string currentFile)
    {
        using var baseline = OpenWorkbook(baselineFile);
        using var current = OpenWorkbook(currentFile);

        var changes = new List<DiffChange>();
        var cellChangeCount = 0;
        var truncated = false;

        DiffSheets(baseline, current, changes, ref cellChangeCount, ref truncated);
        DiffDefinedNames(baseline, current, changes);

        IReadOnlyList<Warning>? warnings = truncated
            ? [new Warning(
                "diff_truncated",
                $"More than {MaxCellChanges} cell changes; reporting the first {MaxCellChanges}. " +
                "Sheet, defined-name and table changes are still complete.")]
            : null;

        return DiffResult.FromChanges(changes, warnings);
    }

    // ----- sheets + cells -----------------------------------------------------

    private static void DiffSheets(
        XLWorkbook baseline, XLWorkbook current, List<DiffChange> changes,
        ref int cellChangeCount, ref bool truncated)
    {
        var baselineSheets = baseline.Worksheets.ToDictionary(ws => ws.Name, StringComparer.OrdinalIgnoreCase);
        var currentSheets = current.Worksheets.ToDictionary(ws => ws.Name, StringComparer.OrdinalIgnoreCase);

        foreach (var (name, ws) in currentSheets)
        {
            if (!baselineSheets.ContainsKey(name))
            {
                changes.Add(new DiffChange
                {
                    Kind = "added",
                    Path = ExcelPaths.SheetPath(ws),
                    Detail = "sheet",
                });
            }
        }

        foreach (var (name, ws) in baselineSheets)
        {
            if (!currentSheets.ContainsKey(name))
            {
                changes.Add(new DiffChange
                {
                    Kind = "removed",
                    Path = ExcelPaths.SheetPath(ws),
                    Detail = "sheet",
                });
            }
        }

        foreach (var (name, baselineSheet) in baselineSheets)
        {
            if (currentSheets.TryGetValue(name, out var currentSheet))
            {
                DiffCells(baselineSheet, currentSheet, changes, ref cellChangeCount, ref truncated);
                DiffTables(baselineSheet, currentSheet, changes);
            }
        }
    }

    private static void DiffCells(
        IXLWorksheet baselineSheet, IXLWorksheet currentSheet, List<DiffChange> changes,
        ref int cellChangeCount, ref bool truncated)
    {
        var baselineUsed = baselineSheet.RangeUsed()?.RangeAddress;
        var currentUsed = currentSheet.RangeUsed()?.RangeAddress;
        if (baselineUsed is null && currentUsed is null)
        {
            return; // both sheets empty
        }

        var firstRow = Math.Min(
            baselineUsed?.FirstAddress.RowNumber ?? int.MaxValue,
            currentUsed?.FirstAddress.RowNumber ?? int.MaxValue);
        var lastRow = Math.Max(
            baselineUsed?.LastAddress.RowNumber ?? 0,
            currentUsed?.LastAddress.RowNumber ?? 0);
        var firstColumn = Math.Min(
            baselineUsed?.FirstAddress.ColumnNumber ?? int.MaxValue,
            currentUsed?.FirstAddress.ColumnNumber ?? int.MaxValue);
        var lastColumn = Math.Max(
            baselineUsed?.LastAddress.ColumnNumber ?? 0,
            currentUsed?.LastAddress.ColumnNumber ?? 0);

        // Deterministic iteration order: row-major over the union rectangle.
        for (var row = firstRow; row <= lastRow; row++)
        {
            for (var column = firstColumn; column <= lastColumn; column++)
            {
                if (truncated)
                {
                    return;
                }

                var baselineCell = baselineSheet.Cell(row, column);
                var currentCell = currentSheet.Cell(row, column);
                var change = CompareCell(currentSheet, baselineCell, currentCell);
                if (change is null)
                {
                    continue;
                }

                if (cellChangeCount >= MaxCellChanges)
                {
                    truncated = true;
                    return;
                }

                changes.Add(change);
                cellChangeCount++;
            }
        }
    }

    /// <summary>Compares one cell address across the two workbooks; null when equal.</summary>
    private static DiffChange? CompareCell(IXLWorksheet currentSheet, IXLCell baselineCell, IXLCell currentCell)
    {
        var baselineContent = CellContent(baselineCell);
        var currentContent = CellContent(currentCell);
        var path = ExcelPaths.CellPath(currentSheet, currentCell.Address);

        var baselineEmpty = baselineContent is null;
        var currentEmpty = currentContent is null;

        if (baselineEmpty && currentEmpty)
        {
            // Both blank in value/formula terms: a number-format-only change is
            // still worth reporting on a non-empty-format cell, but a pair of
            // truly empty cells is not. Compare formats only when one side has a
            // value somewhere — here both are empty, so nothing to report.
            return null;
        }

        if (baselineEmpty)
        {
            return new DiffChange { Kind = "added", Path = path, After = currentContent };
        }

        if (currentEmpty)
        {
            // The cell no longer holds content; report at the (still-valid) path.
            return new DiffChange { Kind = "removed", Path = path, Before = baselineContent };
        }

        if (!string.Equals(baselineContent, currentContent, StringComparison.Ordinal))
        {
            return new DiffChange
            {
                Kind = "modified",
                Path = path,
                Before = baselineContent,
                After = currentContent,
            };
        }

        // Same content: a number-format-only difference is a "format" modify.
        var baselineFormat = NormalizeFormat(baselineCell.Style.NumberFormat.Format);
        var currentFormat = NormalizeFormat(currentCell.Style.NumberFormat.Format);
        if (!string.Equals(baselineFormat, currentFormat, StringComparison.Ordinal))
        {
            return new DiffChange
            {
                Kind = "modified",
                Path = path,
                Before = baselineFormat.Length == 0 ? "General" : baselineFormat,
                After = currentFormat.Length == 0 ? "General" : currentFormat,
                Detail = "format",
            };
        }

        return null;
    }

    /// <summary>
    /// The comparable content of a cell: its formula (with a leading <c>=</c>) or
    /// its typed value as canonical text. Null when the cell holds neither (a
    /// truly empty cell). Number formatting is compared separately.
    /// </summary>
    private static string? CellContent(IXLCell cell)
    {
        if (cell.HasFormula)
        {
            return "=" + cell.FormulaA1;
        }

        var value = cell.Value;
        return value.Type == XLDataType.Blank ? null : ValueText(value);
    }

    /// <summary>A culture-invariant, platform-stable text form of a cell value.</summary>
    private static string ValueText(XLCellValue value) => value.Type switch
    {
        XLDataType.Blank => string.Empty,
        XLDataType.Boolean => value.GetBoolean() ? "true" : "false",
        XLDataType.Number => value.GetNumber().ToString("R", CultureInfo.InvariantCulture),
        XLDataType.Text => value.GetText(),
        XLDataType.Error => value.GetError().ToString(),
        XLDataType.DateTime => value.GetDateTime().ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture),
        XLDataType.TimeSpan => value.GetTimeSpan().ToString("c", CultureInfo.InvariantCulture),
        _ => value.ToString(CultureInfo.InvariantCulture),
    };

    private static string NormalizeFormat(string? format) => format ?? string.Empty;

    // ----- tables (ListObjects) -----------------------------------------------

    private static void DiffTables(
        IXLWorksheet baselineSheet, IXLWorksheet currentSheet, List<DiffChange> changes)
    {
        var baselineTables = baselineSheet.Tables
            .ToDictionary(t => t.Name, StringComparer.OrdinalIgnoreCase);
        var currentTables = currentSheet.Tables
            .ToDictionary(t => t.Name, StringComparer.OrdinalIgnoreCase);

        foreach (var (name, table) in currentTables)
        {
            if (!baselineTables.ContainsKey(name))
            {
                changes.Add(new DiffChange
                {
                    Kind = "added",
                    Path = ExcelPaths.TablePath(currentSheet, name),
                    Detail = "table",
                    After = table.RangeAddress.ToString(),
                });
            }
        }

        foreach (var (name, table) in baselineTables)
        {
            if (!currentTables.ContainsKey(name))
            {
                changes.Add(new DiffChange
                {
                    Kind = "removed",
                    Path = ExcelPaths.TablePath(baselineSheet, name),
                    Detail = "table",
                    Before = table.RangeAddress.ToString(),
                });
            }
        }
    }

    // ----- defined names ------------------------------------------------------

    private static void DiffDefinedNames(XLWorkbook baseline, XLWorkbook current, List<DiffChange> changes)
    {
        var baselineNames = NamesOf(baseline);
        var currentNames = NamesOf(current);

        foreach (var (name, reference) in currentNames)
        {
            if (!baselineNames.TryGetValue(name, out var baselineRef))
            {
                changes.Add(new DiffChange
                {
                    Kind = "added",
                    Path = "/definedName[@name=" + name + "]",
                    Detail = "definedName",
                    After = reference,
                });
            }
            else if (!string.Equals(baselineRef, reference, StringComparison.Ordinal))
            {
                changes.Add(new DiffChange
                {
                    Kind = "modified",
                    Path = "/definedName[@name=" + name + "]",
                    Detail = "definedName",
                    Before = baselineRef,
                    After = reference,
                });
            }
        }

        foreach (var (name, reference) in baselineNames)
        {
            if (!currentNames.ContainsKey(name))
            {
                changes.Add(new DiffChange
                {
                    Kind = "removed",
                    Path = "/definedName[@name=" + name + "]",
                    Detail = "definedName",
                    Before = reference,
                });
            }
        }
    }

    /// <summary>
    /// Every USER defined name (workbook + sheet scope) keyed by its qualified
    /// name → reference text. Builtin bookkeeping names (<c>_xlnm.Print_Area</c>,
    /// <c>_xlnm._FilterDatabase</c>, …) are skipped: they are surfaced through
    /// their own features (print area, autofilter), so diffing them here would be
    /// noise that duplicates a sheet-property change.
    /// </summary>
    private static Dictionary<string, string> NamesOf(XLWorkbook workbook)
    {
        var names = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var name in workbook.DefinedNames.Where(n => !IsBuiltin(n)))
        {
            names[name.Name] = name.RefersTo;
        }

        foreach (var sheet in workbook.Worksheets)
        {
            foreach (var name in sheet.DefinedNames.Where(n => !IsBuiltin(n)))
            {
                // Sheet-scoped names are qualified so a workbook-scoped name of
                // the same text does not collide with a sheet-scoped one.
                names[sheet.Name + "!" + name.Name] = name.RefersTo;
            }
        }

        return names;
    }

    private static bool IsBuiltin(IXLDefinedName name) =>
        name.Name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase);

    private static XLWorkbook OpenWorkbook(string file)
    {
        try
        {
            return new XLWorkbook(file);
        }
        catch (Exception exception)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                $"The file could not be opened as an xlsx workbook: {exception.Message}",
                "Verify both files are real .xlsx workbooks (zip containers); re-export from the source application.",
                innerException: exception);
        }
    }
}
