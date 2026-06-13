using System.Globalization;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// The xlsx accessibility/quality audit (M7 flagship slice). Pure detection
/// lives here; <see cref="ExcelHandler"/> implements <see cref="IAuditor"/> and
/// owns the fix pass (so fixes ride the same atomic snapshot/save pipeline as
/// every other edit).
///
/// <para>Checks (see the M7 contract for the code namespace):</para>
/// <list type="bullet">
/// <item>quality_formula_error — a cell's cached value is an Excel error
/// (#DIV/0!, #VALUE!, #N/A, #NAME?, #NULL!, #NUM!, #REF!). error.</item>
/// <item>quality_broken_ref — a formula's text contains #REF! (a deleted
/// range). error.</item>
/// <item>quality_duplicate_id — the same defined-name appears in more than one
/// scope, or a table name repeats across sheets. error.</item>
/// <item>a11y_no_alt_text — an image or chart has no alt-text / title.
/// warning, autofixable.</item>
/// <item>a11y_merged_data_cells — a merged range sits inside a used data
/// region (screen-reader hostile). warning.</item>
/// <item>a11y_no_doc_title — the core Title property is empty. warning,
/// autofixable.</item>
/// <item>a11y_low_contrast — a cell with an explicit fill has a font/fill
/// contrast below WCAG AA's 4.5:1. warning.</item>
/// <item>quality_hidden_data — a hidden row/column/sheet still carries data.
/// info.</item>
/// </list>
/// </summary>
internal static class ExcelAudit
{
    private const string CategoryAccessibility = "accessibility";
    private const string CategoryQuality = "quality";
    private const string SeverityError = "error";
    private const string SeverityWarning = "warning";
    private const string SeverityInfo = "info";

    /// <summary>Runs the audit and returns findings ordered (errors first), severity-filtered.</summary>
    public static AuditResult Run(string file, AuditOptions opts)
    {
        var findings = new List<AuditFinding>();
        var wantAccessibility = opts.Category is "all" or CategoryAccessibility;
        var wantQuality = opts.Category is "all" or CategoryQuality;

        using (var workbook = ExcelHandler.OpenWorkbookForAudit(file))
        {
            if (wantQuality)
            {
                FormulaErrors(workbook, findings);
                BrokenRefs(workbook, findings);
                DuplicateIds(workbook, findings);
                HiddenData(workbook, findings);
            }

            if (wantAccessibility)
            {
                NoDocTitle(workbook, file, findings);
                MergedDataCells(workbook, findings);
                LowContrast(workbook, findings);
            }
        }

        // Alt-text needs the raw drawing parts (ClosedXML cannot see chart
        // titles/descr); read them in one separate package pass.
        if (wantAccessibility)
        {
            NoAltText(file, findings);
        }

        var minRank = AuditOptions.SeverityRank(opts.MinSeverity);
        var filtered = findings
            .Where(f => AuditOptions.SeverityRank(f.Severity) >= minRank)
            .OrderByDescending(f => AuditOptions.SeverityRank(f.Severity))
            .ThenBy(f => f.Code, StringComparer.Ordinal)
            .ThenBy(f => f.Path, StringComparer.Ordinal)
            .ToList();

        return new AuditResult
        {
            Findings = filtered,
            Summary = new AuditSummary(
                filtered.Count(f => f.Severity == SeverityError),
                filtered.Count(f => f.Severity == SeverityWarning),
                filtered.Count(f => f.Severity == SeverityInfo)),
        };
    }

    // ----- quality ------------------------------------------------------------

    private static void FormulaErrors(XLWorkbook workbook, List<AuditFinding> findings)
    {
        foreach (var sheet in workbook.Worksheets)
        {
            foreach (var cell in sheet.CellsUsed())
            {
                XLCellValue value;
                try
                {
                    value = cell.HasFormula ? cell.CachedValue : cell.Value;
                }
                catch (Exception)
                {
                    continue;
                }

                if (value.Type != XLDataType.Error)
                {
                    continue;
                }

                var path = ExcelPaths.CellPath(sheet, cell.Address);
                findings.Add(new AuditFinding
                {
                    Id = $"quality_formula_error#{path}",
                    Severity = SeverityError,
                    Category = CategoryQuality,
                    Code = "quality_formula_error",
                    Path = path,
                    Message = $"Cell {path} evaluates to {value.ToString(CultureInfo.InvariantCulture)}.",
                    Suggestion = "Fix the formula or the data it depends on; aioffice does not auto-correct " +
                                 "formula errors because the right value is unknowable.",
                    Autofixable = false,
                });
            }
        }
    }

    private static void BrokenRefs(XLWorkbook workbook, List<AuditFinding> findings)
    {
        foreach (var sheet in workbook.Worksheets)
        {
            foreach (var cell in sheet.CellsUsed(XLCellsUsedOptions.AllContents).Where(c => c.HasFormula))
            {
                var formula = cell.FormulaA1;
                if (formula is null || !formula.Contains("#REF!", StringComparison.Ordinal))
                {
                    continue;
                }

                var path = ExcelPaths.CellPath(sheet, cell.Address);
                findings.Add(new AuditFinding
                {
                    Id = $"quality_broken_ref#{path}",
                    Severity = SeverityError,
                    Category = CategoryQuality,
                    Code = "quality_broken_ref",
                    Path = path,
                    Message = $"Formula in {path} references a deleted range (=#REF!): ={formula}.",
                    Suggestion = "Repoint the formula at a live range; a #REF! means the cells it pointed at were deleted.",
                    Autofixable = false,
                });
            }
        }
    }

    private static void DuplicateIds(XLWorkbook workbook, List<AuditFinding> findings)
    {
        // Defined names that collide across scopes (workbook + a sheet share the
        // same name, an easy source of "which one wins?" confusion).
        var nameScopes = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in workbook.DefinedNames.Where(n => !IsBuiltin(n.Name)))
        {
            Track(nameScopes, name.Name, "workbook");
        }

        foreach (var sheet in workbook.Worksheets)
        {
            foreach (var name in sheet.DefinedNames.Where(n => !IsBuiltin(n.Name)))
            {
                Track(nameScopes, name.Name, sheet.Name);
            }
        }

        foreach (var (name, scopes) in nameScopes.Where(kv => kv.Value.Count > 1))
        {
            var path = ExcelNames.PathOf(null, name);
            findings.Add(new AuditFinding
            {
                Id = $"quality_duplicate_id#{path}",
                Severity = SeverityError,
                Category = CategoryQuality,
                Code = "quality_duplicate_id",
                Path = path,
                Message = $"Defined name '{name}' exists in {scopes.Count} scopes ({string.Join(", ", scopes)}).",
                Suggestion = "Rename or remove the duplicate so each name resolves unambiguously.",
                Autofixable = false,
            });
        }

        // Table names must be unique workbook-wide; flag any repeat.
        var tableSheets = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in workbook.Worksheets)
        {
            foreach (var table in sheet.Tables)
            {
                Track(tableSheets, table.Name, sheet.Name);
            }
        }

        foreach (var (name, sheets) in tableSheets.Where(kv => kv.Value.Count > 1))
        {
            findings.Add(new AuditFinding
            {
                Id = $"quality_duplicate_id#table:{name}",
                Severity = SeverityError,
                Category = CategoryQuality,
                Code = "quality_duplicate_id",
                Path = null,
                Message = $"Table name '{name}' is used on {sheets.Count} sheets ({string.Join(", ", sheets)}).",
                Suggestion = "Table names must be unique in a workbook; rename the duplicates.",
                Autofixable = false,
            });
        }
    }

    private static void HiddenData(XLWorkbook workbook, List<AuditFinding> findings)
    {
        foreach (var sheet in workbook.Worksheets)
        {
            var sheetPath = ExcelPaths.SheetPath(sheet);

            if (sheet.Visibility != XLWorksheetVisibility.Visible && sheet.CellsUsed().Any())
            {
                findings.Add(new AuditFinding
                {
                    Id = $"quality_hidden_data#{sheetPath}",
                    Severity = SeverityInfo,
                    Category = CategoryQuality,
                    Code = "quality_hidden_data",
                    Path = sheetPath,
                    Message = $"Hidden sheet '{sheet.Name}' still contains data.",
                    Suggestion = "Confirm the hidden sheet is intentional; readers and screen readers will miss its data.",
                    Autofixable = false,
                });
                continue; // a hidden sheet's hidden rows/cols add nothing actionable
            }

            foreach (var row in sheet.RowsUsed().Where(r => r.IsHidden))
            {
                var path = ExcelPaths.RowPath(sheet, row.RowNumber());
                findings.Add(new AuditFinding
                {
                    Id = $"quality_hidden_data#{path}",
                    Severity = SeverityInfo,
                    Category = CategoryQuality,
                    Code = "quality_hidden_data",
                    Path = path,
                    Message = $"Hidden row {row.RowNumber()} on '{sheet.Name}' contains data.",
                    Suggestion = "Confirm the hidden row is intentional; its data is invisible to readers.",
                    Autofixable = false,
                });
            }

            foreach (var column in sheet.ColumnsUsed().Where(c => c.IsHidden))
            {
                var path = ExcelPaths.ColumnPath(sheet, column.ColumnNumber());
                findings.Add(new AuditFinding
                {
                    Id = $"quality_hidden_data#{path}",
                    Severity = SeverityInfo,
                    Category = CategoryQuality,
                    Code = "quality_hidden_data",
                    Path = path,
                    Message = $"Hidden column {column.ColumnLetter()} on '{sheet.Name}' contains data.",
                    Suggestion = "Confirm the hidden column is intentional; its data is invisible to readers.",
                    Autofixable = false,
                });
            }
        }
    }

    // ----- accessibility ------------------------------------------------------

    private static void NoDocTitle(XLWorkbook workbook, string file, List<AuditFinding> findings)
    {
        if (!string.IsNullOrWhiteSpace(workbook.Properties.Title))
        {
            return;
        }

        findings.Add(new AuditFinding
        {
            Id = "a11y_no_doc_title#/properties",
            Severity = SeverityWarning,
            Category = CategoryAccessibility,
            Code = "a11y_no_doc_title",
            Path = "/properties",
            Message = "The workbook has no document Title.",
            Suggestion = "Set a Title so assistive technology can announce the document: " +
                         "{op:set, path:/properties, props:{title:\"…\"}}. --fix derives one from the filename.",
            Autofixable = true,
        });
    }

    private static void MergedDataCells(XLWorkbook workbook, List<AuditFinding> findings)
    {
        foreach (var sheet in workbook.Worksheets)
        {
            var used = sheet.RangeUsed();
            if (used is null)
            {
                continue;
            }

            var usedAddress = used.RangeAddress;
            foreach (var merged in sheet.MergedRanges)
            {
                var m = merged.RangeAddress;

                // A merge that sits ENTIRELY inside the used data region (not a
                // banner row above/beside it) is the hostile case.
                var insideRows = m.FirstAddress.RowNumber >= usedAddress.FirstAddress.RowNumber &&
                                 m.LastAddress.RowNumber <= usedAddress.LastAddress.RowNumber;
                var insideColumns = m.FirstAddress.ColumnNumber >= usedAddress.FirstAddress.ColumnNumber &&
                                    m.LastAddress.ColumnNumber <= usedAddress.LastAddress.ColumnNumber;
                if (!insideRows || !insideColumns)
                {
                    continue;
                }

                var path = ExcelPaths.RangePath(sheet, m);
                findings.Add(new AuditFinding
                {
                    Id = $"a11y_merged_data_cells#{path}",
                    Severity = SeverityWarning,
                    Category = CategoryAccessibility,
                    Code = "a11y_merged_data_cells",
                    Path = path,
                    Message = $"Merged range {path} sits inside the data region {usedAddress}.",
                    Suggestion = "Unmerge data cells (merge breaks navigation and value reading for screen readers): " +
                                 "{op:set, path:" + path + ", props:{merge:false}}. Use 'center across selection' for banners instead.",
                    Autofixable = false,
                });
            }
        }
    }

    private static void LowContrast(XLWorkbook workbook, List<AuditFinding> findings)
    {
        foreach (var sheet in workbook.Worksheets)
        {
            foreach (var cell in sheet.CellsUsed(XLCellsUsedOptions.All))
            {
                var fill = cell.Style.Fill;
                if (fill.PatternType == XLFillPatternValues.None)
                {
                    continue; // "explicit fill" only — a default (no) fill is not ours to grade
                }

                if (!TryColor(fill.BackgroundColor, out var bg) ||
                    !TryColor(cell.Style.Font.FontColor, out var fg))
                {
                    continue; // theme/indexed colors we cannot resolve to RGB are skipped, not guessed
                }

                var ratio = ContrastRatio(fg, bg);
                if (ratio >= 4.5)
                {
                    continue;
                }

                var path = ExcelPaths.CellPath(sheet, cell.Address);
                findings.Add(new AuditFinding
                {
                    Id = $"a11y_low_contrast#{path}",
                    Severity = SeverityWarning,
                    Category = CategoryAccessibility,
                    Code = "a11y_low_contrast",
                    Path = path,
                    Message = string.Create(
                        CultureInfo.InvariantCulture,
                        $"Cell {path} has a font/fill contrast of {ratio:0.0}:1 (WCAG AA needs 4.5:1)."),
                    Suggestion = "Darken the text or lighten the fill so the contrast reaches 4.5:1: " +
                                 "{op:set, path:" + path + ", props:{color:\"#000000\"}}.",
                    Autofixable = false,
                });
            }
        }
    }

    private static void NoAltText(string file, List<AuditFinding> findings)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return;
        }

        var charts = ExcelCharts.Read(document).ToLookup(c => c.SheetName, StringComparer.OrdinalIgnoreCase);

        foreach (var sheet in workbookPart.Workbook.Descendants<S.Sheet>())
        {
            var sheetName = sheet.Name?.Value ?? string.Empty;
            if (sheet.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart ||
                worksheetPart.DrawingsPart?.WorksheetDrawing is not { } root)
            {
                continue;
            }

            var imageIndex = 0;
            var chartIndex = 0;
            foreach (var anchor in root.ChildElements
                         .Where(c => c is Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor))
            {
                if (anchor.Descendants<Xdr.Picture>().FirstOrDefault() is { } picture)
                {
                    imageIndex++;
                    var props = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
                    if (string.IsNullOrWhiteSpace(props?.Description?.Value) &&
                        string.IsNullOrWhiteSpace(props?.Title?.Value))
                    {
                        var path = string.Create(
                            CultureInfo.InvariantCulture, $"/{ExcelPaths.QuoteSheet(sheetName)}/image[{imageIndex}]");
                        findings.Add(AltTextFinding(path, "Image", sheetName));
                    }
                }
                else if (anchor.Descendants<Xdr.GraphicFrame>().Any(g =>
                             g.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any()))
                {
                    chartIndex++;
                    var props = anchor.Descendants<Xdr.NonVisualDrawingProperties>().FirstOrDefault();
                    var chartInfo = charts[sheetName].FirstOrDefault(c => c.Index == chartIndex);
                    var hasTitle = !string.IsNullOrWhiteSpace(chartInfo?.Title);
                    if (string.IsNullOrWhiteSpace(props?.Description?.Value) &&
                        string.IsNullOrWhiteSpace(props?.Title?.Value) && !hasTitle)
                    {
                        var path = string.Create(
                            CultureInfo.InvariantCulture, $"/{ExcelPaths.QuoteSheet(sheetName)}/chart[{chartIndex}]");
                        findings.Add(AltTextFinding(path, "Chart", sheetName));
                    }
                }
            }
        }
    }

    private static AuditFinding AltTextFinding(string path, string kind, string sheetName) => new()
    {
        Id = $"a11y_no_alt_text#{path}",
        Severity = SeverityWarning,
        Category = CategoryAccessibility,
        Code = "a11y_no_alt_text",
        Path = path,
        Message = $"{kind} {path} on '{sheetName}' has no alt-text.",
        Suggestion = "Add a description so screen readers can convey the visual. --fix inserts a placeholder you should then edit.",
        Autofixable = true,
    };

    // ----- contrast math ------------------------------------------------------

    /// <summary>WCAG 2.x relative-luminance contrast ratio (1:1 .. 21:1).</summary>
    public static double ContrastRatio((double R, double G, double B) a, (double R, double G, double B) b)
    {
        var la = RelativeLuminance(a);
        var lb = RelativeLuminance(b);
        var lighter = Math.Max(la, lb);
        var darker = Math.Min(la, lb);
        return (lighter + 0.05) / (darker + 0.05);
    }

    private static double RelativeLuminance((double R, double G, double B) c) =>
        (0.2126 * Linearize(c.R)) + (0.7152 * Linearize(c.G)) + (0.0722 * Linearize(c.B));

    private static double Linearize(double channel)
    {
        var c = channel / 255.0;
        return c <= 0.03928 ? c / 12.92 : Math.Pow((c + 0.055) / 1.055, 2.4);
    }

    private static bool TryColor(XLColor color, out (double R, double G, double B) rgb)
    {
        rgb = default;
        if (color is null || !color.HasValue)
        {
            return false;
        }

        // Only fully-resolved RGB colors are graded; theme/indexed colors that
        // ClosedXML does not expand stay unjudged rather than mis-judged.
        try
        {
            if (color.ColorType != XLColorType.Color)
            {
                return false;
            }

            var c = color.Color;
            rgb = (c.R, c.G, c.B);
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    // ----- helpers ------------------------------------------------------------

    private static void Track(Dictionary<string, List<string>> map, string key, string scope)
    {
        if (!map.TryGetValue(key, out var scopes))
        {
            scopes = [];
            map[key] = scopes;
        }

        scopes.Add(scope);
    }

    private static bool IsBuiltin(string name) => name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase);
}
