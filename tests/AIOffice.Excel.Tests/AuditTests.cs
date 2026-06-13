using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M7 audit flagship: the xlsx <see cref="IAuditor"/> implementation finds
/// accessibility/quality problems and applies only the safe autofixes. Each
/// fact crafts the exact defect it checks for, then verifies detection, the
/// severity tally, the fix, and the re-audit going clean.
/// </summary>
public sealed class AuditTests : ExcelTestBase
{
    /// <summary>A 4x2 red PNG, identical to the image-suite fixture.</summary>
    private static readonly byte[] PngBytes = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAEElEQVR4nGP4z8AARwzIHABvqgf5gNwAKAAAAABJRU5ErkJggg==");

    private AuditResult AuditAll(string file) =>
        Handler.Audit(Ctx(file), new AuditOptions { Category = "all", MinSeverity = "info" });

    private static AuditFinding? FindCode(AuditResult result, string code) =>
        result.Findings.FirstOrDefault(f => f.Code == code);

    // ----- quality_formula_error ----------------------------------------------

    [Fact]
    public void Detects_a_div_by_zero_formula_error()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 10.0)),
            SetOp("/Sheet1/A2", ("value", 0.0)),
            SetOp("/Sheet1/A3", ("value", "=A1/A2"))).IsOk);

        var result = AuditAll(file);

        var finding = FindCode(result, "quality_formula_error");
        Assert.NotNull(finding);
        Assert.Equal("error", finding!.Severity);
        Assert.Equal("quality", finding.Category);
        Assert.Equal("/Sheet1/A3", finding.Path);
        Assert.False(finding.Autofixable);
        Assert.True(result.Summary.Errors >= 1);
    }

    [Fact]
    public void Detects_a_ref_error_in_a_cached_value()
    {
        var file = CreateWorkbook();
        // Write a cell whose cached value is the #REF! error directly.
        Assert.True(EditOps(file, SetOp("/Sheet1/B2", ("value", "#REF!"), ("valueType", "text"))).IsOk);

        // Inject a real #REF! error value via raw OpenXml so the cached value is
        // an error type (Excel's own representation), not text.
        InjectErrorCell(file, "Sheet1", "C3", "#DIV/0!");

        var result = AuditAll(file);
        var finding = result.Findings.FirstOrDefault(f =>
            f.Code == "quality_formula_error" && f.Path == "/Sheet1/C3");
        Assert.NotNull(finding);
    }

    [Fact]
    public void Detects_a_broken_ref_in_formula_text()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", 5.0))).IsOk);
        InjectFormulaCell(file, "Sheet1", "A2", "SUM(#REF!)", "0");

        var result = AuditAll(file);

        var finding = FindCode(result, "quality_broken_ref");
        Assert.NotNull(finding);
        Assert.Equal("error", finding!.Severity);
        Assert.Equal("/Sheet1/A2", finding.Path);
    }

    // ----- quality_duplicate_id -----------------------------------------------

    [Fact]
    public void Detects_a_defined_name_in_two_scopes()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1:A2", ("values", new JsonArray(new JsonArray(1), new JsonArray(2)))),
            AddOp("/Sheet1/A1:A2", "name", ("name", "Shared")),
            AddOp("/Sheet1/A1:A2", "name", ("name", "Shared"), ("scope", "sheet"))).IsOk);

        var result = AuditAll(file);

        var finding = FindCode(result, "quality_duplicate_id");
        Assert.NotNull(finding);
        Assert.Equal("error", finding!.Severity);
        Assert.Contains("Shared", finding.Message, StringComparison.Ordinal);
    }

    // ----- a11y_merged_data_cells ---------------------------------------------

    [Fact]
    public void Detects_a_merge_inside_the_data_region()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:C3", ("values", new JsonArray(
            new JsonArray(1, 2, 3),
            new JsonArray(4, 5, 6),
            new JsonArray(7, 8, 9))))).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/B2:C2", ("merge", true))).IsOk);

        var result = AuditAll(file);

        var finding = FindCode(result, "a11y_merged_data_cells");
        Assert.NotNull(finding);
        Assert.Equal("warning", finding!.Severity);
        Assert.Equal("accessibility", finding.Category);
        Assert.Equal("/Sheet1/B2:C2", finding.Path);
        Assert.False(finding.Autofixable);
    }

    // ----- a11y_low_contrast --------------------------------------------------

    [Fact]
    public void Detects_low_contrast_text_on_an_explicit_fill()
    {
        var file = CreateWorkbook();
        // Light-grey text on a white fill: well under 4.5:1.
        Assert.True(EditOps(file, SetOp(
            "/Sheet1/A1",
            ("value", "Hi"), ("fill", "#FFFFFF"), ("color", "#DDDDDD"))).IsOk);

        var result = AuditAll(file);

        var finding = FindCode(result, "a11y_low_contrast");
        Assert.NotNull(finding);
        Assert.Equal("warning", finding!.Severity);
        Assert.Equal("/Sheet1/A1", finding.Path);
    }

    [Fact]
    public void High_contrast_text_is_not_flagged()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp(
            "/Sheet1/A1",
            ("value", "Hi"), ("fill", "#FFFFFF"), ("color", "#000000"))).IsOk);

        var result = AuditAll(file);

        Assert.Null(FindCode(result, "a11y_low_contrast"));
    }

    [Fact]
    public void Contrast_math_sits_on_the_wcag_boundary()
    {
        // #767676 on white is ~4.54:1 — just over the 4.5:1 AA threshold, so it
        // must NOT be flagged; #777777 on white is ~4.48:1 and must be.
        var passes = CreateWorkbook("pass.xlsx");
        Assert.True(EditOps(passes, SetOp(
            "/Sheet1/A1", ("value", "x"), ("fill", "#FFFFFF"), ("color", "#767676"))).IsOk);
        Assert.Null(FindCode(AuditAll(passes), "a11y_low_contrast"));

        var fails = CreateWorkbook("fail.xlsx");
        Assert.True(EditOps(fails, SetOp(
            "/Sheet1/A1", ("value", "x"), ("fill", "#FFFFFF"), ("color", "#777777"))).IsOk);
        Assert.NotNull(FindCode(AuditAll(fails), "a11y_low_contrast"));
    }

    // ----- quality_hidden_data (info) -----------------------------------------

    [Fact]
    public void Detects_a_hidden_row_with_data_as_info()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "visible")),
            SetOp("/Sheet1/A2", ("value", "secret")),
            SetOp("/Sheet1/row[2]", ("hidden", true))).IsOk);

        var result = AuditAll(file);

        var finding = result.Findings.FirstOrDefault(f =>
            f.Code == "quality_hidden_data" && f.Path == "/Sheet1/row[2]");
        Assert.NotNull(finding);
        Assert.Equal("info", finding!.Severity);
        Assert.True(result.Summary.Infos >= 1);
    }

    // ----- a11y_no_alt_text + fix ---------------------------------------------

    [Fact]
    public void Detects_image_without_alt_text_then_fix_and_reaudit_is_clean()
    {
        var file = CreateWorkbook();
        File.WriteAllBytes(Path.Combine(Dir, "logo.png"), PngBytes);
        Assert.True(EditOps(file, AddOp(
            "/Sheet1", "image", ("src", "logo.png"), ("anchor", "E2"))).IsOk);

        var before = AuditAll(file);
        var altFinding = FindCode(before, "a11y_no_alt_text");
        Assert.NotNull(altFinding);
        Assert.True(altFinding!.Autofixable);
        Assert.Equal("/Sheet1/image[1]", altFinding.Path);

        // Fix everything autofixable (empty id list = fix all).
        var fixedCount = Handler.Fix(Ctx(file), []);
        Assert.True(fixedCount >= 1);
        AssertValidatorClean(file);

        // Re-audit: the alt-text finding is gone, and the placeholder is on disk.
        var after = AuditAll(file);
        Assert.Null(FindCode(after, "a11y_no_alt_text"));

        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var descr = document.WorkbookPart!.Workbook!.Descendants<S.Sheet>()
            .Select(s => (WorksheetPart)document.WorkbookPart.GetPartById(s.Id!.Value!))
            .SelectMany(wp => wp.DrawingsPart?.WorksheetDrawing?.Descendants<Xdr.Picture>() ?? [])
            .Select(p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Description?.Value)
            .FirstOrDefault(d => d is not null);
        Assert.Equal("(describe this image)", descr);
    }

    // ----- a11y_no_doc_title + fix --------------------------------------------

    [Fact]
    public void Missing_doc_title_is_flagged_then_fixed_from_the_filename()
    {
        var file = CreateWorkbook("revenue.xlsx");

        var before = AuditAll(file);
        var titleFinding = FindCode(before, "a11y_no_doc_title");
        Assert.NotNull(titleFinding);
        Assert.True(titleFinding!.Autofixable);

        var fixedCount = Handler.Fix(Ctx(file), [titleFinding.Id]);
        Assert.Equal(1, fixedCount);
        AssertValidatorClean(file);

        var core = OkData(Handler.Read(Ctx(file, ("view", "properties"))))["core"]!;
        Assert.Equal("revenue", core["title"]!.GetValue<string>());

        var after = AuditAll(file);
        Assert.Null(FindCode(after, "a11y_no_doc_title"));
    }

    // ----- filtering ----------------------------------------------------------

    [Fact]
    public void Category_filter_excludes_the_other_category()
    {
        var file = CreateWorkbook(); // empty workbook → only a11y_no_doc_title fires

        var quality = Handler.Audit(Ctx(file), new AuditOptions { Category = "quality", MinSeverity = "info" });
        Assert.DoesNotContain(quality.Findings, f => f.Category == "accessibility");

        var accessibility = Handler.Audit(
            Ctx(file), new AuditOptions { Category = "accessibility", MinSeverity = "info" });
        Assert.Contains(accessibility.Findings, f => f.Code == "a11y_no_doc_title");
    }

    [Fact]
    public void Min_severity_filter_drops_lower_severities()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "x")),
            SetOp("/Sheet1/A2", ("value", "y")),
            SetOp("/Sheet1/row[2]", ("hidden", true))).IsOk);

        // info-level hidden-data findings disappear at minSeverity=warning.
        var warned = Handler.Audit(Ctx(file), new AuditOptions { Category = "all", MinSeverity = "warning" });
        Assert.DoesNotContain(warned.Findings, f => f.Severity == "info");
    }

    [Fact]
    public void Fix_with_no_autofixable_findings_reports_zero()
    {
        var file = CreateWorkbook();
        // Give it a title so a11y_no_doc_title (the only other autofixable
        // finding) does not fire, leaving a non-autofixable error behind.
        Assert.True(EditOps(file, new EditOp
        {
            Op = "set",
            Path = "/properties",
            Props = new JsonObject { ["title"] = "Titled" },
        }).IsOk);

        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1.0)),
            SetOp("/Sheet1/A2", ("value", 0.0)),
            SetOp("/Sheet1/A3", ("value", "=A1/A2"))).IsOk);

        var onlyErrors = Handler.Fix(Ctx(file), []);
        Assert.Equal(0, onlyErrors);
    }

    // ----- raw fixture helpers ------------------------------------------------

    /// <summary>Writes a cell whose stored value is an Excel error type.</summary>
    private static void InjectErrorCell(string file, string sheetName, string address, string error)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var (worksheetPart, sheetData) = SheetData(document, sheetName);
        var cell = EnsureCell(sheetData, address);
        cell.DataType = S.CellValues.Error;
        cell.CellValue = new S.CellValue(error);
        cell.CellFormula = null;
        worksheetPart.Worksheet!.Save();
    }

    /// <summary>Writes a formula cell with a chosen formula text and cached value.</summary>
    private static void InjectFormulaCell(
        string file, string sheetName, string address, string formula, string cached)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var (worksheetPart, sheetData) = SheetData(document, sheetName);
        var cell = EnsureCell(sheetData, address);
        cell.CellFormula = new S.CellFormula(formula);
        cell.CellValue = new S.CellValue(cached);
        cell.DataType = null;
        worksheetPart.Worksheet!.Save();
    }

    private static (WorksheetPart Part, S.SheetData Data) SheetData(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>()
            .First(s => s.Name?.Value == sheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        var sheetData = worksheetPart.Worksheet!.GetFirstChild<S.SheetData>()!;
        return (worksheetPart, sheetData);
    }

    private static S.Cell EnsureCell(S.SheetData sheetData, string address)
    {
        var rowNumber = uint.Parse(new string(address.Where(char.IsDigit).ToArray()));
        var row = sheetData.Elements<S.Row>().FirstOrDefault(r => r.RowIndex?.Value == rowNumber);
        if (row is null)
        {
            row = new S.Row { RowIndex = rowNumber };
            sheetData.AppendChild(row);
        }

        var cell = row.Elements<S.Cell>().FirstOrDefault(c => c.CellReference?.Value == address);
        if (cell is null)
        {
            cell = new S.Cell { CellReference = address };
            row.AppendChild(cell);
        }

        return cell;
    }
}
