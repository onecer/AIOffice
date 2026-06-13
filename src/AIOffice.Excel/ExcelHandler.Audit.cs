using System.Diagnostics;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// The <see cref="IAuditor"/> half of the xlsx handler (M7). Detection is in
/// <see cref="ExcelAudit"/>; this file wires it to the verb surface and owns the
/// safe-only fix pass, which rides the same snapshot/atomic-save discipline as
/// every other edit.
/// </summary>
public sealed partial class ExcelHandler : IAuditor
{
    private const string AltTextPlaceholder = "(describe this image)";

    public AuditResult Audit(CommandContext ctx, AuditOptions opts)
    {
        var file = RequireFile(ctx, mustExist: true);
        return ExcelAudit.Run(file, opts);
    }

    public int Fix(CommandContext ctx, IReadOnlyList<string> findingIds)
    {
        var file = RequireFile(ctx, mustExist: true);

        // Re-audit (every category) so finding ids are fresh against the file on
        // disk; an empty id list means "fix everything autofixable".
        var result = ExcelAudit.Run(file, new AuditOptions { Category = "all", MinSeverity = "info" });
        var wanted = findingIds.Count == 0 ? null : findingIds.ToHashSet(StringComparer.Ordinal);
        var targets = result.Findings
            .Where(f => f.Autofixable && (wanted is null || wanted.Contains(f.Id)))
            .ToList();
        if (targets.Count == 0)
        {
            return 0;
        }

        var snapshot = _snapshots.Save(file); // every fix pass is undoable
        try
        {
            return ApplyFixes(file, targets);
        }
        catch (Exception)
        {
            File.Copy(snapshot.Path, file, overwrite: true);
            throw;
        }
    }

    /// <summary>
    /// Applies the safe autofixes. Doc-title fixes go through ClosedXML (a
    /// property write); alt-text fixes are authored raw on the drawing parts
    /// (ClosedXML cannot see chart titles/picture descr).
    /// </summary>
    private static int ApplyFixes(string file, IReadOnlyList<AuditFinding> targets)
    {
        var fixedCount = 0;

        var titleFix = targets.FirstOrDefault(f => f.Code == "a11y_no_doc_title");
        if (titleFix is not null)
        {
            using var workbook = OpenWorkbook(file);
            workbook.Properties.Title = TitleFromFile(file, workbook);
            workbook.Save();
            fixedCount++;
        }

        var altFixes = targets.Where(f => f.Code == "a11y_no_alt_text").ToList();
        if (altFixes.Count > 0)
        {
            fixedCount += FixAltText(file, altFixes);
        }

        return fixedCount;
    }

    private static string TitleFromFile(string file, XLWorkbook workbook)
    {
        var fromName = Path.GetFileNameWithoutExtension(file);
        if (!string.IsNullOrWhiteSpace(fromName))
        {
            return fromName;
        }

        return workbook.Worksheets.FirstOrDefault()?.Name ?? "Workbook";
    }

    /// <summary>Writes a placeholder description on every image/chart finding's drawing node.</summary>
    private static int FixAltText(string file, IReadOnlyList<AuditFinding> altFixes)
    {
        var wantedPaths = altFixes.Select(f => f.Path).Where(p => p is not null).ToHashSet(StringComparer.Ordinal);

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return 0;
        }

        var fixedCount = 0;
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
            var touched = false;
            foreach (var anchor in root.ChildElements
                         .Where(c => c is Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor))
            {
                if (anchor.Descendants<Xdr.Picture>().FirstOrDefault() is { } picture)
                {
                    imageIndex++;
                    var path = $"/{ExcelPaths.QuoteSheet(sheetName)}/image[{imageIndex}]";
                    if (wantedPaths.Contains(path) &&
                        picture.NonVisualPictureProperties?.NonVisualDrawingProperties is { } props)
                    {
                        props.Description = AltTextPlaceholder;
                        touched = true;
                        fixedCount++;
                    }
                }
                else if (anchor.Descendants<Xdr.GraphicFrame>().Any(g =>
                             g.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any()))
                {
                    chartIndex++;
                    var path = $"/{ExcelPaths.QuoteSheet(sheetName)}/chart[{chartIndex}]";
                    if (wantedPaths.Contains(path) &&
                        anchor.Descendants<Xdr.NonVisualDrawingProperties>().FirstOrDefault() is { } props)
                    {
                        props.Description = AltTextPlaceholder;
                        touched = true;
                        fixedCount++;
                    }
                }
            }

            if (touched)
            {
                root.Save();
            }
        }

        return fixedCount;
    }

    /// <summary>Opens a workbook for the audit's read-only scan (shared error mapping).</summary>
    internal static XLWorkbook OpenWorkbookForAudit(string file) => OpenWorkbook(file);
}
