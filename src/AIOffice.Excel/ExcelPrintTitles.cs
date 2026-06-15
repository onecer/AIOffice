using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// (1.7) Clearing the print-title repeat bands (rows-to-repeat / cols-to-repeat).
/// ClosedXML 0.105 can SET a band (<c>SetRowsToRepeatAtTop</c>) but throws on every
/// clear form (0,0 and "" both throw), and the reserved <c>_xlnm.Print_Titles</c>
/// name is not exposed through its NamedRanges API. So a CLEAR is done raw on the
/// saved bytes: the print-titles defined name is <c>Sheet!$A:$A,Sheet!$1:$1</c> —
/// drop the row-band or column-band segment, and remove the whole defined name when
/// both are gone. Setting still rides ClosedXML (it works and round-trips).
/// </summary>
internal static partial class ExcelPrintTitles
{
    /// <summary>One queued print-title clear (which axis, on which sheet).</summary>
    internal sealed record ClearSpec(string SheetName, bool ClearRows, bool ClearCols);

    [GeneratedRegex(@"^.+!\$?[0-9]+:\$?[0-9]+$")]
    private static partial Regex RowBandSegment();

    [GeneratedRegex(@"^.+!\$?[A-Za-z]+:\$?[A-Za-z]+$")]
    private static partial Regex ColumnBandSegment();

    /// <summary>Applies queued print-title clears to the file ClosedXML just saved.</summary>
    public static void Apply(string file, IReadOnlyList<ClearSpec> specs)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbook = document.WorkbookPart?.Workbook;
        if (workbook?.DefinedNames is not { } definedNames)
        {
            return;
        }

        foreach (var spec in specs)
        {
            var sheetIndex = SheetLocalIndex(workbook, spec.SheetName);
            if (sheetIndex is null)
            {
                continue;
            }

            var defined = definedNames.Elements<S.DefinedName>().FirstOrDefault(n =>
                n.Name?.Value == "_xlnm.Print_Titles" && n.LocalSheetId?.Value == sheetIndex);
            if (defined is null)
            {
                continue;
            }

            var segments = (defined.Text ?? string.Empty)
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Where(segment =>
                    !(spec.ClearRows && RowBandSegment().IsMatch(segment)) &&
                    !(spec.ClearCols && ColumnBandSegment().IsMatch(segment)))
                .ToList();

            if (segments.Count == 0)
            {
                defined.Remove();
            }
            else
            {
                defined.Text = string.Join(",", segments);
            }
        }

        workbook.Save();
    }

    /// <summary>The 0-based local sheet id (definedName scope) for a sheet name, or null.</summary>
    private static uint? SheetLocalIndex(S.Workbook workbook, string sheetName)
    {
        var sheets = workbook.Sheets?.Elements<S.Sheet>().ToList() ?? [];
        for (var i = 0; i < sheets.Count; i++)
        {
            if (string.Equals(sheets[i].Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                return (uint)i;
            }
        }

        return null;
    }
}
