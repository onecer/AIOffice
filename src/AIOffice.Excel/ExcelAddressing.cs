using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>What an xlsx path resolved to.</summary>
internal enum ExcelTargetKind
{
    Sheet,
    Cell,
    Range,
    Row,
}

/// <summary>A resolved xlsx address: the worksheet plus an optional cell/range/row.</summary>
internal sealed record ExcelTarget
{
    public required ExcelTargetKind Kind { get; init; }

    public required IXLWorksheet Sheet { get; init; }

    /// <summary>Set when <see cref="Kind"/> is Cell.</summary>
    public IXLCell? Cell { get; init; }

    /// <summary>Set when <see cref="Kind"/> is Range.</summary>
    public IXLRange? Range { get; init; }

    /// <summary>1-based worksheet row number when <see cref="Kind"/> is Row.</summary>
    public int? RowNumber { get; init; }
}

/// <summary>
/// xlsx addressing: <c>/Sheet1/A1</c>, <c>/Sheet1/A1:C10</c>, <c>/Sheet1/row[3]</c>,
/// <c>/'Q3 Data'/B2</c>. Resolution failures throw <c>invalid_path</c> with
/// nearest-match candidates, as the envelope contract requires.
/// </summary>
internal static partial class ExcelPaths
{
    [GeneratedRegex("^[A-Za-z_][A-Za-z0-9_.-]*$")]
    private static partial Regex BareName();

    [GeneratedRegex("^[A-Z]{1,3}[0-9]{1,7}(:[A-Z]{1,3}[0-9]{1,7})?$")]
    private static partial Regex CellOrRange();

    /// <summary>Quotes a sheet name when it would not survive the path grammar bare.</summary>
    public static string QuoteSheet(string name) =>
        BareName().IsMatch(name) && !CellOrRange().IsMatch(name)
            ? name
            : "'" + name.Replace("'", "''", StringComparison.Ordinal) + "'";

    public static string SheetPath(IXLWorksheet sheet) => "/" + QuoteSheet(sheet.Name);

    public static string CellPath(IXLWorksheet sheet, IXLAddress address) =>
        SheetPath(sheet) + "/" + address.ColumnLetter +
        address.RowNumber.ToString(CultureInfo.InvariantCulture);

    public static string RangePath(IXLWorksheet sheet, IXLRangeAddress address)
    {
        var first = address.FirstAddress;
        var last = address.LastAddress;
        if (first.RowNumber == last.RowNumber && first.ColumnNumber == last.ColumnNumber)
        {
            return CellPath(sheet, first);
        }

        return string.Create(
            CultureInfo.InvariantCulture,
            $"{SheetPath(sheet)}/{first.ColumnLetter}{first.RowNumber}:{last.ColumnLetter}{last.RowNumber}");
    }

    /// <summary>Resolves an xlsx path against an open workbook.</summary>
    /// <exception cref="AiofficeException"><c>invalid_path</c> with candidates when the address does not resolve.</exception>
    public static ExcelTarget Resolve(XLWorkbook workbook, string pathText)
    {
        var path = DocPath.Parse(pathText);
        var sheet = ResolveSheet(workbook, path, pathText);

        if (path.Segments.Count == 1)
        {
            return new ExcelTarget { Kind = ExcelTargetKind.Sheet, Sheet = sheet };
        }

        if (path.Segments.Count > 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"xlsx paths have at most two segments (sheet, then cell/range/row): {pathText}",
                "Use /Sheet1/A1, /Sheet1/A1:C10 or /Sheet1/row[3].",
                candidates: ExampleTargets(sheet));
        }

        var segment = path.Segments[1];
        switch (segment.Kind)
        {
            case PathSegmentKind.Cell:
                return new ExcelTarget
                {
                    Kind = ExcelTargetKind.Cell,
                    Sheet = sheet,
                    Cell = sheet.Cell(segment.Start!.Value.ToString()),
                };

            case PathSegmentKind.Range:
                return new ExcelTarget
                {
                    Kind = ExcelTargetKind.Range,
                    Sheet = sheet,
                    Range = sheet.Range($"{segment.Start!.Value}:{segment.End!.Value}"),
                };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "row", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } rowNumber:
                return new ExcelTarget { Kind = ExcelTargetKind.Row, Sheet = sheet, RowNumber = rowNumber };

            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"'{segment.ToCanonicalString()}' is not a cell, range or row[n] in: {pathText}",
                    "After the sheet name use A1, A1:C10 or row[3]; column letters are uppercase.",
                    candidates: ExampleTargets(sheet));
        }
    }

    private static IXLWorksheet ResolveSheet(XLWorkbook workbook, DocPath path, string pathText)
    {
        var first = path.Segments[0];
        var sheetName = first.Kind switch
        {
            PathSegmentKind.Name => first.Name,
            PathSegmentKind.Element when first.Index is null => first.Name,
            _ => null,
        };

        if (sheetName is null)
        {
            var hint = first.Kind is PathSegmentKind.Cell or PathSegmentKind.Range
                ? $"'{first.ToCanonicalString()}' parsed as a cell reference, but xlsx paths start with a sheet name. " +
                  "Quote sheet names that look like cells: /'Q3'/A1."
                : "Sheets are addressed by name, not by index, in v0.";
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"xlsx paths must start with a sheet name: {pathText}",
                hint,
                candidates: SheetCandidates(workbook, first.Name ?? first.ToCanonicalString()));
        }

        if (!workbook.TryGetWorksheet(sheetName, out var sheet))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No sheet named '{sheetName}' exists in the workbook.",
                "Sheet names are matched case-insensitively; pick one of the candidates.",
                candidates: SheetCandidates(workbook, sheetName));
        }

        return sheet;
    }

    /// <summary>All sheet paths ordered by edit distance to the requested name (nearest first).</summary>
    public static IReadOnlyList<string> SheetCandidates(XLWorkbook workbook, string requested) =>
        [.. workbook.Worksheets
            .OrderBy(ws => Levenshtein(requested, ws.Name))
            .ThenBy(ws => ws.Position)
            .Select(SheetPath)
            .Take(5)];

    private static List<string> ExampleTargets(IXLWorksheet sheet)
    {
        var basePath = SheetPath(sheet);
        return [basePath + "/A1", basePath + "/A1:C10", basePath + "/row[1]"];
    }

    /// <summary>Classic Levenshtein edit distance, case-insensitive.</summary>
    internal static int Levenshtein(string a, string b)
    {
        a = a.ToUpperInvariant();
        b = b.ToUpperInvariant();
        var previous = new int[b.Length + 1];
        var current = new int[b.Length + 1];
        for (var j = 0; j <= b.Length; j++)
        {
            previous[j] = j;
        }

        for (var i = 1; i <= a.Length; i++)
        {
            current[0] = i;
            for (var j = 1; j <= b.Length; j++)
            {
                var substitution = previous[j - 1] + (a[i - 1] == b[j - 1] ? 0 : 1);
                current[j] = Math.Min(Math.Min(previous[j] + 1, current[j - 1] + 1), substitution);
            }

            (previous, current) = (current, previous);
        }

        return previous[b.Length];
    }
}
