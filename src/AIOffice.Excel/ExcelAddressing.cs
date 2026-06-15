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
    Column,
    Chart,
    Pivot,
    ConditionalFormat,
    Image,
    DataValidation,
    Sparkline,
    Comment,
    Table,
    Slicer,
    Embed,
    FormControl,
    DataTable,
    Scenario,
    LinkedPicture,
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

    /// <summary>1-based worksheet column number when <see cref="Kind"/> is Column.</summary>
    public int? ColumnNumber { get; init; }

    /// <summary>1-based per-sheet chart index when <see cref="Kind"/> is Chart.</summary>
    public int? ChartIndex { get; init; }

    /// <summary>1-based per-sheet pivot index when <see cref="Kind"/> is Pivot (null when addressed by name).</summary>
    public int? PivotIndex { get; init; }

    /// <summary>Pivot table name when <see cref="Kind"/> is Pivot and the path used <c>[@name=…]</c>.</summary>
    public string? PivotName { get; init; }

    /// <summary>1-based per-sheet conditional-format index when <see cref="Kind"/> is ConditionalFormat.</summary>
    public int? ConditionalFormatIndex { get; init; }

    /// <summary>1-based per-sheet image index when <see cref="Kind"/> is Image.</summary>
    public int? ImageIndex { get; init; }

    /// <summary>1-based per-sheet data-validation index when <see cref="Kind"/> is DataValidation.</summary>
    public int? DataValidationIndex { get; init; }

    /// <summary>1-based per-sheet sparkline index when <see cref="Kind"/> is Sparkline.</summary>
    public int? SparklineIndex { get; init; }

    /// <summary>Bare thread GUID when <see cref="Kind"/> is Comment (<c>comment[@id=…]</c>).</summary>
    public string? CommentId { get; init; }

    /// <summary>Table name when <see cref="Kind"/> is Table (<c>table[@name=…]</c>).</summary>
    public string? TableName { get; init; }

    /// <summary>1-based per-sheet slicer index when <see cref="Kind"/> is Slicer (null when addressed by name).</summary>
    public int? SlicerIndex { get; init; }

    /// <summary>Slicer name when <see cref="Kind"/> is Slicer and the path used <c>[@name=…]</c>.</summary>
    public string? SlicerName { get; init; }

    /// <summary>1-based per-sheet embed index when <see cref="Kind"/> is Embed.</summary>
    public int? EmbedIndex { get; init; }

    /// <summary>1-based per-sheet form-control index when <see cref="Kind"/> is FormControl.</summary>
    public int? FormControlIndex { get; init; }

    /// <summary>1-based per-sheet data-table index when <see cref="Kind"/> is DataTable (1.4).</summary>
    public int? DataTableIndex { get; init; }

    /// <summary>Scenario name when <see cref="Kind"/> is Scenario (<c>scenario[@name=…]</c>, 1.5).</summary>
    public string? ScenarioName { get; init; }

    /// <summary>1-based per-sheet linked-picture index when <see cref="Kind"/> is LinkedPicture (1.7).</summary>
    public int? LinkedPictureIndex { get; init; }
}

/// <summary>
/// xlsx addressing: <c>/Sheet1/A1</c>, <c>/Sheet1/A1:C10</c>, <c>/Sheet1/row[3]</c>,
/// <c>/Sheet1/col[C]</c>, <c>/Sheet1/chart[1]</c>, <c>/Sheet1/pivot[1]</c>,
/// <c>/Pivot/pivot[@name=Sales]</c>, <c>/Sheet1/conditionalFormat[1]</c>,
/// <c>/Sheet1/image[1]</c>, <c>/Sheet1/dataValidation[1]</c>,
/// <c>/Sheet1/sparkline[1]</c>, <c>/Sheet1/comment[@id=GUID]</c>,
/// <c>/'Q3 Data'/B2</c>.
/// Resolution failures throw <c>invalid_path</c> with nearest-match candidates,
/// as the envelope contract requires.
/// </summary>
internal static partial class ExcelPaths
{
    [GeneratedRegex("^[A-Za-z_][A-Za-z0-9_.-]*$")]
    private static partial Regex BareName();

    [GeneratedRegex("^[A-Z]{1,3}[0-9]{1,7}(:[A-Z]{1,3}[0-9]{1,7})?$")]
    private static partial Regex CellOrRange();

    /// <summary>
    /// The stable-name pivot form <c>/Sheet/pivot[@name=X]</c> (bare or
    /// <c>'quoted'</c> name). Pre-parsed here because the shared DocPath
    /// grammar has no attribute predicates.
    /// </summary>
    [GeneratedRegex(@"^(?<sheet>/.+)/(?i:pivot)\[@name=(?:'(?<quoted>(?:[^']|'')+)'|(?<bare>[^\]]+))\]$")]
    private static partial Regex PivotByName();

    /// <summary>
    /// The stable-name table form <c>/Sheet/table[@name=X]</c> (bare or
    /// <c>'quoted'</c> name). Pre-parsed here, exactly like the pivot form,
    /// because the shared DocPath grammar has no attribute predicates for names
    /// with spaces/specials.
    /// </summary>
    [GeneratedRegex(@"^(?<sheet>/.+)/(?i:table)\[@name=(?:'(?<quoted>(?:[^']|'')+)'|(?<bare>[^\]]+))\]$")]
    private static partial Regex TableByName();

    /// <summary>
    /// The stable-name slicer form <c>/Sheet/slicer[@name=X]</c> (bare or
    /// <c>'quoted'</c> name), peeled off like the pivot/table forms.
    /// </summary>
    [GeneratedRegex(@"^(?<sheet>/.+)/(?i:slicer)\[@name=(?:'(?<quoted>(?:[^']|'')+)'|(?<bare>[^\]]+))\]$")]
    private static partial Regex SlicerByName();

    /// <summary>
    /// The stable-name scenario form <c>/Sheet/scenario[@name=X]</c> (bare or
    /// <c>'quoted'</c> name), peeled off like the pivot/table/slicer forms (1.5).
    /// </summary>
    [GeneratedRegex(@"^(?<sheet>/.+)/(?i:scenario)\[@name=(?:'(?<quoted>(?:[^']|'')+)'|(?<bare>[^\]]+))\]$")]
    private static partial Regex ScenarioByName();

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
        // /Sheet/pivot[@name=X] never survives the shared DocPath grammar, so
        // it is peeled off first (precedent: pptx parses shape[@id=N] itself).
        var byName = PivotByName().Match(pathText);
        if (byName.Success)
        {
            var sheetTarget = Resolve(workbook, byName.Groups["sheet"].Value);
            if (sheetTarget.Kind != ExcelTargetKind.Sheet)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"pivot[@name=…] must follow a sheet name: {pathText}",
                    "Use /SheetName/pivot[@name=X]; quote names with specials: pivot[@name='Q3 Sales'].");
            }

            var name = byName.Groups["quoted"].Success
                ? byName.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal)
                : byName.Groups["bare"].Value;
            return new ExcelTarget { Kind = ExcelTargetKind.Pivot, Sheet = sheetTarget.Sheet, PivotName = name };
        }

        // /Sheet/table[@name=X] gets the same id-form peel-off as pivots.
        var byTable = TableByName().Match(pathText);
        if (byTable.Success)
        {
            var sheetTarget = Resolve(workbook, byTable.Groups["sheet"].Value);
            if (sheetTarget.Kind != ExcelTargetKind.Sheet)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"table[@name=…] must follow a sheet name: {pathText}",
                    "Use /SheetName/table[@name=X]; quote names with specials: table[@name='Q3 Sales'].");
            }

            var name = byTable.Groups["quoted"].Success
                ? byTable.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal)
                : byTable.Groups["bare"].Value;
            return new ExcelTarget { Kind = ExcelTargetKind.Table, Sheet = sheetTarget.Sheet, TableName = name };
        }

        // /Sheet/slicer[@name=X] gets the same id-form peel-off as pivots/tables.
        var bySlicer = SlicerByName().Match(pathText);
        if (bySlicer.Success)
        {
            var sheetTarget = Resolve(workbook, bySlicer.Groups["sheet"].Value);
            if (sheetTarget.Kind != ExcelTargetKind.Sheet)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"slicer[@name=…] must follow a sheet name: {pathText}",
                    "Use /SheetName/slicer[@name=X]; quote names with specials: slicer[@name='Slicer Region'].");
            }

            var name = bySlicer.Groups["quoted"].Success
                ? bySlicer.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal)
                : bySlicer.Groups["bare"].Value;
            return new ExcelTarget { Kind = ExcelTargetKind.Slicer, Sheet = sheetTarget.Sheet, SlicerName = name };
        }

        // /Sheet/scenario[@name=X] gets the same id-form peel-off (1.5).
        var byScenario = ScenarioByName().Match(pathText);
        if (byScenario.Success)
        {
            var sheetTarget = Resolve(workbook, byScenario.Groups["sheet"].Value);
            if (sheetTarget.Kind != ExcelTargetKind.Sheet)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"scenario[@name=…] must follow a sheet name: {pathText}",
                    "Use /SheetName/scenario[@name=X]; quote names with specials: scenario[@name='Best Case'].");
            }

            var name = byScenario.Groups["quoted"].Success
                ? byScenario.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal)
                : byScenario.Groups["bare"].Value;
            return new ExcelTarget { Kind = ExcelTargetKind.Scenario, Sheet = sheetTarget.Sheet, ScenarioName = name };
        }

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
                $"xlsx paths have at most two segments (sheet, then cell/range/element): {pathText}",
                "Use /Sheet1/A1, /Sheet1/A1:C10, /Sheet1/row[3], /Sheet1/chart[1], /Sheet1/pivot[1], " +
                "/Sheet1/conditionalFormat[1] or /Sheet1/image[1].",
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

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "col", StringComparison.OrdinalIgnoreCase) &&
                segment.Letter is { } columnLetter:
                return new ExcelTarget
                {
                    Kind = ExcelTargetKind.Column,
                    Sheet = sheet,
                    ColumnNumber = new CellRef(columnLetter, 1).ColumnNumber,
                };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "col", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } columnNumber:
                // Numeric col index is accepted as a convenience; the canonical
                // form aioffice emits is the letter form, col[C].
                return new ExcelTarget { Kind = ExcelTargetKind.Column, Sheet = sheet, ColumnNumber = columnNumber };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "chart", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } chartIndex:
                return new ExcelTarget { Kind = ExcelTargetKind.Chart, Sheet = sheet, ChartIndex = chartIndex };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "pivot", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } pivotIndex:
                return new ExcelTarget { Kind = ExcelTargetKind.Pivot, Sheet = sheet, PivotIndex = pivotIndex };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "conditionalFormat", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } formatIndex:
                return new ExcelTarget
                {
                    Kind = ExcelTargetKind.ConditionalFormat,
                    Sheet = sheet,
                    ConditionalFormatIndex = formatIndex,
                };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "image", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } imageIndex:
                return new ExcelTarget { Kind = ExcelTargetKind.Image, Sheet = sheet, ImageIndex = imageIndex };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "linkedPicture", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } linkedPictureIndex:
                return new ExcelTarget
                {
                    Kind = ExcelTargetKind.LinkedPicture,
                    Sheet = sheet,
                    LinkedPictureIndex = linkedPictureIndex,
                };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "dataValidation", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } validationIndex:
                return new ExcelTarget
                {
                    Kind = ExcelTargetKind.DataValidation,
                    Sheet = sheet,
                    DataValidationIndex = validationIndex,
                };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "sparkline", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } sparklineIndex:
                return new ExcelTarget { Kind = ExcelTargetKind.Sparkline, Sheet = sheet, SparklineIndex = sparklineIndex };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "comment", StringComparison.OrdinalIgnoreCase) &&
                segment is { Id: { } commentId, IdAttribute: "id" }:
                return new ExcelTarget { Kind = ExcelTargetKind.Comment, Sheet = sheet, CommentId = commentId };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "slicer", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } slicerIndex:
                return new ExcelTarget { Kind = ExcelTargetKind.Slicer, Sheet = sheet, SlicerIndex = slicerIndex };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "embed", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } embedIndex:
                return new ExcelTarget { Kind = ExcelTargetKind.Embed, Sheet = sheet, EmbedIndex = embedIndex };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "formControl", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } formControlIndex:
                return new ExcelTarget
                {
                    Kind = ExcelTargetKind.FormControl,
                    Sheet = sheet,
                    FormControlIndex = formControlIndex,
                };

            case PathSegmentKind.Element when
                string.Equals(segment.Name, "dataTable", StringComparison.OrdinalIgnoreCase) &&
                segment.Index is { } dataTableIndex:
                return new ExcelTarget
                {
                    Kind = ExcelTargetKind.DataTable,
                    Sheet = sheet,
                    DataTableIndex = dataTableIndex,
                };

            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"'{segment.ToCanonicalString()}' is not a cell, range, row[n], col[C], chart[n], pivot[n], " +
                    $"conditionalFormat[n], image[n], dataValidation[n], sparkline[n], slicer[n], embed[n] or comment[@id=…] in: {pathText}",
                    "After the sheet name use A1, A1:C10, row[3], col[C], chart[1], pivot[1] (or pivot[@name=X]), " +
                    "conditionalFormat[1], image[1], dataValidation[1], sparkline[1], slicer[1], embed[1] or comment[@id=GUID]; " +
                    "column letters are uppercase.",
                    candidates: ExampleTargets(sheet));
        }
    }

    /// <summary>The canonical stable-name pivot path aioffice emits: <c>/Sheet/pivot[@name=X]</c>.</summary>
    public static string PivotPath(IXLWorksheet sheet, string pivotName) =>
        $"{SheetPath(sheet)}/pivot[@name={QuoteSheet(pivotName)}]";

    /// <summary>The canonical stable-name table path aioffice emits: <c>/Sheet/table[@name=X]</c>.</summary>
    public static string TablePath(IXLWorksheet sheet, string tableName) =>
        $"{SheetPath(sheet)}/table[@name={QuoteSheet(tableName)}]";

    public static string ConditionalFormatPath(IXLWorksheet sheet, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"{SheetPath(sheet)}/conditionalFormat[{index}]");

    public static string ImagePath(IXLWorksheet sheet, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"{SheetPath(sheet)}/image[{index}]");

    public static string LinkedPicturePath(IXLWorksheet sheet, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"{SheetPath(sheet)}/linkedPicture[{index}]");

    public static string DataValidationPath(IXLWorksheet sheet, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"{SheetPath(sheet)}/dataValidation[{index}]");

    public static string SparklinePath(IXLWorksheet sheet, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"{SheetPath(sheet)}/sparkline[{index}]");

    public static string SlicerPath(IXLWorksheet sheet, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"{SheetPath(sheet)}/slicer[{index}]");

    public static string EmbedPath(IXLWorksheet sheet, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"{SheetPath(sheet)}/embed[{index}]");

    public static string FormControlPath(IXLWorksheet sheet, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"{SheetPath(sheet)}/formControl[{index}]");

    public static string DataTablePath(IXLWorksheet sheet, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"{SheetPath(sheet)}/dataTable[{index}]");

    /// <summary>The canonical column path aioffice emits: <c>/Sheet1/col[C]</c> (letter form).</summary>
    public static string ColumnPath(IXLWorksheet sheet, int columnNumber) =>
        $"{SheetPath(sheet)}/col[{ExcelCharts.ColumnLetters(columnNumber)}]";

    public static string RowPath(IXLWorksheet sheet, int rowNumber) =>
        string.Create(CultureInfo.InvariantCulture, $"{SheetPath(sheet)}/row[{rowNumber}]");

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
        return
        [
            basePath + "/A1", basePath + "/A1:C10", basePath + "/row[1]", basePath + "/col[A]",
            basePath + "/chart[1]", basePath + "/pivot[1]", basePath + "/conditionalFormat[1]",
            basePath + "/image[1]", basePath + "/dataValidation[1]", basePath + "/sparkline[1]",
        ];
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
