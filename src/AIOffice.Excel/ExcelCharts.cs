using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>One chart series captured at op time: cached values plus the workbook references.</summary>
internal sealed record ChartSeriesSpec(
    string Name,
    string? NameRef,
    IReadOnlyList<double?> Values,
    string ValuesRef);

/// <summary>
/// A fully validated chart-add, extracted from the in-memory workbook while the
/// edit batch is applied. The raw OpenXml write happens later, in the post-save
/// pass (<see cref="ExcelCharts.Apply"/>), because ClosedXML cannot author
/// chart parts itself.
/// </summary>
internal sealed record ChartAddSpec(
    string SheetName,
    string Kind,
    string? Title,
    string DataRange,
    string Anchor,
    IReadOnlyList<string> Categories,
    string CategoriesRef,
    IReadOnlyList<ChartSeriesSpec> Series,
    int AnchorColumn,
    int AnchorRow,
    int WidthCells,
    int HeightCells,
    // Numeric X values (scatter only): the first dataRange column parsed as numbers.
    IReadOnlyList<double?>? XValues = null);

/// <summary>A validated chart removal (1-based per-sheet index at apply time).</summary>
internal sealed record ChartRemoveSpec(string SheetName, int Index);

/// <summary>A chart as read back from the raw package (for get / read --view structure).</summary>
internal sealed record ChartInfo(
    string SheetName,
    int Index,
    string Path,
    string Kind,
    string? Title,
    string? DataRange,
    string? Anchor,
    int Series);

/// <summary>
/// Collects chart ops during an edit batch so they can be validated up front
/// (atomicity: a bad op aborts before any byte is written) and applied after
/// ClosedXML has saved. Tracks projected per-sheet chart counts so indices in
/// a multi-op batch stay consistent.
/// </summary>
internal sealed class ChartOpBatch
{
    private readonly string _file;
    private Dictionary<string, int>? _projectedCounts;

    public ChartOpBatch(string file) => _file = file;

    /// <summary>Chart ops in batch order (<see cref="ChartAddSpec"/> / <see cref="ChartRemoveSpec"/>).</summary>
    public List<object> Ops { get; } = [];

    public bool IsEmpty => Ops.Count == 0;

    /// <summary>Queues an add and returns the chart's projected 1-based index on its sheet.</summary>
    public int Add(ChartAddSpec spec)
    {
        var counts = ProjectedCounts();
        var next = counts.GetValueOrDefault(spec.SheetName) + 1;
        counts[spec.SheetName] = next;
        Ops.Add(spec);
        return next;
    }

    /// <summary>Queues a removal, validating the index against the projected state.</summary>
    public void Remove(string sheetName, string sheetPath, int index, int opIndex)
    {
        var counts = ProjectedCounts();
        var current = counts.GetValueOrDefault(sheetName);
        if (index > current)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"ops[{opIndex}]: no chart[{index}] on sheet '{sheetName}' ({current} chart(s) exist).",
                current > 0
                    ? "Chart indices are 1-based per sheet; run 'aioffice read --view structure' to list them."
                    : "This sheet has no charts; add one with {op:add, type:chart, path:" + sheetPath + ", props:{kind:\"bar\", dataRange:\"A1:B5\", anchor:\"D2\"}}.",
                candidates: current > 0
                    ? [.. Enumerable.Range(1, current).Select(i => $"{sheetPath}/chart[{i}]")]
                    : [sheetPath]);
        }

        counts[sheetName] = current - 1;
        Ops.Add(new ChartRemoveSpec(sheetName, index));
    }

    private Dictionary<string, int> ProjectedCounts()
    {
        if (_projectedCounts is null)
        {
            _projectedCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            using var document = SpreadsheetDocument.Open(_file, isEditable: false);
            foreach (var info in ExcelCharts.Read(document))
            {
                _projectedCounts[info.SheetName] = info.Index; // indices are sequential per sheet
            }
        }

        return _projectedCounts;
    }
}

/// <summary>
/// The xlsx chart layer. ClosedXML cannot author charts, so this slice works on
/// raw OpenXml: <see cref="ParseAdd"/> validates an add op and captures cached
/// data from the in-memory workbook; <see cref="Apply"/> writes
/// ChartPart/DrawingsPart XML in a second pass over the file ClosedXML saved
/// (measured: ClosedXML 0.105 preserves these parts byte-identical on later
/// saves, so charts survive subsequent edits).
/// </summary>
internal static partial class ExcelCharts
{
    /// <summary>The chart kinds aioffice can create. Everything else is unsupported_feature.</summary>
    public static readonly IReadOnlyList<string> Kinds = ["bar", "line", "pie", "scatter", "area"];

    private static readonly IReadOnlyList<string> AddProps =
        ["kind", "dataRange", "anchor", "title", "widthCells", "heightCells"];

    private const string ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string MainNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string SpreadsheetDrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    private const int MaxColumns = 16384; // XFD
    private const int MaxRows = 1048576;

    private const uint CategoryAxisId = 100001u;
    private const uint ValueAxisId = 100002u;

    [GeneratedRegex("^([A-Z]{1,3})([0-9]{1,7})$")]
    private static partial Regex CellPattern();

    [GeneratedRegex("^([A-Z]{1,3})([0-9]{1,7}):([A-Z]{1,3})([0-9]{1,7})$")]
    private static partial Regex RangePattern();

    /// <summary>Parses chart formula refs like <c>'Q3 Data'!$A$2:$A$5</c> or <c>Sheet1!$B$1</c>.</summary>
    [GeneratedRegex(@"^(?:'(?<q>(?:[^']|'')+)'|(?<bare>[^'!]+))!\$?(?<c1>[A-Z]{1,3})\$?(?<r1>[0-9]{1,7})(?::\$?(?<c2>[A-Z]{1,3})\$?(?<r2>[0-9]{1,7}))?$")]
    private static partial Regex FormulaRefPattern();

    // ----- op-time validation & data capture ---------------------------------

    /// <summary>
    /// Validates an <c>add chart</c> op against the in-memory workbook and
    /// captures everything the post-save writer needs. Data semantics: the
    /// first column of <c>dataRange</c> is categories; each later column is a
    /// series; the first row supplies series names when every series cell in
    /// it is text.
    /// </summary>
    public static ChartAddSpec ParseAdd(IXLWorksheet sheet, EditOp op, int opIndex)
    {
        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add chart needs props.",
                "Pass props like {\"kind\":\"bar\",\"dataRange\":\"A1:B5\",\"anchor\":\"D2\"}.");
        }

        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown chart prop '{key}'.",
                    "Supported chart props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }

        var kind = RequiredString(props, "kind", opIndex, "bar");
        if (!Kinds.Contains(kind, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{opIndex}]: chart kind '{kind}' is not supported.",
                "Supported chart kinds: bar, line, pie, scatter, area. Bubble, radar and combo " +
                "charts land later; a scatter chart is the usual stand-in for bubble data.",
                candidates: Kinds);
        }

        var (firstColumn, firstRow, lastColumn, lastRow) =
            ParseDataRange(RequiredString(props, "dataRange", opIndex, "A1:B5"), opIndex);
        var columns = lastColumn - firstColumn + 1;
        if (columns < 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: dataRange needs at least two columns (categories, then one or more series).",
                "Widen the range, e.g. A1:B5 — first column categories, later columns series values.");
        }

        if (kind == "pie" && columns != 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a pie chart takes exactly one series (two columns); dataRange has {columns}.",
                "Narrow dataRange to categories plus one value column, or use kind bar/line for multi-series data.");
        }

        var (anchorColumn, anchorRow) = ParseAnchor(RequiredString(props, "anchor", opIndex, "D2"), opIndex);
        var widthCells = OptionalInt(props, "widthCells", opIndex) ?? 8;
        var heightCells = OptionalInt(props, "heightCells", opIndex) ?? 15;
        if (widthCells < 1 || heightCells < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: widthCells and heightCells must be at least 1.",
                "Defaults are widthCells 8 and heightCells 15; omit them unless you need another size.");
        }

        if (anchorColumn - 1 + widthCells > MaxColumns || anchorRow - 1 + heightCells > MaxRows)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: the chart would extend past the sheet edge.",
                "Move the anchor up/left or shrink widthCells/heightCells.");
        }

        // First row = series names only when EVERY series cell in it is text.
        var hasHeader = true;
        for (var column = firstColumn + 1; column <= lastColumn; column++)
        {
            if (!EvaluatedValue(sheet.Cell(firstRow, column)).IsText)
            {
                hasHeader = false;
                break;
            }
        }

        var dataFirstRow = hasHeader ? firstRow + 1 : firstRow;
        if (dataFirstRow > lastRow)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: dataRange has a header row but no data rows.",
                "Extend the range to include at least one row of values under the header.");
        }

        var categories = new List<string>(lastRow - dataFirstRow + 1);
        for (var row = dataFirstRow; row <= lastRow; row++)
        {
            categories.Add(ExcelValues.SafeFormatted(sheet.Cell(row, firstColumn)));
        }

        // Scatter plots numbers against numbers: the first column is the X
        // axis and must be numeric (categories stay text for bar/line/pie/area).
        List<double?>? xValues = null;
        if (kind == "scatter")
        {
            xValues = new List<double?>(lastRow - dataFirstRow + 1);
            for (var row = dataFirstRow; row <= lastRow; row++)
            {
                var cell = sheet.Cell(row, firstColumn);
                var value = EvaluatedValue(cell);
                xValues.Add(value.Type switch
                {
                    XLDataType.Blank => null,
                    XLDataType.Number => value.GetNumber(),
                    XLDataType.DateTime => value.GetDateTime().ToOADate(),
                    XLDataType.TimeSpan => value.GetTimeSpan().TotalDays,
                    _ => throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{opIndex}]: a scatter chart needs numeric X values, but " +
                        $"{ExcelPaths.CellPath(sheet, cell.Address)} is {ExcelValues.TypeName(value.Type)}.",
                        "Point the first dataRange column at numbers, or use kind \"line\" for category data."),
                });
            }
        }

        var categoriesRef = RangeRef(sheet.Name, firstColumn, dataFirstRow, firstColumn, lastRow);
        var series = new List<ChartSeriesSpec>(columns - 1);
        for (var column = firstColumn + 1; column <= lastColumn; column++)
        {
            var ordinal = column - firstColumn;
            var values = new List<double?>(lastRow - dataFirstRow + 1);
            for (var row = dataFirstRow; row <= lastRow; row++)
            {
                values.Add(NumericChartValue(sheet, sheet.Cell(row, column), opIndex));
            }

            series.Add(new ChartSeriesSpec(
                Name: hasHeader
                    ? ExcelValues.SafeFormatted(sheet.Cell(firstRow, column))
                    : string.Create(CultureInfo.InvariantCulture, $"Series {ordinal}"),
                NameRef: hasHeader ? CellRefText(sheet.Name, column, firstRow) : null,
                Values: values,
                ValuesRef: RangeRef(sheet.Name, column, dataFirstRow, column, lastRow)));
        }

        return new ChartAddSpec(
            SheetName: sheet.Name,
            Kind: kind,
            Title: OptionalString(props, "title"),
            DataRange: string.Create(
                CultureInfo.InvariantCulture,
                $"{ColumnLetters(firstColumn)}{firstRow}:{ColumnLetters(lastColumn)}{lastRow}"),
            Anchor: string.Create(CultureInfo.InvariantCulture, $"{ColumnLetters(anchorColumn)}{anchorRow}"),
            Categories: categories,
            CategoriesRef: categoriesRef,
            Series: series,
            AnchorColumn: anchorColumn,
            AnchorRow: anchorRow,
            WidthCells: widthCells,
            HeightCells: heightCells,
            XValues: xValues);
    }

    private static (int FirstColumn, int FirstRow, int LastColumn, int LastRow) ParseDataRange(
        string text, int opIndex)
    {
        var match = RangePattern().Match(text.ToUpperInvariant());
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{text}' is not a usable dataRange.",
                "dataRange is a plain range on the chart's own sheet, e.g. A1:B5 (no sheet prefix, no path).");
        }

        var start = new CellRef(match.Groups[1].Value, int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture));
        var end = new CellRef(match.Groups[3].Value, int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture));
        if (start.ColumnNumber > end.ColumnNumber || start.Row > end.Row)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: dataRange start must not be past its end: {text}",
                "Write the range top-left to bottom-right, e.g. A1:B5.");
        }

        return (start.ColumnNumber, start.Row, end.ColumnNumber, end.Row);
    }

    private static (int Column, int Row) ParseAnchor(string text, int opIndex)
    {
        var match = CellPattern().Match(text.ToUpperInvariant());
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{text}' is not a usable anchor.",
                "anchor is the top-left cell the chart hangs from, e.g. D2.");
        }

        var cell = new CellRef(match.Groups[1].Value, int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture));
        return (cell.ColumnNumber, cell.Row);
    }

    private static XLCellValue EvaluatedValue(IXLCell cell)
    {
        try
        {
            return cell.Value;
        }
        catch (Exception)
        {
            return cell.CachedValue;
        }
    }

    private static double? NumericChartValue(IXLWorksheet sheet, IXLCell cell, int opIndex)
    {
        var value = EvaluatedValue(cell);
        return value.Type switch
        {
            XLDataType.Blank => null,
            XLDataType.Number => value.GetNumber(),
            XLDataType.DateTime => value.GetDateTime().ToOADate(),
            XLDataType.TimeSpan => value.GetTimeSpan().TotalDays,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: chart series data must be numeric, but " +
                $"{ExcelPaths.CellPath(sheet, cell.Address)} is {ExcelValues.TypeName(value.Type)}.",
                "Only the FIRST dataRange column may hold text (the categories); " +
                "point the later columns at numeric cells."),
        };
    }

    // ----- raw read-back ------------------------------------------------------

    /// <summary>All charts in the package, in sheet order then drawing order (1-based per sheet).</summary>
    public static List<ChartInfo> Read(SpreadsheetDocument document)
    {
        var result = new List<ChartInfo>();
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return result;
        }

        foreach (var sheet in workbookPart.Workbook.Descendants<S.Sheet>())
        {
            if (sheet.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart ||
                worksheetPart.DrawingsPart is not { } drawings)
            {
                continue;
            }

            var sheetName = sheet.Name?.Value ?? string.Empty;
            var index = 0;
            foreach (var (anchor, chartPart) in ChartAnchors(drawings))
            {
                index++;
                result.Add(Describe(sheetName, index, anchor, chartPart));
            }
        }

        return result;
    }

    /// <summary>Drawing anchors that host a chart graphic frame, in document order.</summary>
    private static IEnumerable<(OpenXmlCompositeElement Anchor, ChartPart Part)> ChartAnchors(DrawingsPart drawings)
    {
        if (drawings.WorksheetDrawing is not { } root)
        {
            yield break;
        }

        foreach (var child in root.ChildElements)
        {
            if (child is not (Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor))
            {
                continue;
            }

            var anchor = (OpenXmlCompositeElement)child;
            if (anchor.Descendants<C.ChartReference>().FirstOrDefault()?.Id?.Value is not { } relationshipId)
            {
                continue;
            }

            var pair = drawings.Parts.FirstOrDefault(p => p.RelationshipId == relationshipId);
            if (pair.OpenXmlPart is ChartPart chartPart)
            {
                yield return (anchor, chartPart);
            }
        }
    }

    private static ChartInfo Describe(string sheetName, int index, OpenXmlCompositeElement anchor, ChartPart chartPart)
    {
        var chartSpace = chartPart.ChartSpace;
        var plotArea = chartSpace?.Descendants<C.PlotArea>().FirstOrDefault();
        var group = plotArea?.ChildElements
            .FirstOrDefault(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

        var series = group?.ChildElements.Count(e => e.LocalName == "ser") ?? 0;
        return new ChartInfo(
            SheetName: sheetName,
            Index: index,
            Path: string.Create(
                CultureInfo.InvariantCulture,
                $"/{ExcelPaths.QuoteSheet(sheetName)}/chart[{index}]"),
            Kind: KindName(group?.LocalName),
            Title: TitleText(chartSpace),
            DataRange: DataRangeOf(group),
            Anchor: AnchorRefOf(anchor),
            Series: series);
    }

    private static string KindName(string? localName)
    {
        if (string.IsNullOrEmpty(localName))
        {
            return "unknown";
        }

        var name = localName.EndsWith("Chart", StringComparison.Ordinal) ? localName[..^5] : localName;
        return name.EndsWith("3D", StringComparison.Ordinal) ? name[..^2] : name;
    }

    private static string? TitleText(C.ChartSpace? chartSpace)
    {
        var title = chartSpace?.Descendants<C.Title>().FirstOrDefault();
        if (title is null)
        {
            return null;
        }

        var text = string.Concat(title.Descendants<A.Text>().Select(t => t.Text));
        if (text.Length == 0)
        {
            text = string.Concat(title.Descendants<C.NumericValue>().Select(v => v.Text));
        }

        return text.Length == 0 ? null : text;
    }

    /// <summary>Reconstructs the original dataRange as the union of every series ref on the primary sheet.</summary>
    private static string? DataRangeOf(OpenXmlElement? group)
    {
        if (group is null)
        {
            return null;
        }

        string? primarySheet = null;
        int minColumn = int.MaxValue, minRow = int.MaxValue, maxColumn = 0, maxRow = 0;
        foreach (var formula in group.ChildElements
                     .Where(e => e.LocalName == "ser")
                     .SelectMany(s => s.Descendants<C.Formula>()))
        {
            var match = FormulaRefPattern().Match(formula.Text);
            if (!match.Success)
            {
                continue;
            }

            var sheetOfRef = match.Groups["q"].Success
                ? match.Groups["q"].Value.Replace("''", "'", StringComparison.Ordinal)
                : match.Groups["bare"].Value;
            primarySheet ??= sheetOfRef;
            if (!string.Equals(primarySheet, sheetOfRef, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var start = new CellRef(
                match.Groups["c1"].Value,
                int.Parse(match.Groups["r1"].Value, CultureInfo.InvariantCulture));
            var end = match.Groups["c2"].Success
                ? new CellRef(match.Groups["c2"].Value, int.Parse(match.Groups["r2"].Value, CultureInfo.InvariantCulture))
                : start;
            minColumn = Math.Min(minColumn, start.ColumnNumber);
            minRow = Math.Min(minRow, start.Row);
            maxColumn = Math.Max(maxColumn, end.ColumnNumber);
            maxRow = Math.Max(maxRow, end.Row);
        }

        if (primarySheet is null)
        {
            return null;
        }

        return string.Create(
            CultureInfo.InvariantCulture,
            $"{ColumnLetters(minColumn)}{minRow}:{ColumnLetters(maxColumn)}{maxRow}");
    }

    private static string? AnchorRefOf(OpenXmlCompositeElement anchor)
    {
        var from = anchor switch
        {
            Xdr.TwoCellAnchor two => two.FromMarker,
            Xdr.OneCellAnchor one => one.FromMarker,
            _ => null,
        };
        if (from?.ColumnId?.Text is not { } columnText || from.RowId?.Text is not { } rowText ||
            !int.TryParse(columnText, NumberStyles.None, CultureInfo.InvariantCulture, out var column) ||
            !int.TryParse(rowText, NumberStyles.None, CultureInfo.InvariantCulture, out var row))
        {
            return null;
        }

        return string.Create(CultureInfo.InvariantCulture, $"{ColumnLetters(column + 1)}{row + 1}");
    }

    // ----- post-save apply ----------------------------------------------------

    /// <summary>
    /// Applies queued chart ops to the file ClosedXML just saved. All semantic
    /// validation already happened at op time, so this pass is mechanical.
    /// </summary>
    public static void Apply(string file, IReadOnlyList<object> ops)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        foreach (var op in ops)
        {
            switch (op)
            {
                case ChartAddSpec add:
                    AddChart(document, add);
                    break;
                case ChartRemoveSpec remove:
                    RemoveChart(document, remove);
                    break;
            }
        }
    }

    private static WorksheetPart SheetPartOrThrow(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is { } relationshipId &&
            workbookPart!.GetPartById(relationshipId) is WorksheetPart worksheetPart)
        {
            return worksheetPart;
        }

        throw new AiofficeException(
            ErrorCodes.InternalError,
            $"Sheet '{sheetName}' disappeared between validation and the chart write pass.",
            "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
    }

    private static void AddChart(SpreadsheetDocument document, ChartAddSpec spec)
    {
        var worksheetPart = SheetPartOrThrow(document, spec.SheetName);
        var drawings = worksheetPart.DrawingsPart;
        if (drawings is null)
        {
            drawings = worksheetPart.AddNewPart<DrawingsPart>();
        }

        if (drawings.WorksheetDrawing is null)
        {
            var root = new Xdr.WorksheetDrawing();
            root.AddNamespaceDeclaration("xdr", SpreadsheetDrawingNs);
            root.AddNamespaceDeclaration("a", MainNs);
            drawings.WorksheetDrawing = root;
        }

        EnsureDrawingElement(worksheetPart, drawings);

        var chartPart = drawings.AddNewPart<ChartPart>();
        chartPart.ChartSpace = BuildChartSpace(spec);
        AppendAnchor(drawings, chartPart, spec);
    }

    /// <summary>
    /// Makes sure the worksheet references its drawings part. The schema puts
    /// <c>drawing</c> before <c>tableParts</c>/<c>extLst</c>, and ClosedXML
    /// re-emits it in the correct slot on later saves (measured).
    /// </summary>
    private static void EnsureDrawingElement(WorksheetPart worksheetPart, DrawingsPart drawings)
    {
        if (worksheetPart.Worksheet is not { } worksheet)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                "A worksheet part has no worksheet XML.",
                "Restore a snapshot ('aioffice snapshot list') or re-export the file from its source.");
        }

        if (worksheet.Elements<S.Drawing>().Any())
        {
            return;
        }

        var drawing = new S.Drawing { Id = worksheetPart.GetIdOfPart(drawings) };
        var successor = worksheet.Elements<OpenXmlElement>().FirstOrDefault(e =>
            e is S.LegacyDrawing or S.LegacyDrawingHeaderFooter or S.DrawingHeaderFooter or S.Picture
                or S.OleObjects or S.Controls or S.WebPublishItems or S.TableParts or S.ExtensionList);
        if (successor is null)
        {
            worksheet.Append(drawing);
        }
        else
        {
            worksheet.InsertBefore(drawing, successor);
        }
    }

    private static C.ChartSpace BuildChartSpace(ChartAddSpec spec)
    {
        var plotArea = new C.PlotArea(new C.Layout());
        plotArea.Append(BuildChartGroup(spec));
        if (spec.Kind is "bar" or "line" or "area")
        {
            plotArea.Append(new C.CategoryAxis(
                new C.AxisId { Val = CategoryAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                new C.CrossingAxis { Val = ValueAxisId }));
            plotArea.Append(new C.ValueAxis(
                new C.AxisId { Val = ValueAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Left },
                new C.CrossingAxis { Val = CategoryAxisId }));
        }
        else if (spec.Kind == "scatter")
        {
            // Scatter plots value-against-value: both axes are value axes.
            plotArea.Append(new C.ValueAxis(
                new C.AxisId { Val = CategoryAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                new C.CrossingAxis { Val = ValueAxisId }));
            plotArea.Append(new C.ValueAxis(
                new C.AxisId { Val = ValueAxisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Left },
                new C.CrossingAxis { Val = CategoryAxisId }));
        }

        var chart = new C.Chart();
        if (spec.Title is { } titleText)
        {
            chart.Append(new C.Title(
                new C.ChartText(new C.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(titleText))))),
                new C.Overlay { Val = false }));
        }

        chart.Append(new C.AutoTitleDeleted { Val = spec.Title is null });
        chart.Append(plotArea);
        chart.Append(new C.PlotVisibleOnly { Val = true });

        var chartSpace = new C.ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", ChartNs);
        chartSpace.AddNamespaceDeclaration("a", MainNs);
        chartSpace.AddNamespaceDeclaration("r", RelNs);
        chartSpace.Append(chart);
        return chartSpace;
    }

    private static OpenXmlCompositeElement BuildChartGroup(ChartAddSpec spec)
    {
        switch (spec.Kind)
        {
            case "bar":
            {
                var group = new C.BarChart(
                    new C.BarDirection { Val = C.BarDirectionValues.Column },
                    new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
                    new C.VaryColors { Val = false });
                for (var i = 0; i < spec.Series.Count; i++)
                {
                    var series = new C.BarChartSeries();
                    AppendSeriesChildren(series, spec, i);
                    group.Append(series);
                }

                group.Append(new C.AxisId { Val = CategoryAxisId });
                group.Append(new C.AxisId { Val = ValueAxisId });
                return group;
            }

            case "line":
            {
                var group = new C.LineChart(
                    new C.Grouping { Val = C.GroupingValues.Standard },
                    new C.VaryColors { Val = false });
                for (var i = 0; i < spec.Series.Count; i++)
                {
                    var series = new C.LineChartSeries();
                    AppendSeriesChildren(series, spec, i);
                    group.Append(series);
                }

                group.Append(new C.AxisId { Val = CategoryAxisId });
                group.Append(new C.AxisId { Val = ValueAxisId });
                return group;
            }

            case "area":
            {
                var group = new C.AreaChart(
                    new C.Grouping { Val = C.GroupingValues.Standard },
                    new C.VaryColors { Val = false });
                for (var i = 0; i < spec.Series.Count; i++)
                {
                    var series = new C.AreaChartSeries();
                    AppendSeriesChildren(series, spec, i);
                    group.Append(series);
                }

                group.Append(new C.AxisId { Val = CategoryAxisId });
                group.Append(new C.AxisId { Val = ValueAxisId });
                return group;
            }

            case "scatter":
            {
                var group = new C.ScatterChart(
                    new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
                    new C.VaryColors { Val = false });
                for (var i = 0; i < spec.Series.Count; i++)
                {
                    group.Append(BuildScatterSeries(spec, i));
                }

                group.Append(new C.AxisId { Val = CategoryAxisId });
                group.Append(new C.AxisId { Val = ValueAxisId });
                return group;
            }

            default: // "pie" — ParseAdd rejected everything else
            {
                var group = new C.PieChart(new C.VaryColors { Val = true });
                var series = new C.PieChartSeries();
                AppendSeriesChildren(series, spec, 0);
                group.Append(series);
                group.Append(new C.FirstSliceAngle { Val = 0 });
                return group;
            }
        }
    }

    /// <summary>idx/order/tx/spPr/xVal/yVal/smooth — scatter plots X numbers against Y numbers.</summary>
    private static C.ScatterChartSeries BuildScatterSeries(ChartAddSpec spec, int ordinal)
    {
        var data = spec.Series[ordinal];
        var series = new C.ScatterChartSeries();
        series.Append(new C.Index { Val = (uint)ordinal });
        series.Append(new C.Order { Val = (uint)ordinal });
        series.Append(SeriesText(data));

        // Markers only: no connecting line, the classic scatter look.
        series.Append(new C.ChartShapeProperties(new A.Outline(new A.NoFill())));

        series.Append(new C.XValues(new C.NumberReference(
            new C.Formula(spec.CategoriesRef), NumberCache(spec.XValues!))));
        series.Append(new C.YValues(new C.NumberReference(
            new C.Formula(data.ValuesRef), NumberCache(data.Values))));
        series.Append(new C.Smooth { Val = false });
        return series;
    }

    /// <summary>idx/order/tx/cat/val — the shared series skeleton (schema order).</summary>
    private static void AppendSeriesChildren(OpenXmlCompositeElement series, ChartAddSpec spec, int ordinal)
    {
        var data = spec.Series[ordinal];
        series.Append(new C.Index { Val = (uint)ordinal });
        series.Append(new C.Order { Val = (uint)ordinal });
        series.Append(SeriesText(data));

        var stringCache = new C.StringCache(new C.PointCount { Val = (uint)spec.Categories.Count });
        for (var i = 0; i < spec.Categories.Count; i++)
        {
            stringCache.Append(new C.StringPoint
            {
                Index = (uint)i,
                NumericValue = new C.NumericValue(spec.Categories[i]),
            });
        }

        series.Append(new C.CategoryAxisData(new C.StringReference(
            new C.Formula(spec.CategoriesRef), stringCache)));

        series.Append(new C.Values(new C.NumberReference(
            new C.Formula(data.ValuesRef), NumberCache(data.Values))));
    }

    private static C.SeriesText SeriesText(ChartSeriesSpec data) => data.NameRef is { } nameRef
        ? new C.SeriesText(new C.StringReference(
            new C.Formula(nameRef),
            new C.StringCache(
                new C.PointCount { Val = 1u },
                new C.StringPoint { Index = 0u, NumericValue = new C.NumericValue(data.Name) })))
        : new C.SeriesText(new C.NumericValue(data.Name));

    private static C.NumberingCache NumberCache(IReadOnlyList<double?> values)
    {
        var cache = new C.NumberingCache(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)values.Count });
        for (var i = 0; i < values.Count; i++)
        {
            if (values[i] is { } value)
            {
                cache.Append(new C.NumericPoint
                {
                    Index = (uint)i,
                    NumericValue = new C.NumericValue(value.ToString(CultureInfo.InvariantCulture)),
                });
            }
        }

        return cache;
    }

    private static void AppendAnchor(DrawingsPart drawings, ChartPart chartPart, ChartAddSpec spec)
    {
        var root = drawings.WorksheetDrawing!;
        var frameId = root.Descendants<Xdr.NonVisualDrawingProperties>()
            .Select(p => p.Id?.Value ?? 0u)
            .DefaultIfEmpty(1u)
            .Max() + 1;

        var fromColumn = spec.AnchorColumn - 1;
        var fromRow = spec.AnchorRow - 1;
        var anchor = new Xdr.TwoCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId(Invariant(fromColumn)),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId(Invariant(fromRow)),
                new Xdr.RowOffset("0")),
            new Xdr.ToMarker(
                new Xdr.ColumnId(Invariant(fromColumn + spec.WidthCells)),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId(Invariant(fromRow + spec.HeightCells)),
                new Xdr.RowOffset("0")));
        anchor.Append(new Xdr.GraphicFrame(
            new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties
                {
                    Id = frameId,
                    Name = string.Create(CultureInfo.InvariantCulture, $"Chart {frameId}"),
                },
                new Xdr.NonVisualGraphicFrameDrawingProperties()),
            new Xdr.Transform(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 0, Cy = 0 }),
            new A.Graphic(new A.GraphicData(
                new C.ChartReference { Id = drawings.GetIdOfPart(chartPart) })
            {
                Uri = ChartNs,
            })));
        anchor.Append(new Xdr.ClientData());
        root.Append(anchor);
    }

    private static void RemoveChart(SpreadsheetDocument document, ChartRemoveSpec spec)
    {
        var worksheetPart = SheetPartOrThrow(document, spec.SheetName);
        var drawings = worksheetPart.DrawingsPart;
        var anchors = drawings is null ? [] : ChartAnchors(drawings).ToList();
        if (drawings is null || spec.Index > anchors.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"chart[{spec.Index}] on '{spec.SheetName}' disappeared between validation and the chart write pass.",
                "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
        }

        var (anchor, chartPart) = anchors[spec.Index - 1];
        anchor.Remove();
        drawings.DeletePart(chartPart);

        var anchorsLeft = drawings.WorksheetDrawing?.ChildElements
            .Any(c => c is Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor) ?? false;
        if (!anchorsLeft)
        {
            var relationshipId = worksheetPart.GetIdOfPart(drawings);
            var drawingElements = worksheetPart.Worksheet?.Elements<S.Drawing>()
                .Where(d => d.Id?.Value == relationshipId).ToList() ?? [];
            foreach (var drawing in drawingElements)
            {
                drawing.Remove();
            }

            worksheetPart.DeletePart(drawings);
        }
    }

    // ----- small shared helpers ----------------------------------------------

    private static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);

    /// <summary>1-based column number to letters (1=A, 27=AA). Oversized for malformed foreign refs.</summary>
    internal static string ColumnLetters(int column)
    {
        Span<char> buffer = stackalloc char[8];
        var i = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--i] = (char)('A' + (column % 26));
            column /= 26;
        }

        return new string(buffer[i..]);
    }

    /// <summary>
    /// Chart formulas always quote the sheet name — quoted refs are valid for
    /// every name, which sidesteps Excel's bare-name rules entirely.
    /// </summary>
    private static string SheetRef(string sheetName) =>
        "'" + sheetName.Replace("'", "''", StringComparison.Ordinal) + "'";

    private static string RangeRef(string sheetName, int firstColumn, int firstRow, int lastColumn, int lastRow) =>
        string.Create(
            CultureInfo.InvariantCulture,
            $"{SheetRef(sheetName)}!${ColumnLetters(firstColumn)}${firstRow}:${ColumnLetters(lastColumn)}${lastRow}");

    private static string CellRefText(string sheetName, int column, int row) =>
        string.Create(
            CultureInfo.InvariantCulture,
            $"{SheetRef(sheetName)}!${ColumnLetters(column)}${row}");

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
            $"ops[{opIndex}]: add chart needs the '{key}' prop.",
            $"Pass it as a string, e.g. {{\"{key}\":\"{example}\"}}.");
    }

    private static string? OptionalString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static int? OptionalInt(JsonObject props, string key, int opIndex)
    {
        if (!props.TryGetPropertyValue(key, out var node) || node is null)
        {
            return null;
        }

        if (node is JsonValue value)
        {
            if (value.GetValueKind() == JsonValueKind.Number && value.TryGetValue<int>(out var number))
            {
                return number;
            }

            if (value.GetValueKind() == JsonValueKind.String &&
                int.TryParse(value.GetValue<string>(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: '{key}' must be a whole number of cells.",
            $"Pass e.g. {{\"{key}\":8}}.");
    }
}
