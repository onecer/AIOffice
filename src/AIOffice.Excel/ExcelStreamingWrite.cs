using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// The M6 flagship: an IN-PLACE streaming rewrite that edits a LARGE existing
/// workbook without loading the whole DOM. It handles the single most common
/// mutation — setting a cell value or formula (single cell and bulk 2D) on a
/// target sheet — by streaming the source sheet part through an
/// <see cref="OpenXmlReader"/> into an <see cref="OpenXmlWriter"/> for a fresh
/// part, splicing the targeted cells in in row/column order, then atomically
/// swapping the part.
///
/// <para>Activation: file size over <see cref="ThresholdBytes"/> (default
/// 15 MB) OR the explicit <c>stream=true</c> arg — but ONLY when EVERY op in the
/// batch is a streamable cell/range write (see <see cref="TryPlan"/>). Anything
/// else (a chart add, a style-only set, a defined-name op…) transparently falls
/// back to the full ClosedXML DOM path; the user is never told, because both
/// paths produce the same workbook. A 50 MB equality test pins that.</para>
///
/// <para>Streaming-capable ops (documented in PARITY):</para>
/// <list type="bullet">
/// <item><c>set value</c> on a single cell (number/boolean/text/blank/formula
/// or ISO date);</item>
/// <item><c>set values</c> bulk 2D at a cell anchor or an exact range, same
/// value vocabulary.</item>
/// </list>
/// Other set props (bold/fill/numberFormat/merge/…) make the op non-streamable
/// → DOM fallback. Dates ARE streamable here (unlike the bare-sheet bulk path):
/// the existing styles part already carries number formats, so a date cell is
/// written with the workbook's default date style id (cloned from any existing
/// date-formatted cell, else a style is appended).
///
/// <para>Formula cells written via streaming carry NO cached value (the SAX
/// pass does not evaluate). The edit sets the workbook's <c>fullCalcOnLoad</c>
/// recalc flag so Excel recomputes on open, and a <c>formula_not_evaluated</c>
/// warning is raised — the honest behavior, pinned by a test.</para>
/// </summary>
internal static class ExcelStreamingWrite
{
    /// <summary>Files larger than this take the in-place streaming write path (when the batch is streamable).</summary>
    public const long ThresholdBytes = 15L * 1024 * 1024;

    public static bool ShouldConsider(string file, bool streamArg) =>
        streamArg || new FileInfo(file).Length > ThresholdBytes;

    /// <summary>One cell write: a value/formula at a 1-based (row, column).</summary>
    internal readonly record struct CellWrite(int Row, int Column, ExcelValues.ParsedValue Value);

    /// <summary>A planned streaming edit: every targeted cell, grouped per sheet, in batch order.</summary>
    internal sealed class Plan
    {
        /// <summary>Per-sheet merged writes: sheet name → (row → (column → value)). Later ops win.</summary>
        public Dictionary<string, SortedDictionary<int, SortedDictionary<int, ExcelValues.ParsedValue>>> BySheet { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        public bool HasFormulas { get; set; }

        public bool HasDates { get; set; }

        public void Add(string sheetName, int row, int column, ExcelValues.ParsedValue value)
        {
            if (!BySheet.TryGetValue(sheetName, out var rows))
            {
                BySheet[sheetName] = rows = new SortedDictionary<int, SortedDictionary<int, ExcelValues.ParsedValue>>();
            }

            if (!rows.TryGetValue(row, out var cells))
            {
                rows[row] = cells = [];
            }

            cells[column] = value;
            if (value.IsFormula)
            {
                HasFormulas = true;
            }
            else if (value.Value.Type is ClosedXML.Excel.XLDataType.DateTime or ClosedXML.Excel.XLDataType.TimeSpan)
            {
                HasDates = true;
            }
        }
    }

    /// <summary>
    /// Tries to plan the whole batch as a streaming edit. Returns null (DOM
    /// fallback) the moment any op is not a streamable cell/range write, or
    /// addresses a sheet/cell the SAX path can't safely resolve. A null plan is
    /// NOT an error — the caller silently runs the DOM path.
    /// </summary>
    public static Plan? TryPlan(string file, IReadOnlyList<EditOp> ops)
    {
        // Discover sheet names without a DOM load. A corrupt/unopenable package
        // falls back to the DOM path, which raises the canonical format_corrupt.
        HashSet<string> sheetNames;
        try
        {
            using var document = SpreadsheetDocument.Open(file, isEditable: false);
            var workbookPart = document.WorkbookPart;
            if (workbookPart?.Workbook is null)
            {
                return null;
            }

            sheetNames = workbookPart.Workbook
                .Descendants<S.Sheet>()
                .Select(s => s.Name?.Value)
                .Where(n => n is not null)
                .Select(n => n!)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }
        catch (Exception)
        {
            return null;
        }

        var plan = new Plan();
        foreach (var op in ops)
        {
            // A malformed op (bad valueType, unparseable date…) raises here in
            // planning; let the DOM path produce the canonical typed error.
            try
            {
                if (!PlanOp(op, sheetNames, plan))
                {
                    return null;
                }
            }
            catch (AiofficeException)
            {
                return null;
            }
        }

        return plan.BySheet.Count == 0 ? null : plan;
    }

    private static bool PlanOp(EditOp op, HashSet<string> sheetNames, Plan plan)
    {
        if (!string.Equals(op.Op, "set", StringComparison.Ordinal))
        {
            return false; // only `set` streams; add/remove/replace/move take the DOM
        }

        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            return false;
        }

        // The only streamable props are value (+valueType) and values. Anything
        // else (bold, fill, merge, numberFormat, hyperlink, freeze, …) → DOM.
        foreach (var (key, _) in props)
        {
            if (key is not ("value" or "values" or "valueType"))
            {
                return false;
            }
        }

        // A path that is not a plain /Sheet/Cell or /Sheet/Range can't stream
        // (named ranges, row/col/element targets, id-forms). Resolve cheaply.
        if (!TryParseSheetCellOrRange(op.Path, out var sheetName, out var start, out var end))
        {
            return false;
        }

        if (!sheetNames.TryGetValue(sheetName, out var canonicalSheet))
        {
            return false; // unknown sheet → let the DOM path raise the proper invalid_path
        }

        if (props.TryGetPropertyValue("value", out var valueNode))
        {
            if (props.ContainsKey("values"))
            {
                return false; // value + values together: hand to the DOM path's validation
            }

            var parsed = ExcelValues.Parse(valueNode, ValueTypeArg(props));
            for (var r = start.Row; r <= end.Row; r++)
            {
                for (var c = start.ColumnNumber; c <= end.ColumnNumber; c++)
                {
                    plan.Add(canonicalSheet, r, c, parsed);
                }
            }

            return true;
        }

        if (props.TryGetPropertyValue("values", out var valuesNode))
        {
            return PlanValues(valuesNode, canonicalSheet, start, end, plan);
        }

        return false;
    }

    /// <summary>Plans a bulk 2D <c>values</c> write (anchor or exact-range form).</summary>
    private static bool PlanValues(
        System.Text.Json.Nodes.JsonNode? valuesNode, string sheetName, CellRef start, CellRef end, Plan plan)
    {
        if (valuesNode is not System.Text.Json.Nodes.JsonArray array || array.Count == 0)
        {
            return false;
        }

        var isRange = start.Row != end.Row || start.ColumnNumber != end.ColumnNumber;
        var rangeRows = end.Row - start.Row + 1;
        var rangeColumns = end.ColumnNumber - start.ColumnNumber + 1;

        for (var r = 0; r < array.Count; r++)
        {
            if (array[r] is not System.Text.Json.Nodes.JsonArray rowArray)
            {
                return false; // flat array on a cell anchor: let the DOM path raise the 2D hint
            }

            if (isRange && (array.Count != rangeRows || rowArray.Count != rangeColumns))
            {
                return false; // shape mismatch: DOM path produces the exact both-shapes error
            }

            for (var c = 0; c < rowArray.Count; c++)
            {
                var row = start.Row + r;
                var column = start.ColumnNumber + c;
                if (row > MaxRows || column > MaxColumns)
                {
                    return false; // past the sheet edge: DOM path raises the precise error
                }

                plan.Add(sheetName, row, column, ExcelValues.Parse(rowArray[c]));
            }
        }

        return true;
    }

    private const int MaxRows = 1048576;
    private const int MaxColumns = 16384;

    private static string? ValueTypeArg(System.Text.Json.Nodes.JsonObject props) =>
        props.TryGetPropertyValue("valueType", out var node) && node is System.Text.Json.Nodes.JsonValue value &&
        value.GetValueKind() == System.Text.Json.JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    /// <summary>
    /// Parses a plain <c>/Sheet/Cell</c> or <c>/Sheet/Range</c> path. Returns
    /// false for anything else (row/col/element/id-form/named) so the caller
    /// falls back to the DOM. Reuses the streaming-read parser's contract.
    /// </summary>
    private static bool TryParseSheetCellOrRange(string pathText, out string sheetName, out CellRef start, out CellRef end)
    {
        sheetName = string.Empty;
        start = end = default;
        try
        {
            return ExcelStreaming.TryParseCellOrRange(pathText, out sheetName, out start, out end);
        }
        catch (AiofficeException)
        {
            return false; // malformed path: the DOM path will raise the canonical invalid_path
        }
    }

    // ----- apply (the SAX rewrite) --------------------------------------------

    /// <summary>
    /// Applies the plan in place: rewrites each target sheet's part with the
    /// edits spliced in, appends any new shared strings, and sets the recalc
    /// flag when formulas were written. Returns warnings (formula_not_evaluated).
    /// </summary>
    public static List<Warning>? Apply(string file, Plan plan)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart ?? throw Corrupt("the package has no workbook part");

        // Shared strings: one append-only pass. Text values become shared-string
        // indices so the rewrite matches what ClosedXML would have produced.
        var sharedStrings = new SharedStringAppender(workbookPart);
        var dateStyleIndex = new Lazy<uint>(() => DateStyleIndex(workbookPart));

        foreach (var (sheetName, rows) in plan.BySheet)
        {
            var sheetElement = workbookPart.Workbook
                ?.Descendants<S.Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                ?? throw Corrupt($"sheet '{sheetName}' is missing from workbook.xml");
            if (sheetElement.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart part)
            {
                throw Corrupt($"sheet '{sheetName}' points at a missing worksheet part");
            }

            RewriteSheet(part, rows, sharedStrings, dateStyleIndex);
        }

        sharedStrings.Save();

        // The calcChain is a recompute-order cache; once cells change underneath
        // it (a streamed value can be a dependency of a formula elsewhere) it is
        // stale. Drop it and tell Excel to recompute everything on open — the
        // honest, correct choice for a partial in-place edit.
        RemoveCalcChain(workbookPart);
        SetFullRecalc(workbookPart);

        return plan.HasFormulas
            ?
            [
                new Warning(
                    ErrorCodes.FormulaNotEvaluated,
                    "Formula cells were written via the in-place streaming path and carry no cached value yet. " +
                    "Excel recomputes them on open (the workbook's full-recalc flag is set)."),
            ]
            : null;
    }

    /// <summary>Drops the calcChain part (a stale recompute cache after an in-place edit).</summary>
    private static void RemoveCalcChain(WorkbookPart workbookPart)
    {
        if (workbookPart.CalculationChainPart is { } calcChain)
        {
            workbookPart.DeletePart(calcChain);
        }
    }

    private static AiofficeException Corrupt(string what) => new(
        ErrorCodes.InternalError,
        $"In-place streaming write failed: {what}.",
        "The edit was rolled back; re-run it. If this persists, report a bug with the workbook.");

    /// <summary>
    /// Streams one sheet part: copies it verbatim except sheetData, where the
    /// edits are merged into the existing rows in row/column order (overwrite,
    /// insert new cell in column order, insert new row in row order).
    /// </summary>
    private static void RewriteSheet(
        WorksheetPart part,
        SortedDictionary<int, SortedDictionary<int, ExcelValues.ParsedValue>> edits,
        SharedStringAppender sharedStrings,
        Lazy<uint> dateStyleIndex)
    {
        List<OpenXmlAttribute> rootAttributes = [];
        List<KeyValuePair<string, string>> rootNamespaces = [];
        var preElements = new List<OpenXmlElement>();
        var postElements = new List<OpenXmlElement>();
        var sourceRows = new List<S.Row>();

        using (var reader = new OpenXmlPartReader(part))
        {
            while (reader.Read())
            {
                if (reader.ElementType != typeof(S.Worksheet) || !reader.IsStartElement)
                {
                    continue;
                }

                rootAttributes = [.. reader.Attributes];
                rootNamespaces = [.. reader.NamespaceDeclarations];
                if (!reader.ReadFirstChild())
                {
                    break;
                }

                var seenSheetData = false;
                do
                {
                    if (!reader.IsStartElement)
                    {
                        continue;
                    }

                    if (reader.ElementType == typeof(S.SheetData))
                    {
                        seenSheetData = true;
                        ReadRows(reader, sourceRows);
                        continue;
                    }

                    (seenSheetData ? postElements : preElements).Add(reader.LoadCurrentElement()!);
                }
                while (reader.ReadNextSibling());

                break;
            }
        }

        UpdateDimension(preElements, sourceRows, edits);

        using var writer = OpenXmlWriter.Create(part);
        writer.WriteStartElement(new S.Worksheet(), rootAttributes, rootNamespaces);
        foreach (var element in preElements)
        {
            writer.WriteElement(element);
        }

        writer.WriteStartElement(new S.SheetData());
        WriteMergedRows(writer, sourceRows, edits, sharedStrings, dateStyleIndex);
        writer.WriteEndElement(); // sheetData

        foreach (var element in postElements)
        {
            writer.WriteElement(element);
        }

        writer.WriteEndElement(); // worksheet
    }

    /// <summary>Loads the existing rows in document order (each row's cells come with it).</summary>
    private static void ReadRows(OpenXmlPartReader reader, List<S.Row> rows)
    {
        if (!reader.ReadFirstChild())
        {
            return;
        }

        do
        {
            if (reader.ElementType == typeof(S.Row) && reader.IsStartElement)
            {
                rows.Add((S.Row)reader.LoadCurrentElement()!);
            }
        }
        while (reader.ReadNextSibling());
    }

    /// <summary>
    /// Merges the edits into the source rows in ascending row order: existing
    /// rows are emitted (with their cells overwritten/inserted in column order),
    /// brand-new rows are spliced at the right position.
    /// </summary>
    private static void WriteMergedRows(
        OpenXmlWriter writer,
        List<S.Row> sourceRows,
        SortedDictionary<int, SortedDictionary<int, ExcelValues.ParsedValue>> edits,
        SharedStringAppender sharedStrings,
        Lazy<uint> dateStyleIndex)
    {
        var sourceByRow = new SortedDictionary<int, S.Row>();
        var impliedRow = 0;
        foreach (var row in sourceRows)
        {
            var rowNumber = (int?)row.RowIndex?.Value ?? impliedRow + 1;
            impliedRow = rowNumber;
            sourceByRow[rowNumber] = row;
        }

        var allRowNumbers = new SortedSet<int>(sourceByRow.Keys);
        foreach (var rowNumber in edits.Keys)
        {
            allRowNumbers.Add(rowNumber);
        }

        foreach (var rowNumber in allRowNumbers)
        {
            var hasSource = sourceByRow.TryGetValue(rowNumber, out var sourceRow);
            var hasEdits = edits.TryGetValue(rowNumber, out var rowEdits);

            if (hasSource && !hasEdits)
            {
                writer.WriteElement(sourceRow!); // untouched row, byte-for-byte
                continue;
            }

            WriteRow(writer, rowNumber, hasSource ? sourceRow : null, rowEdits, sharedStrings, dateStyleIndex);
        }
    }

    /// <summary>
    /// Writes one row, merging the edited cells into the source row's cells in
    /// ascending column order (edit overwrites source at the same column).
    /// </summary>
    private static void WriteRow(
        OpenXmlWriter writer,
        int rowNumber,
        S.Row? sourceRow,
        SortedDictionary<int, ExcelValues.ParsedValue>? rowEdits,
        SharedStringAppender sharedStrings,
        Lazy<uint> dateStyleIndex)
    {
        // Clone the row's attributes (style/height/customFormat/outline) but not
        // its children; we re-emit the cells ourselves in column order.
        var newRow = new S.Row();
        if (sourceRow is not null)
        {
            foreach (var attribute in sourceRow.GetAttributes())
            {
                newRow.SetAttribute(attribute);
            }
        }

        newRow.RowIndex ??= (uint)rowNumber;

        writer.WriteStartElement(newRow);

        var sourceCells = new SortedDictionary<int, S.Cell>();
        if (sourceRow is not null)
        {
            var impliedColumn = 0;
            foreach (var cell in sourceRow.Elements<S.Cell>())
            {
                var column = ColumnOfRef(cell.CellReference?.Value) ?? impliedColumn + 1;
                impliedColumn = column;
                sourceCells[column] = cell;
            }
        }

        var columns = new SortedSet<int>(sourceCells.Keys);
        if (rowEdits is not null)
        {
            foreach (var column in rowEdits.Keys)
            {
                columns.Add(column);
            }
        }

        foreach (var column in columns)
        {
            if (rowEdits is not null && rowEdits.TryGetValue(column, out var value))
            {
                var styleId = sourceCells.TryGetValue(column, out var existing) ? existing.StyleIndex?.Value : null;
                if (BuildCell(rowNumber, column, value, styleId, sharedStrings, dateStyleIndex) is { } cell)
                {
                    writer.WriteElement(cell);
                }
            }
            else
            {
                writer.WriteElement(sourceCells[column]); // untouched cell, byte-for-byte
            }
        }

        writer.WriteEndElement(); // row
    }

    /// <summary>
    /// Builds one cell. Text → shared string (matching the DOM path). Formula →
    /// formula text, no cached value (recalc-on-open). Date → date style id.
    /// Null for a blank that has no style to preserve (nothing to write).
    /// </summary>
    private static S.Cell? BuildCell(
        int rowNumber,
        int column,
        ExcelValues.ParsedValue value,
        uint? styleId,
        SharedStringAppender sharedStrings,
        Lazy<uint> dateStyleIndex)
    {
        var reference = ExcelCharts.ColumnLetters(column) + rowNumber.ToString(CultureInfo.InvariantCulture);

        if (value.IsFormula)
        {
            var text = value.Formula!;
            var cell = new S.Cell
            {
                CellReference = reference,
                CellFormula = new S.CellFormula(text.StartsWith('=') ? text[1..] : text),
            };
            if (styleId is { } sid)
            {
                cell.StyleIndex = sid;
            }

            return cell;
        }

        switch (value.Value.Type)
        {
            case ClosedXML.Excel.XLDataType.Number:
            {
                var cell = new S.Cell
                {
                    CellReference = reference,
                    CellValue = new S.CellValue(value.Value.GetNumber()),
                };
                if (styleId is { } sid)
                {
                    cell.StyleIndex = sid;
                }

                return cell;
            }

            case ClosedXML.Excel.XLDataType.Boolean:
            {
                var cell = new S.Cell
                {
                    CellReference = reference,
                    DataType = S.CellValues.Boolean,
                    CellValue = new S.CellValue(value.Value.GetBoolean() ? "1" : "0"),
                };
                if (styleId is { } sid)
                {
                    cell.StyleIndex = sid;
                }

                return cell;
            }

            case ClosedXML.Excel.XLDataType.Text:
            {
                var index = sharedStrings.Intern(value.Value.GetText());
                var cell = new S.Cell
                {
                    CellReference = reference,
                    DataType = S.CellValues.SharedString,
                    CellValue = new S.CellValue(index.ToString(CultureInfo.InvariantCulture)),
                };
                if (styleId is { } sid)
                {
                    cell.StyleIndex = sid;
                }

                return cell;
            }

            case ClosedXML.Excel.XLDataType.DateTime:
            {
                var serial = value.Value.GetDateTime().ToOADate();
                return new S.Cell
                {
                    CellReference = reference,
                    StyleIndex = styleId ?? dateStyleIndex.Value,
                    CellValue = new S.CellValue(serial.ToString("R", CultureInfo.InvariantCulture)),
                };
            }

            case ClosedXML.Excel.XLDataType.TimeSpan:
            {
                var serial = value.Value.GetTimeSpan().TotalDays;
                return new S.Cell
                {
                    CellReference = reference,
                    StyleIndex = styleId ?? dateStyleIndex.Value,
                    CellValue = new S.CellValue(serial.ToString("R", CultureInfo.InvariantCulture)),
                };
            }

            default: // blank
                if (styleId is { } blankStyle)
                {
                    return new S.Cell { CellReference = reference, StyleIndex = blankStyle };
                }

                return null;
        }
    }

    /// <summary>Repoints the cloned dimension to cover the source plus the new edits.</summary>
    private static void UpdateDimension(
        List<OpenXmlElement> preElements,
        List<S.Row> sourceRows,
        SortedDictionary<int, SortedDictionary<int, ExcelValues.ParsedValue>> edits)
    {
        if (preElements.OfType<S.SheetDimension>().FirstOrDefault() is not { } dimension)
        {
            return;
        }

        int firstRow = int.MaxValue, lastRow = 0, firstColumn = int.MaxValue, lastColumn = 0;

        void Note(int row, int column)
        {
            firstRow = Math.Min(firstRow, row);
            lastRow = Math.Max(lastRow, row);
            firstColumn = Math.Min(firstColumn, column);
            lastColumn = Math.Max(lastColumn, column);
        }

        // Existing extent (parse the source dimension if present, else scan).
        if (dimension.Reference?.Value is { } existing && TryParseRange(existing, out var es, out var ee))
        {
            Note(es.Row, es.ColumnNumber);
            Note(ee.Row, ee.ColumnNumber);
        }
        else
        {
            var impliedRow = 0;
            foreach (var row in sourceRows)
            {
                var rowNumber = (int?)row.RowIndex?.Value ?? impliedRow + 1;
                impliedRow = rowNumber;
                var impliedColumn = 0;
                foreach (var cell in row.Elements<S.Cell>())
                {
                    var column = ColumnOfRef(cell.CellReference?.Value) ?? impliedColumn + 1;
                    impliedColumn = column;
                    Note(rowNumber, column);
                }
            }
        }

        foreach (var (rowNumber, cells) in edits)
        {
            foreach (var column in cells.Keys)
            {
                Note(rowNumber, column);
            }
        }

        if (lastRow == 0)
        {
            return;
        }

        var start = ExcelCharts.ColumnLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture);
        var end = ExcelCharts.ColumnLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture);
        dimension.Reference = start == end ? start : start + ":" + end;
    }

    // ----- shared strings -----------------------------------------------------

    /// <summary>
    /// Append-only shared-string interner over the existing sst part (loaded
    /// once). Existing entries are indexed for reuse; new text is appended. The
    /// part is rewritten on <see cref="Save"/> only when something was added.
    /// </summary>
    private sealed class SharedStringAppender
    {
        private readonly WorkbookPart _workbookPart;
        private readonly List<string> _items = [];
        private readonly Dictionary<string, int> _index = new(StringComparer.Ordinal);
        private bool _dirty;

        public SharedStringAppender(WorkbookPart workbookPart)
        {
            _workbookPart = workbookPart;
            if (workbookPart.SharedStringTablePart?.SharedStringTable is { } table)
            {
                foreach (var item in table.Elements<S.SharedStringItem>())
                {
                    var text = SharedStringText(item);
                    if (!_index.ContainsKey(text))
                    {
                        _index[text] = _items.Count;
                    }

                    _items.Add(text);
                }
            }
        }

        public int Intern(string text)
        {
            if (_index.TryGetValue(text, out var existing))
            {
                return existing;
            }

            var index = _items.Count;
            _items.Add(text);
            _index[text] = index;
            _dirty = true;
            return index;
        }

        public void Save()
        {
            if (!_dirty)
            {
                return;
            }

            var part = _workbookPart.SharedStringTablePart ??
                       _workbookPart.AddNewPart<SharedStringTablePart>();
            using var writer = OpenXmlWriter.Create(part);
            writer.WriteStartElement(new S.SharedStringTable
            {
                Count = (uint)_items.Count,
                UniqueCount = (uint)_items.Count,
            });
            foreach (var text in _items)
            {
                writer.WriteStartElement(new S.SharedStringItem());
                var t = new S.Text(text);
                if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1])))
                {
                    t.Space = SpaceProcessingModeValues.Preserve;
                }

                writer.WriteElement(t);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        private static string SharedStringText(S.SharedStringItem item) =>
            string.Concat(item
                .Descendants<S.Text>()
                .Where(t => t.Ancestors<S.PhoneticRun>().FirstOrDefault() is null)
                .Select(t => t.Text));
    }

    // ----- styles (date) ------------------------------------------------------

    /// <summary>
    /// Finds (or appends) a cellXf with the built-in short-date number format
    /// (id 14), returning its 0-based index. Mirrors the date style ClosedXML
    /// would apply, so a streamed date renders identically. When the workbook
    /// has no styles part (a minimal generated workbook), a complete minimal
    /// stylesheet is created first so every style index resolves.
    /// </summary>
    private static uint DateStyleIndex(WorkbookPart workbookPart)
    {
        var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet ??= EmptyStylesheet();
        var stylesheet = stylesPart.Stylesheet;

        // A valid stylesheet needs at least one font/fill/border and a base
        // cellXf; ClosedXML throws on load otherwise. Backfill any missing
        // collection (a hand-rolled workbook may omit them entirely).
        EnsureMinimalCollections(stylesheet);
        var cellFormats = stylesheet.CellFormats!;

        const uint shortDateFormatId = 14; // built-in "m/d/yyyy"
        var formats = cellFormats.Elements<S.CellFormat>().ToList();
        for (var i = 0; i < formats.Count; i++)
        {
            if (formats[i].NumberFormatId?.Value == shortDateFormatId && formats[i].ApplyNumberFormat?.Value == true)
            {
                return (uint)i;
            }
        }

        cellFormats.AppendChild(new S.CellFormat
        {
            NumberFormatId = shortDateFormatId,
            FontId = 0,
            FillId = 0,
            BorderId = 0,
            FormatId = 0,
            ApplyNumberFormat = true,
        });
        cellFormats.Count = (uint)(formats.Count + 1);
        stylesheet.Save();
        return (uint)formats.Count;
    }

    private static S.Stylesheet EmptyStylesheet() => new();

    /// <summary>Backfills the mandatory fonts/fills/borders/cellStyleXfs/cellXfs so style 0 resolves.</summary>
    private static void EnsureMinimalCollections(S.Stylesheet stylesheet)
    {
        if (stylesheet.Fonts is null || !stylesheet.Fonts.Elements<S.Font>().Any())
        {
            stylesheet.Fonts = new S.Fonts(new S.Font(
                new S.FontSize { Val = 11D },
                new S.FontName { Val = "Calibri" }))
            { Count = 1 };
        }

        if (stylesheet.Fills is null || stylesheet.Fills.Elements<S.Fill>().Count() < 2)
        {
            // Excel requires fill 0 = none and fill 1 = gray125 (the reserved pair).
            stylesheet.Fills = new S.Fills(
                new S.Fill(new S.PatternFill { PatternType = S.PatternValues.None }),
                new S.Fill(new S.PatternFill { PatternType = S.PatternValues.Gray125 }))
            { Count = 2 };
        }

        if (stylesheet.Borders is null || !stylesheet.Borders.Elements<S.Border>().Any())
        {
            stylesheet.Borders = new S.Borders(new S.Border(
                new S.LeftBorder(), new S.RightBorder(), new S.TopBorder(),
                new S.BottomBorder(), new S.DiagonalBorder()))
            { Count = 1 };
        }

        if (stylesheet.CellStyleFormats is null || !stylesheet.CellStyleFormats.Elements<S.CellFormat>().Any())
        {
            stylesheet.CellStyleFormats = new S.CellStyleFormats(new S.CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
            })
            { Count = 1 };
        }

        if (stylesheet.CellFormats is null || !stylesheet.CellFormats.Elements<S.CellFormat>().Any())
        {
            stylesheet.CellFormats = new S.CellFormats(new S.CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
            })
            { Count = 1 };
        }
    }

    // ----- recalc -------------------------------------------------------------

    /// <summary>Sets the workbook's full-recalc-on-load flag so Excel recomputes streamed formulas.</summary>
    private static void SetFullRecalc(WorkbookPart workbookPart)
    {
        if (workbookPart.Workbook is not { } workbook)
        {
            return;
        }

        workbook.CalculationProperties ??= new S.CalculationProperties { CalculationId = 0 };
        workbook.CalculationProperties.FullCalculationOnLoad = true;
        workbook.CalculationProperties.ForceFullCalculation = true;
        workbook.Save();
    }

    // ----- small helpers ------------------------------------------------------

    private static int? ColumnOfRef(string? reference)
    {
        if (string.IsNullOrEmpty(reference))
        {
            return null;
        }

        var n = 0;
        foreach (var c in reference)
        {
            if (c is >= 'A' and <= 'Z')
            {
                n = (n * 26) + (c - 'A' + 1);
            }
            else if (c is >= 'a' and <= 'z')
            {
                n = (n * 26) + (c - 'a' + 1);
            }
            else
            {
                break;
            }
        }

        return n == 0 ? null : n;
    }

    private static bool TryParseRange(string text, out CellRef start, out CellRef end)
    {
        start = end = default;
        var colon = text.IndexOf(':');
        if (colon < 0)
        {
            return TryParseCell(text, out start) && (end = start) == start;
        }

        return TryParseCell(text[..colon], out start) && TryParseCell(text[(colon + 1)..], out end);
    }

    private static bool TryParseCell(string text, out CellRef cell)
    {
        cell = default;
        var i = 0;
        while (i < text.Length && char.IsAsciiLetter(text[i]))
        {
            i++;
        }

        if (i == 0 || i == text.Length)
        {
            return false;
        }

        if (!int.TryParse(text.AsSpan(i), NumberStyles.None, CultureInfo.InvariantCulture, out var row))
        {
            return false;
        }

        cell = new CellRef(text[..i].ToUpperInvariant(), row);
        return true;
    }
}
