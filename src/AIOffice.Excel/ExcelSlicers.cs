using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>The slicer source kinds aioffice can wire (a table column or a pivot field).</summary>
internal enum SlicerSource
{
    Table,
    Pivot,
}

/// <summary>
/// A fully validated slicer-add, captured from the in-memory workbook while the
/// edit batch is applied. The raw OpenXml write happens later in the post-save
/// pass (<see cref="ExcelSlicers.Apply"/>), because ClosedXML cannot author
/// slicer parts itself (measured: ClosedXML 0.105 preserves them byte-identical
/// on later saves, exactly like charts).
/// </summary>
internal sealed record SlicerAddSpec(
    string SheetName,
    SlicerSource Source,
    string SlicerName,
    string Column,
    string Caption,
    string Style,
    int AnchorColumn,
    int AnchorRow,
    int WidthCells,
    int HeightCells,
    // Table source: the ListObject name (id is package-assigned, resolved at
    // apply time) + the 0-based column index inside it.
    string TableName,
    uint TableColumnIndex,
    // Pivot source: the host pivot's name (tabId/cacheId/item count are
    // package-assigned, resolved at apply time).
    string PivotName);

/// <summary>A validated slicer removal (1-based per-sheet index at apply time).</summary>
internal sealed record SlicerRemoveSpec(string SheetName, int Index);

/// <summary>A slicer as read back from the raw package (for get / read --view structure).</summary>
internal sealed record SlicerInfo(
    string SheetName,
    int Index,
    string Path,
    string Name,
    string SourceKind,
    string Source,
    string Column,
    string Caption);

/// <summary>
/// Excel slicers (M8 dashboard feature): a table-column slicer or a pivot-field
/// slicer, authored entirely on raw OpenXml because ClosedXML cannot create
/// them. The part layout is the one Excel itself writes and the OpenXmlValidator
/// (Office2019) accepts with zero errors:
///
/// <list type="bullet">
/// <item>A <c>SlicerCachePart</c> on the workbook holds the
/// <c>x14:slicerCacheDefinition</c>. A TABLE slicer carries the connection in an
/// <c>x15:tableSlicerCache</c> ext (no <c>&lt;data&gt;</c>); a PIVOT slicer
/// carries <c>x14:slicerCachePivotTables</c> + an <c>x14:data/x14:tabular</c>
/// with one item per distinct field value.</item>
/// <item>The workbook <c>extLst</c> references the cache: an <c>x15:slicerCaches</c>
/// ext for table slicers, an <c>x14:slicerCaches</c> ext for pivot slicers.</item>
/// <item>A <c>SlicersPart</c> on the host worksheet holds the <c>x14:slicers</c> /
/// <c>x14:slicer</c>; the worksheet <c>extLst</c> points at it via
/// <c>x14:slicerList</c>; and a worksheet-drawing <c>mc:AlternateContent</c>
/// anchor carries the <c>sle:slicer</c> graphic frame (with a shape fallback).</item>
/// </list>
///
/// Unsupported source kinds (e.g. a slicer on a bare range, or an OLAP/timeline
/// slicer) return <c>unsupported_feature</c> naming the workaround.
/// </summary>
internal static partial class ExcelSlicers
{
    public const string X14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    public const string X15Ns = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
    private const string XmNs = "http://schemas.microsoft.com/office/excel/2006/main";
    private const string MainNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string DrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private const string McNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string SlicerDrawingNs = "http://schemas.microsoft.com/office/drawing/2010/slicer";

    // The extension Uris Excel stamps on each ext (stable GUIDs from the OOXML spec).
    private const string TableSlicerCacheExtUri = "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}";
    private const string WorkbookSlicerCachesX15Uri = "{46BE6895-7355-4a93-B00E-2C351335B9C9}";
    private const string WorkbookSlicerCachesX14Uri = "{BBE1A952-AA13-448e-AADC-164F8A28A991}";
    private const string WorksheetSlicerListUri = "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}";

    private const string DefaultStyle = "SlicerStyleLight1";
    private const int DefaultWidthCells = 2;
    private const int DefaultHeightCells = 7;

    // ----- op-time validation & data capture ---------------------------------

    /// <summary>
    /// Validates an <c>add slicer</c> op against the in-memory workbook and
    /// captures everything the post-save writer needs. The op path names the
    /// table (<c>/Sheet1/table[@name=Sales]</c>) or pivot
    /// (<c>/Sheet1/pivot[1]</c>); props.column (table) or props.field (pivot)
    /// names the source column/field.
    /// </summary>
    public static SlicerAddSpec ParseAdd(XLWorkbook workbook, ExcelTarget target, EditOp op, int opIndex)
    {
        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add slicer needs props.",
                "On a table: {op:add, type:slicer, path:/Sheet1/table[@name=Sales], props:{column:\"Region\"}}. " +
                "On a pivot: {op:add, type:slicer, path:/Sheet1/pivot[1], props:{field:\"Region\"}}.");
        }

        return target.Kind switch
        {
            ExcelTargetKind.Table => ParseTableSlicer(target, op, opIndex),
            ExcelTargetKind.Pivot => ParsePivotSlicer(workbook, target, op, opIndex),
            _ => throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{opIndex}]: a slicer source must be a table or a pivot, not '{target.Kind}'.",
                "Address a table (/Sheet1/table[@name=Sales]) or a pivot (/Sheet1/pivot[1]); " +
                "slicers on bare ranges, OLAP cubes and timelines are not supported."),
        };
    }

    private static SlicerAddSpec ParseTableSlicer(ExcelTarget target, EditOp op, int opIndex)
    {
        var table = ExcelTables.Find(target); // invalid_path with candidates when absent
        var column = RequiredString(op.Props!, "column", opIndex,
            "the table column to slice, e.g. {\"column\":\"Region\"}");

        var fields = table.Fields.ToList();
        var fieldIndex = fields.FindIndex(f =>
            string.Equals(f.Name, column, StringComparison.OrdinalIgnoreCase));
        if (fieldIndex < 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: table '{table.Name}' has no column '{column}'.",
                "Slicers are keyed by column name; the table's columns are listed as candidates.",
                candidates: [.. fields.Select(f => f.Name)]);
        }

        var resolvedColumn = fields[fieldIndex].Name;
        var (anchorColumn, anchorRow, width, height) = ReadGeometry(target.Sheet, op.Props!, opIndex);
        return new SlicerAddSpec(
            SheetName: target.Sheet.Name,
            Source: SlicerSource.Table,
            SlicerName: SlicerNameFor(resolvedColumn),
            Column: resolvedColumn,
            Caption: OptionalString(op.Props!, "caption") ?? resolvedColumn,
            Style: ResolveStyle(op.Props!),
            AnchorColumn: anchorColumn,
            AnchorRow: anchorRow,
            WidthCells: width,
            HeightCells: height,
            TableName: table.Name,
            TableColumnIndex: (uint)fieldIndex,
            PivotName: string.Empty);
    }

    private static SlicerAddSpec ParsePivotSlicer(XLWorkbook workbook, ExcelTarget target, EditOp op, int opIndex)
    {
        var (pivot, _) = ExcelPivots.Find(target); // invalid_path with candidates when absent
        var field = RequiredString(op.Props!, "field", opIndex,
            "the pivot field to slice, e.g. {\"field\":\"Region\"}");

        var fieldNames = pivot.PivotCache?.FieldNames?.ToList() ?? [];
        var sourceField = fieldNames.FirstOrDefault(name =>
            string.Equals(name, field, StringComparison.OrdinalIgnoreCase));
        if (sourceField is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: pivot '{pivot.Name}' has no field '{field}'.",
                "Slicers are keyed by the source field name; the fields are listed as candidates.",
                candidates: fieldNames.Count > 0 ? fieldNames : [field]);
        }

        var (anchorColumn, anchorRow, width, height) = ReadGeometry(target.Sheet, op.Props!, opIndex);
        return new SlicerAddSpec(
            SheetName: target.Sheet.Name,
            Source: SlicerSource.Pivot,
            SlicerName: SlicerNameFor(sourceField),
            Column: sourceField,
            Caption: OptionalString(op.Props!, "caption") ?? sourceField,
            Style: ResolveStyle(op.Props!),
            AnchorColumn: anchorColumn,
            AnchorRow: anchorRow,
            WidthCells: width,
            HeightCells: height,
            TableName: string.Empty,
            TableColumnIndex: 0,
            PivotName: pivot.Name);
    }

    private static (int Column, int Row, int Width, int Height) ReadGeometry(
        IXLWorksheet sheet, JsonObject props, int opIndex)
    {
        // props.x is the top-left anchor cell (e.g. "E2"); default is two columns
        // right of the sheet's used range so the slicer never lands on data.
        var anchorText = OptionalString(props, "x");
        int anchorColumn;
        int anchorRow;
        if (anchorText is not null)
        {
            var address = ParseCell(anchorText, opIndex);
            anchorColumn = address.Column;
            anchorRow = address.Row;
        }
        else
        {
            var used = sheet.RangeUsed();
            anchorColumn = (used?.RangeAddress.LastAddress.ColumnNumber ?? 1) + 2;
            anchorRow = 1;
        }

        var width = OptionalInt(props, "widthCells", opIndex) ?? DefaultWidthCells;
        var height = OptionalInt(props, "heightCells", opIndex) ?? DefaultHeightCells;
        if (width < 1 || height < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: slicer widthCells/heightCells must be at least 1.",
                "Drop the props to use the default 2x7 box, or pass positive cell counts.");
        }

        return (anchorColumn, anchorRow, width, height);
    }

    /// <summary>The proposed slicer name (spaces/quotes become underscores, as Excel does); made unique at apply time.</summary>
    private static string SlicerNameFor(string column) =>
        "Slicer_" + column.Replace(' ', '_').Replace('\'', '_');

    private static string ResolveStyle(JsonObject props)
    {
        var style = OptionalString(props, "style");
        if (string.IsNullOrEmpty(style))
        {
            return DefaultStyle;
        }

        // Accept either the full "SlicerStyleLight1" or the short "light1" form.
        if (style.StartsWith("SlicerStyle", StringComparison.OrdinalIgnoreCase))
        {
            return style;
        }

        foreach (var tier in new[] { "Light", "Dark", "Other" })
        {
            if (style.StartsWith(tier, StringComparison.OrdinalIgnoreCase) &&
                style.Length > tier.Length &&
                int.TryParse(style.AsSpan(tier.Length), NumberStyles.None, CultureInfo.InvariantCulture, out var n))
            {
                return string.Create(CultureInfo.InvariantCulture, $"SlicerStyle{tier}{n}");
            }
        }

        return style; // a custom (workbook) style name passes through untouched
    }

    // ----- batch (op accumulation, mirrors ChartOpBatch) ----------------------

    /// <summary>
    /// Collects slicer ops during an edit batch so they validate up front
    /// (a bad op aborts before any byte is written) and apply after ClosedXML
    /// has saved. Projected per-sheet slicer counts keep remove indices in a
    /// multi-op batch consistent.
    /// </summary>
    public sealed class Batch
    {
        private readonly string _file;
        private Dictionary<string, int>? _projectedCounts;

        public Batch(string file) => _file = file;

        /// <summary>Slicer ops in batch order (<see cref="SlicerAddSpec"/> / <see cref="SlicerRemoveSpec"/>).</summary>
        public List<object> Ops { get; } = [];

        public bool IsEmpty => Ops.Count == 0;

        /// <summary>Queues an add and returns the slicer's projected 1-based index on its sheet.</summary>
        public int Add(SlicerAddSpec spec)
        {
            var counts = ProjectedCounts();
            var next = counts.GetValueOrDefault(spec.SheetName) + 1;
            counts[spec.SheetName] = next;
            Ops.Add(spec);
            return next;
        }

        /// <summary>Queues a removal, validating the 1-based index against the projected state.</summary>
        public void Remove(string sheetName, string sheetPath, int index, int opIndex)
        {
            var counts = ProjectedCounts();
            var current = counts.GetValueOrDefault(sheetName);
            if (index > current || index < 1)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"ops[{opIndex}]: no slicer[{index}] on sheet '{sheetName}' ({current} slicer(s) exist).",
                    current > 0
                        ? "Slicer indices are 1-based per sheet; run 'aioffice read --view structure' to list them."
                        : "This sheet has no slicers; add one with {op:add, type:slicer, path:" + sheetPath +
                          "/table[@name=Sales], props:{column:\"Region\"}}.",
                    candidates: current > 0
                        ? [.. Enumerable.Range(1, current).Select(i => $"{sheetPath}/slicer[{i}]")]
                        : [sheetPath]);
            }

            counts[sheetName] = current - 1;
            Ops.Add(new SlicerRemoveSpec(sheetName, index));
        }

        private Dictionary<string, int> ProjectedCounts()
        {
            if (_projectedCounts is null)
            {
                _projectedCounts = new Dictionary<string, int>(StringComparer.Ordinal);
                using var document = SpreadsheetDocument.Open(_file, isEditable: false);
                foreach (var info in Read(document))
                {
                    _projectedCounts[info.SheetName] = info.Index; // indices are sequential per sheet
                }
            }

            return _projectedCounts;
        }
    }

    // ----- apply (raw OpenXml write) ------------------------------------------

    public static void Apply(string file, IReadOnlyList<object> ops)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        foreach (var op in ops)
        {
            switch (op)
            {
                case SlicerAddSpec add:
                    AddSlicer(document, add);
                    break;
                case SlicerRemoveSpec remove:
                    RemoveSlicer(document, remove);
                    break;
            }
        }
    }

    private static void AddSlicer(SpreadsheetDocument document, SlicerAddSpec rawSpec)
    {
        var workbookPart = document.WorkbookPart
            ?? throw Disappeared(rawSpec.SheetName);
        var worksheetPart = SheetPartOrThrow(document, rawSpec.SheetName);

        // Make the name unique against every slicer cache already in the package
        // (a multi-slicer batch on the same column would otherwise collide).
        var spec = rawSpec with { SlicerName = UniqueName(workbookPart, rawSpec.SlicerName) };

        var cachePart = workbookPart.AddNewPart<SlicerCachePart>();
        cachePart.SlicerCacheDefinition = spec.Source == SlicerSource.Table
            ? BuildTableCacheDefinition(spec, worksheetPart)
            : BuildPivotCacheDefinition(spec, workbookPart);
        var cacheRelId = workbookPart.GetIdOfPart(cachePart);

        AppendWorkbookSlicerCache(workbookPart, spec, cacheRelId);

        var slicersPart = worksheetPart.SlicersParts.FirstOrDefault() ?? worksheetPart.AddNewPart<SlicersPart>();
        if (slicersPart.Slicers is null)
        {
            var slicers = new X14.Slicers();
            slicers.AddNamespaceDeclaration("x14", X14Ns);
            slicers.AddNamespaceDeclaration("xm", XmNs);
            slicersPart.Slicers = slicers;
        }

        slicersPart.Slicers.Append(new X14.Slicer
        {
            Name = spec.SlicerName,
            Cache = spec.SlicerName,
            Caption = spec.Caption,
            RowHeight = 241300U,
            Style = spec.Style,
            ColumnCount = 1U,
        });

        AppendWorksheetSlicerList(worksheetPart, slicersPart);
        AppendDrawingAnchor(worksheetPart, spec);

        workbookPart.Workbook!.Save();
        worksheetPart.Worksheet!.Save();
        cachePart.SlicerCacheDefinition!.Save();
        slicersPart.Slicers!.Save();
        worksheetPart.DrawingsPart!.WorksheetDrawing!.Save();
    }

    private static X14.SlicerCacheDefinition BuildTableCacheDefinition(SlicerAddSpec spec, WorksheetPart worksheetPart)
    {
        var tableId = TableIdOrThrow(worksheetPart, spec);
        var definition = new X14.SlicerCacheDefinition
        {
            Name = spec.SlicerName,
            SourceName = spec.Column,
        };
        definition.AddNamespaceDeclaration("x14", X14Ns);
        definition.AddNamespaceDeclaration("xm", XmNs);

        // No <data> for a table slicer: the connection lives in an x15 ext.
        var extensionList = new X14.SlicerCacheDefinitionExtensionList();
        var extension = new S.SlicerCacheDefinitionExtension { Uri = TableSlicerCacheExtUri };
        extension.AddNamespaceDeclaration("x15", X15Ns);
        extension.Append(new X15.TableSlicerCache
        {
            TableId = tableId,
            Column = spec.TableColumnIndex + 1, // x15 column is 1-based within the table
        });
        extensionList.Append(extension);
        definition.Append(extensionList);
        return definition;
    }

    private static X14.SlicerCacheDefinition BuildPivotCacheDefinition(SlicerAddSpec spec, WorkbookPart workbookPart)
    {
        var (tabId, cacheId, itemCount) = PivotIdsOrThrow(workbookPart, spec);
        var definition = new X14.SlicerCacheDefinition
        {
            Name = spec.SlicerName,
            SourceName = spec.Column,
        };
        definition.AddNamespaceDeclaration("x14", X14Ns);
        definition.AddNamespaceDeclaration("xm", XmNs);

        var pivotTables = new X14.SlicerCachePivotTables();
        pivotTables.Append(new X14.SlicerCachePivotTable { TabId = tabId, Name = spec.PivotName });
        definition.Append(pivotTables);

        // The tabular cache needs one selectable item per distinct field value;
        // the schema rejects an empty <items>, so floor at one.
        var safeCount = Math.Max(1, itemCount);
        var data = new X14.SlicerCacheData();
        var tabular = new X14.TabularSlicerCache { PivotCacheId = cacheId };
        var items = new X14.TabularSlicerCacheItems { Count = (uint)safeCount };
        for (var i = 0; i < safeCount; i++)
        {
            items.Append(new X14.TabularSlicerCacheItem { Atom = (uint)i, IsSelected = true });
        }

        tabular.Append(items);
        data.Append(tabular);
        definition.Append(data);
        return definition;
    }

    private static void AppendWorkbookSlicerCache(WorkbookPart workbookPart, SlicerAddSpec spec, string cacheRelId)
    {
        var workbook = workbookPart.Workbook!;
        var extensionList = workbook.GetFirstChild<S.WorkbookExtensionList>();
        if (extensionList is null)
        {
            extensionList = new S.WorkbookExtensionList();
            workbook.Append(extensionList);
        }

        if (spec.Source == SlicerSource.Table)
        {
            var caches = FindOrCreateX15Caches(extensionList);
            caches.Append(new X14.SlicerCache { Id = cacheRelId });
        }
        else
        {
            var caches = FindOrCreateX14Caches(extensionList);
            caches.Append(new X14.SlicerCache { Id = cacheRelId });
        }
    }

    private static X15.SlicerCaches FindOrCreateX15Caches(S.WorkbookExtensionList extensionList)
    {
        var existing = extensionList.Elements<S.WorkbookExtension>()
            .SelectMany(e => e.Elements<X15.SlicerCaches>())
            .FirstOrDefault();
        if (existing is not null)
        {
            return existing;
        }

        var extension = new S.WorkbookExtension { Uri = WorkbookSlicerCachesX15Uri };
        extension.AddNamespaceDeclaration("x15", X15Ns);
        var caches = new X15.SlicerCaches();
        extension.Append(caches);
        extensionList.Append(extension);
        return caches;
    }

    private static X14.SlicerCaches FindOrCreateX14Caches(S.WorkbookExtensionList extensionList)
    {
        var existing = extensionList.Elements<S.WorkbookExtension>()
            .SelectMany(e => e.Elements<X14.SlicerCaches>())
            .FirstOrDefault();
        if (existing is not null)
        {
            return existing;
        }

        var extension = new S.WorkbookExtension { Uri = WorkbookSlicerCachesX14Uri };
        extension.AddNamespaceDeclaration("x14", X14Ns);
        var caches = new X14.SlicerCaches();
        extension.Append(caches);
        extensionList.Append(extension);
        return caches;
    }

    private static void AppendWorksheetSlicerList(WorksheetPart worksheetPart, SlicersPart slicersPart)
    {
        var worksheet = worksheetPart.Worksheet!;
        var slicersRelId = worksheetPart.GetIdOfPart(slicersPart);

        var extensionList = worksheet.GetFirstChild<S.WorksheetExtensionList>();
        if (extensionList is null)
        {
            extensionList = new S.WorksheetExtensionList();
            worksheet.Append(extensionList);
        }

        var slicerList = extensionList.Elements<S.WorksheetExtension>()
            .SelectMany(e => e.Elements<X14.SlicerList>())
            .FirstOrDefault();
        if (slicerList is null)
        {
            var extension = new S.WorksheetExtension { Uri = WorksheetSlicerListUri };
            extension.AddNamespaceDeclaration("x14", X14Ns);
            slicerList = new X14.SlicerList();
            extension.Append(slicerList);
            extensionList.Append(extension);
        }

        slicerList.Append(new X14.SlicerRef { Id = slicersRelId });
    }

    private static void AppendDrawingAnchor(WorksheetPart worksheetPart, SlicerAddSpec spec)
    {
        var drawings = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
        if (drawings.WorksheetDrawing is null)
        {
            var root = new Xdr.WorksheetDrawing();
            root.AddNamespaceDeclaration("xdr", DrawingNs);
            root.AddNamespaceDeclaration("a", MainNs);
            drawings.WorksheetDrawing = root;
        }

        EnsureDrawingElement(worksheetPart, drawings);

        var root2 = drawings.WorksheetDrawing!;
        var frameId = root2.Descendants<Xdr.NonVisualDrawingProperties>()
            .Select(p => p.Id?.Value ?? 0u)
            .Concat(SlicerFrameIds(root2))
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
        anchor.Append(BuildSlicerFrame(spec.SlicerName, frameId));
        anchor.Append(new Xdr.ClientData());
        root2.Append(anchor);
    }

    /// <summary>The cNvPr ids already used by sle:slicer alternate-content frames.</summary>
    private static IEnumerable<uint> SlicerFrameIds(Xdr.WorksheetDrawing root) =>
        root.Descendants<AlternateContent>()
            .SelectMany(ac => ac.Descendants())
            .Where(e => e.LocalName == "cNvPr")
            .Select(e => e.GetAttributes().FirstOrDefault(a => a.LocalName == "id").Value)
            .Where(v => v is not null)
            .Select(v => uint.TryParse(v, NumberStyles.None, CultureInfo.InvariantCulture, out var n) ? n : 0u);

    /// <summary>
    /// The <c>mc:AlternateContent</c> graphic frame Excel writes for a slicer: an
    /// <c>sle:slicer</c> choice (the real slicer) with a shape fallback for
    /// readers that do not understand slicers.
    /// </summary>
    private static OpenXmlElement BuildSlicerFrame(string name, uint frameId)
    {
        var idText = Invariant((int)frameId);
        var inner =
            $"<mc:AlternateContent xmlns:mc=\"{McNs}\">" +
            $"<mc:Choice xmlns:sle=\"{SlicerDrawingNs}\" Requires=\"sle\">" +
            $"<xdr:graphicFrame xmlns:xdr=\"{DrawingNs}\" macro=\"\">" +
            "<xdr:nvGraphicFramePr>" +
            $"<xdr:cNvPr id=\"{idText}\" name=\"{XmlEscape(name)}\"/>" +
            "<xdr:cNvGraphicFramePr/>" +
            "</xdr:nvGraphicFramePr>" +
            $"<xdr:xfrm><a:off x=\"0\" y=\"0\" xmlns:a=\"{MainNs}\"/><a:ext cx=\"0\" cy=\"0\" xmlns:a=\"{MainNs}\"/></xdr:xfrm>" +
            $"<a:graphic xmlns:a=\"{MainNs}\">" +
            $"<a:graphicData uri=\"{SlicerDrawingNs}\">" +
            $"<sle:slicer name=\"{XmlEscape(name)}\"/>" +
            "</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>" +
            "<mc:Fallback>" +
            $"<xdr:sp xmlns:xdr=\"{DrawingNs}\" macro=\"\" textlink=\"\"><xdr:nvSpPr>" +
            $"<xdr:cNvPr id=\"{idText}\" name=\"{XmlEscape(name)}\"/>" +
            $"<xdr:cNvSpPr><a:spLocks xmlns:a=\"{MainNs}\" noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr>" +
            $"<xdr:spPr><a:xfrm xmlns:a=\"{MainNs}\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>" +
            $"<a:prstGeom xmlns:a=\"{MainNs}\" prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>" +
            "</mc:Fallback></mc:AlternateContent>";

        var holder = new Xdr.TwoCellAnchor { InnerXml = inner };
        var element = holder.FirstChild!;
        element.Remove();
        return element;
    }

    /// <summary>
    /// Makes sure the worksheet references its drawings part, inserting
    /// <c>&lt;drawing&gt;</c> in its schema slot (before tableParts/extLst).
    /// </summary>
    private static void EnsureDrawingElement(WorksheetPart worksheetPart, DrawingsPart drawings)
    {
        var worksheet = worksheetPart.Worksheet!;
        if (worksheet.Elements<S.Drawing>().Any())
        {
            return;
        }

        var drawing = new S.Drawing { Id = worksheetPart.GetIdOfPart(drawings) };
        var successor = worksheet.Elements<OpenXmlElement>().FirstOrDefault(e =>
            e is S.LegacyDrawing or S.LegacyDrawingHeaderFooter or S.DrawingHeaderFooter or S.Picture
                or S.OleObjects or S.Controls or S.WebPublishItems or S.TableParts or S.WorksheetExtensionList
                or S.ExtensionList);
        if (successor is null)
        {
            worksheet.Append(drawing);
        }
        else
        {
            worksheet.InsertBefore(drawing, successor);
        }
    }

    // ----- remove -------------------------------------------------------------

    private static void RemoveSlicer(SpreadsheetDocument document, SlicerRemoveSpec spec)
    {
        var workbookPart = document.WorkbookPart ?? throw Disappeared(spec.SheetName);
        var worksheetPart = SheetPartOrThrow(document, spec.SheetName);
        var slicersPart = worksheetPart.SlicersParts.FirstOrDefault();
        var slicerElements = slicersPart?.Slicers?.Elements<X14.Slicer>().ToList() ?? [];
        if (slicersPart is null || spec.Index > slicerElements.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"slicer[{spec.Index}] on '{spec.SheetName}' disappeared between validation and the slicer write pass.",
                "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
        }

        var slicer = slicerElements[spec.Index - 1];
        var slicerName = slicer.Name?.Value;
        slicer.Remove();

        // Drop the workbook slicer cache + its part for this slicer's cache.
        if (slicerName is not null)
        {
            RemoveWorkbookSlicerCache(workbookPart, slicerName);
        }

        // Drop the drawing frame whose sle:slicer matches this slicer.
        if (slicerName is not null && worksheetPart.DrawingsPart?.WorksheetDrawing is { } root)
        {
            RemoveSlicerFrame(root, slicerName);
        }

        // If the worksheet has no slicers left, tear down the slicers part +
        // the worksheet slicerList ext entry (keep the file tidy + valid).
        if (slicersPart.Slicers is { } slicers && !slicers.Elements<X14.Slicer>().Any())
        {
            var relId = worksheetPart.GetIdOfPart(slicersPart);
            RemoveWorksheetSlicerListEntry(worksheetPart, relId);
            worksheetPart.DeletePart(slicersPart);
        }

        workbookPart.Workbook!.Save();
        worksheetPart.Worksheet!.Save();
        worksheetPart.DrawingsPart?.WorksheetDrawing?.Save();
    }

    private static void RemoveWorkbookSlicerCache(WorkbookPart workbookPart, string slicerName)
    {
        // The slicer's cache part is the SlicerCacheDefinition whose Name matches.
        var cachePart = workbookPart.SlicerCacheParts.FirstOrDefault(p =>
            string.Equals(p.SlicerCacheDefinition?.Name?.Value, slicerName, StringComparison.Ordinal));
        if (cachePart is null)
        {
            return;
        }

        var relId = workbookPart.GetIdOfPart(cachePart);
        foreach (var cacheRef in workbookPart.Workbook!.Descendants<X14.SlicerCache>()
                     .Where(c => string.Equals(c.Id?.Value, relId, StringComparison.Ordinal))
                     .ToList())
        {
            var caches = cacheRef.Parent;
            cacheRef.Remove();
            PruneEmptyCachesExt(caches);
        }

        workbookPart.DeletePart(cachePart);
    }

    private static void PruneEmptyCachesExt(OpenXmlElement? caches)
    {
        if (caches is null || caches.HasChildren)
        {
            return;
        }

        var extension = caches.Parent;
        caches.Remove();
        if (extension is S.WorkbookExtension && !extension.HasChildren)
        {
            var list = extension.Parent;
            extension.Remove();
            if (list is S.WorkbookExtensionList && !list.HasChildren)
            {
                list.Remove();
            }
        }
    }

    private static void RemoveSlicerFrame(Xdr.WorksheetDrawing root, string slicerName)
    {
        var frame = root.Descendants<AlternateContent>()
            .FirstOrDefault(ac => ac.Descendants()
                .Any(e => e.LocalName == "slicer" && e.NamespaceUri == SlicerDrawingNs &&
                          string.Equals(
                              e.GetAttributes().FirstOrDefault(a => a.LocalName == "name").Value,
                              slicerName,
                              StringComparison.Ordinal)));
        // Only the anchor is removed: a shared drawing part may host other
        // anchors (charts, images), so the part itself is left for ClosedXML to
        // re-emit, exactly as it preserves chart drawings across its saves.
        var anchor = frame?.Ancestors<Xdr.TwoCellAnchor>().FirstOrDefault();
        anchor?.Remove();
    }

    private static void RemoveWorksheetSlicerListEntry(WorksheetPart worksheetPart, string relId)
    {
        var extensionList = worksheetPart.Worksheet!.GetFirstChild<S.WorksheetExtensionList>();
        if (extensionList is null)
        {
            return;
        }

        foreach (var slicerRef in extensionList.Descendants<X14.SlicerRef>()
                     .Where(r => string.Equals(r.Id?.Value, relId, StringComparison.Ordinal))
                     .ToList())
        {
            var slicerList = slicerRef.Parent;
            slicerRef.Remove();
            if (slicerList is X14.SlicerList && !slicerList.HasChildren)
            {
                var extension = slicerList.Parent;
                slicerList.Remove();
                if (extension is S.WorksheetExtension && !extension.HasChildren)
                {
                    extension.Remove();
                }
            }
        }

        if (!extensionList.HasChildren)
        {
            extensionList.Remove();
        }
    }

    // ----- read (get / structure) ---------------------------------------------

    /// <summary>Reads every slicer in the workbook, keyed by host sheet, 1-based per sheet.</summary>
    public static List<SlicerInfo> Read(SpreadsheetDocument document)
    {
        var infos = new List<SlicerInfo>();
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return infos;
        }

        // Index the cache definitions by slicer name so each slicer can report
        // its source (table column or pivot field).
        var caches = workbookPart.SlicerCacheParts
            .Select(p => p.SlicerCacheDefinition)
            .Where(d => d is not null)
            .ToDictionary(d => d!.Name?.Value ?? string.Empty, d => d!, StringComparer.Ordinal);

        foreach (var sheet in workbookPart.Workbook.Descendants<S.Sheet>())
        {
            var sheetName = sheet.Name?.Value ?? string.Empty;
            if (sheet.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var index = 0;
            foreach (var slicersPart in worksheetPart.SlicersParts)
            {
                foreach (var slicer in slicersPart.Slicers?.Elements<X14.Slicer>() ?? [])
                {
                    index++;
                    var cacheName = slicer.Cache?.Value ?? slicer.Name?.Value ?? string.Empty;
                    caches.TryGetValue(cacheName, out var definition);
                    var sourceKind = definition?.GetFirstChild<X14.SlicerCachePivotTables>() is not null
                        ? "pivot"
                        : "table";
                    var source = sourceKind == "pivot"
                        ? definition?.GetFirstChild<X14.SlicerCachePivotTables>()?
                            .GetFirstChild<X14.SlicerCachePivotTable>()?.Name?.Value ?? string.Empty
                        : TableNameForCache(definition, worksheetPart);
                    infos.Add(new SlicerInfo(
                        SheetName: sheetName,
                        Index: index,
                        Path: string.Create(CultureInfo.InvariantCulture, $"/{ExcelPaths.QuoteSheet(sheetName)}/slicer[{index}]"),
                        Name: slicer.Name?.Value ?? string.Empty,
                        SourceKind: sourceKind,
                        Source: source,
                        Column: definition?.SourceName?.Value ?? string.Empty,
                        Caption: slicer.Caption?.Value ?? string.Empty));
                }
            }
        }

        return infos;
    }

    private static string TableNameForCache(X14.SlicerCacheDefinition? definition, WorksheetPart worksheetPart)
    {
        var tableId = definition?.SlicerCacheDefinitionExtensionList
            ?.Descendants<X15.TableSlicerCache>().FirstOrDefault()?.TableId?.Value;
        if (tableId is null)
        {
            return string.Empty;
        }

        var table = worksheetPart.TableDefinitionParts
            .Select(p => p.Table)
            .FirstOrDefault(t => t?.Id?.Value == tableId);
        return table?.Name?.Value ?? string.Empty;
    }

    /// <summary>Describes one slicer for <c>get</c> output.</summary>
    public static object Describe(SlicerInfo info) => new
    {
        path = info.Path,
        kind = "slicer",
        sheet = info.SheetName,
        name = info.Name,
        source = info.SourceKind,
        sourceName = info.Source,
        column = info.Column,
        caption = info.Caption,
    };

    // ----- shared helpers -----------------------------------------------------

    /// <summary>A slicer name not already used by any cache definition in the package.</summary>
    private static string UniqueName(WorkbookPart workbookPart, string baseName)
    {
        var existing = workbookPart.SlicerCacheParts
            .Select(p => p.SlicerCacheDefinition?.Name?.Value)
            .Where(n => n is not null)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (!existing.Contains(baseName))
        {
            return baseName;
        }

        for (var n = 1; ; n++)
        {
            var candidate = string.Create(CultureInfo.InvariantCulture, $"{baseName}{n}");
            if (!existing.Contains(candidate))
            {
                return candidate;
            }
        }
    }

    private static uint TableIdOrThrow(WorksheetPart worksheetPart, SlicerAddSpec spec)
    {
        // Resolve the ListObject id from the package at apply time (ClosedXML
        // assigns it on save; the op-time spec only knows the column index).
        var table = worksheetPart.TableDefinitionParts
            .Select(p => p.Table)
            .FirstOrDefault(t => string.Equals(
                t?.Name?.Value, spec.TableName, StringComparison.OrdinalIgnoreCase));
        return table?.Id?.Value ?? throw Disappeared(spec.SheetName);
    }

    private static (uint TabId, uint CacheId, int ItemCount) PivotIdsOrThrow(WorkbookPart workbookPart, SlicerAddSpec spec)
    {
        // Resolve the pivot's host-sheet tabId, its cacheId, and the distinct
        // value count of the sliced field (from the raw pivotCacheDefinition).
        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            foreach (var pivotPart in worksheetPart.PivotTableParts)
            {
                if (!string.Equals(pivotPart.PivotTableDefinition?.Name?.Value, spec.PivotName, StringComparison.Ordinal))
                {
                    continue;
                }

                var cacheId = pivotPart.PivotTableDefinition?.CacheId?.Value ?? 0u;
                var hostRelId = workbookPart.GetIdOfPart(worksheetPart);
                var hostSheet = workbookPart.Workbook!.Descendants<S.Sheet>()
                    .FirstOrDefault(s => string.Equals(s.Id?.Value, hostRelId, StringComparison.Ordinal));
                var tabId = hostSheet?.SheetId?.Value ?? 0u;
                var itemCount = SharedItemCount(pivotPart, spec.Column);
                return (tabId, cacheId, itemCount);
            }
        }

        throw Disappeared(spec.SheetName);
    }

    /// <summary>The distinct (shared-item) count of one cache field, read raw; 0 when unknown.</summary>
    private static int SharedItemCount(PivotTablePart pivotPart, string fieldName)
    {
        var cacheDefinition = pivotPart.PivotTableCacheDefinitionPart?.PivotCacheDefinition;
        var field = cacheDefinition?.CacheFields?
            .OfType<S.CacheField>()
            .FirstOrDefault(f => string.Equals(f.Name?.Value, fieldName, StringComparison.OrdinalIgnoreCase));
        var shared = field?.SharedItems;
        if (shared is null)
        {
            return 0;
        }

        // Prefer the live child count; fall back to the declared @count.
        var children = shared.ChildElements.Count;
        return children > 0 ? children : (int)(shared.Count?.Value ?? 0u);
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

        throw Disappeared(sheetName);
    }

    private static AiofficeException Disappeared(string sheetName) => new(
        ErrorCodes.InternalError,
        $"Sheet '{sheetName}' disappeared between validation and the slicer write pass.",
        "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");

    private static string Invariant(int value) => value.ToString(CultureInfo.InvariantCulture);

    private static string XmlEscape(string text) => text
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal)
        .Replace("\"", "&quot;", StringComparison.Ordinal);

    private static string RequiredString(JsonObject props, string key, int opIndex, string example)
    {
        if (props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
            value.GetValueKind() == JsonValueKind.String)
        {
            var text = value.GetValue<string>();
            if (!string.IsNullOrWhiteSpace(text))
            {
                return text;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add slicer needs props.{key}.",
            $"Pass {example}.");
    }

    private static string? OptionalString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static int? OptionalInt(JsonObject props, string key, int opIndex)
    {
        if (!props.TryGetPropertyValue(key, out var node) || node is not JsonValue value)
        {
            return null;
        }

        return value.GetValueKind() switch
        {
            JsonValueKind.Number => value.GetValue<int>(),
            JsonValueKind.String when int.TryParse(
                value.GetValue<string>(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) => parsed,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: slicer {key} must be a whole number.",
                "Pass a positive integer like 3."),
        };
    }

    private static (int Column, int Row) ParseCell(string text, int opIndex)
    {
        var match = AnchorCell().Match(text);
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: slicer x anchor '{text}' is not a cell like E2.",
                "Pass a single-cell A1 reference, e.g. {\"x\":\"E2\"}.");
        }

        var column = 0;
        foreach (var c in match.Groups[1].Value.ToUpperInvariant())
        {
            column = (column * 26) + (c - 'A' + 1);
        }

        var row = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
        return (column, row);
    }

    [System.Text.RegularExpressions.GeneratedRegex("^([A-Za-z]{1,3})([0-9]{1,7})$")]
    private static partial System.Text.RegularExpressions.Regex AnchorCell();
}
