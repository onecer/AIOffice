using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    private static readonly string[] TableProps =
        ["borders", "borderColor", "borderWidthPt", "shading", "headerRow", "columnWidths", "width", "alignment", "cellPaddingCm", "rtl"];

    private static readonly string[] CellProps =
        ["text", "mergeRight", "mergeDown", "shading", "valign"];

    private const double TwipsPerEmu = 1.0 / 635.0;

    // ------------------------------------------------------------- table set

    /// <summary>
    /// <c>{"op":"set","path":"/body/table[1]","props":{…}}</c>: table-level
    /// formatting — borders (all|outer|none) with borderColor/borderWidthPt,
    /// shading, headerRow (bold + w:tblHeader repeat-on-pages), columnWidths
    /// (tblGrid + per-cell tcW), width (length or "NN%"), alignment and
    /// cellPaddingCm (default cell margins).
    /// </summary>
    private static object ApplySetTable(Table table, ResolvedNode node, JsonObject props)
    {
        var applied = new List<string>();
        foreach (var (name, value) in props)
        {
            switch (name)
            {
                case "borders":
                    ApplyTableBorders(table, NodeToString(value), color: null, sizeEighths: null);
                    break;

                case "borderColor":
                    ApplyTableBorders(table, kind: null, WordFormatting.ParseHexColor(NodeToString(value)), sizeEighths: null);
                    break;

                case "borderWidthPt":
                    ApplyTableBorders(table, kind: null, color: null, ParseBorderWidthEighths(NodeToString(value)));
                    break;

                case "shading":
                    ApplyShading(EnsureTblPr(table), NodeToString(value), set: s => EnsureTblPr(table).Shading = s);
                    break;

                case "headerRow":
                    ApplyHeaderRow(table, node, WordFormatting.ParseBool(name, NodeToString(value)));
                    break;

                case "columnWidths":
                    ApplyColumnWidths(table, value as JsonArray ?? throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        "columnWidths must be an array with one entry per column.",
                        "Example: {\"props\":{\"columnWidths\":[\"3cm\",\"auto\",\"2.5cm\"]}}; use \"auto\" to let a column size itself."));
                    break;

                case "width":
                    ApplyTableWidth(table, NodeToString(value));
                    break;

                case "alignment":
                    EnsureTblPr(table).TableJustification = new TableJustification { Val = ParseTableAlignment(NodeToString(value)) };
                    break;

                case "cellPaddingCm":
                    ApplyCellPadding(table, NodeToString(value));
                    break;

                case "rtl":
                    EnsureTblPr(table).BiDiVisual = WordFormatting.ParseBool(name, NodeToString(value))
                        ? new BiDiVisual()
                        : new BiDiVisual { Val = OnOffOnlyValues.Off };
                    break;

                default:
                    throw new AiofficeException(
                        ErrorCodes.UnsupportedFeature,
                        $"Property '{name}' is not supported on table.",
                        $"Did you mean '{WordFormatting.Nearest(name, TableProps)}'? Supported table properties: {string.Join(", ", TableProps)}.",
                        candidates: TableProps);
            }

            applied.Add(name);
        }

        return new { op = "set", path = node.CanonicalPath, type = "table", properties = applied };
    }

    /// <summary>tblPr must be the table's first child.</summary>
    private static TableProperties EnsureTblPr(Table table)
    {
        var tblPr = table.GetFirstChild<TableProperties>();
        if (tblPr is null)
        {
            tblPr = new TableProperties();
            table.InsertAt(tblPr, 0);
        }

        return tblPr;
    }

    private static TableCellProperties EnsureTcPr(TableCell cell)
    {
        var tcPr = cell.GetFirstChild<TableCellProperties>();
        if (tcPr is null)
        {
            tcPr = new TableCellProperties();
            cell.InsertAt(tcPr, 0);
        }

        return tcPr;
    }

    /// <summary>w:sz is eighths of a point: 0.5pt → 4. Word's floor is 2 (¼pt).</summary>
    private static uint ParseBorderWidthEighths(string value)
    {
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var points) &&
            points is > 0 and <= 12)
        {
            return (uint)Math.Max(2, Math.Round(points * 8));
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"borderWidthPt must be a number of points in (0, 12], got '{value}'.",
            "Use e.g. 0.5 for hairline borders or 1.5 for heavy ones.");
    }

    /// <summary>
    /// Rebuilds or retouches w:tblBorders. A kind (all|outer|none) decides the
    /// six border slots; color/size alone retouch every non-none slot of the
    /// existing set (an absent set counts as "all", matching what create made).
    /// </summary>
    private static void ApplyTableBorders(Table table, string? kind, string? color, uint? sizeEighths)
    {
        if (kind is not (null or "all" or "outer" or "none"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown borders value '{kind}'.",
                "Use borders \"all\" (grid), \"outer\" (box only) or \"none\".",
                candidates: ["all", "outer", "none"]);
        }

        var tblPr = EnsureTblPr(table);

        if (kind is not null)
        {
            var size = sizeEighths ?? 4;
            BorderValues Outer() => kind == "none" ? BorderValues.None : BorderValues.Single;
            BorderValues Inner() => kind == "all" ? BorderValues.Single : BorderValues.None;

            tblPr.TableBorders = new TableBorders(
                Border(new TopBorder(), Outer(), size, color),
                Border(new LeftBorder(), Outer(), size, color),
                Border(new BottomBorder(), Outer(), size, color),
                Border(new RightBorder(), Outer(), size, color),
                Border(new InsideHorizontalBorder(), Inner(), size, color),
                Border(new InsideVerticalBorder(), Inner(), size, color));
            return;
        }

        var borders = tblPr.TableBorders;
        if (borders is null)
        {
            ApplyTableBorders(table, "all", color, sizeEighths);
            return;
        }

        foreach (var border in borders.ChildElements.OfType<BorderType>())
        {
            if (border.Val?.Value == BorderValues.None || border.Val?.Value == BorderValues.Nil)
            {
                continue;
            }

            if (color is not null)
            {
                border.Color = color;
            }

            if (sizeEighths is { } size)
            {
                border.Size = size;
            }
        }
    }

    private static T Border<T>(T border, BorderValues val, uint size, string? color)
        where T : BorderType
    {
        border.Val = val;
        if (val != BorderValues.None)
        {
            border.Size = size;
            if (color is not null)
            {
                border.Color = color;
            }
        }

        return border;
    }

    /// <summary>w:shd clear-pattern fill; "none" removes the shading.</summary>
    private static void ApplyShading(OpenXmlElement owner, string value, Action<Shading?> set)
    {
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            set(null);
            return;
        }

        set(new Shading
        {
            Val = ShadingPatternValues.Clear,
            Color = "auto",
            Fill = WordFormatting.ParseHexColor(value),
        });
    }

    /// <summary>
    /// headerRow true: bolds row 1 and marks it w:tblHeader so it repeats on
    /// every page; empty paragraphs get a bold empty run so later text inherits
    /// the weight. false removes the repeat mark (bolding is left alone).
    /// </summary>
    private static void ApplyHeaderRow(Table table, ResolvedNode node, bool on)
    {
        var row = table.ChildElements.OfType<TableRow>().FirstOrDefault() ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"{node.CanonicalPath} has no rows, so there is no header row to style.",
            "Add a row first: {\"op\":\"add\",\"path\":\"" + node.CanonicalPath + "\",\"type\":\"tr\",\"props\":{\"cells\":[…]}}.");

        var trPr = row.GetFirstChild<TableRowProperties>();
        if (!on)
        {
            trPr?.RemoveAllChildren<TableHeader>();
            return;
        }

        if (trPr is null)
        {
            trPr = new TableRowProperties();
            row.InsertAt(trPr, 0);
        }

        if (trPr.GetFirstChild<TableHeader>() is null)
        {
            trPr.AppendChild(new TableHeader());
        }

        foreach (var paragraph in row.Descendants<Paragraph>())
        {
            if (!paragraph.ChildElements.OfType<Run>().Any())
            {
                paragraph.AppendChild(new Run(NewText(string.Empty)));
            }

            foreach (var run in paragraph.ChildElements.OfType<Run>())
            {
                WordFormatting.SetRunProp(run, "bold", "true");
            }
        }
    }

    /// <summary>
    /// One width per grid column: lengths become fixed twips (w:gridCol + every
    /// cell's tcW, merged cells summing their columns), "auto" leaves the column
    /// self-sizing. All-fixed tables also get w:tblLayout fixed so Word honors
    /// the numbers.
    /// </summary>
    private static void ApplyColumnWidths(Table table, JsonArray widths)
    {
        var columns = GridColumnCount(table);
        if (widths.Count != columns)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"columnWidths has {widths.Count} entries but the table has {columns} column(s).",
                "Pass exactly one width per column; use \"auto\" for columns that should size themselves.");
        }

        var twips = new long?[columns]; // null = auto
        for (var i = 0; i < columns; i++)
        {
            var raw = NodeToString(widths[i]);
            twips[i] = raw.Equals("auto", StringComparison.OrdinalIgnoreCase)
                ? null
                : Math.Max(1, (long)Math.Round(ParseLengthEmu($"columnWidths[{i}]", raw) * TwipsPerEmu));
        }

        var grid = table.GetFirstChild<TableGrid>();
        if (grid is null)
        {
            grid = new TableGrid();
            var anchor = table.GetFirstChild<TableProperties>();
            if (anchor is null)
            {
                table.InsertAt(grid, 0);
            }
            else
            {
                table.InsertAfter(grid, anchor);
            }
        }

        grid.RemoveAllChildren<GridColumn>();
        foreach (var width in twips)
        {
            grid.AppendChild(width is { } w
                ? new GridColumn { Width = w.ToString(CultureInfo.InvariantCulture) }
                : new GridColumn());
        }

        foreach (var row in table.ChildElements.OfType<TableRow>())
        {
            var cursor = 0;
            foreach (var cell in row.ChildElements.OfType<TableCell>())
            {
                var span = CellSpan(cell);
                var slice = twips.Skip(cursor).Take(Math.Min(span, Math.Max(0, columns - cursor))).ToList();
                cursor += span;

                var tcPr = EnsureTcPr(cell);
                tcPr.TableCellWidth = slice.Count > 0 && slice.All(w => w is not null)
                    ? new TableCellWidth
                    {
                        Type = TableWidthUnitValues.Dxa,
                        Width = slice.Sum(w => w!.Value).ToString(CultureInfo.InvariantCulture),
                    }
                    : new TableCellWidth { Type = TableWidthUnitValues.Auto, Width = "0" };
            }
        }

        if (twips.All(w => w is not null))
        {
            EnsureTblPr(table).TableLayout = new TableLayout { Type = TableLayoutValues.Fixed };
        }
        else
        {
            EnsureTblPr(table).TableLayout = null;
        }
    }

    /// <summary>"100%"-style percentages become w:tblW pct (fiftieths of a percent); lengths become dxa.</summary>
    private static void ApplyTableWidth(Table table, string value)
    {
        var tblPr = EnsureTblPr(table);
        if (value.EndsWith('%'))
        {
            var number = value[..^1].Trim();
            if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out var pct) ||
                pct is <= 0 or > 100)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Table width '{value}' is not a valid percentage.",
                    "Use a percentage in (0, 100], e.g. \"100%\", or a length like \"12cm\".");
            }

            tblPr.TableWidth = new TableWidth
            {
                Type = TableWidthUnitValues.Pct,
                Width = ((int)Math.Round(pct * 50)).ToString(CultureInfo.InvariantCulture),
            };
            return;
        }

        var twips = Math.Max(1, (long)Math.Round(ParseLengthEmu("width", value) * TwipsPerEmu));
        tblPr.TableWidth = new TableWidth
        {
            Type = TableWidthUnitValues.Dxa,
            Width = twips.ToString(CultureInfo.InvariantCulture),
        };
    }

    private static TableRowAlignmentValues ParseTableAlignment(string value) => value.ToLowerInvariant() switch
    {
        "left" or "start" => TableRowAlignmentValues.Left,
        "center" or "centre" => TableRowAlignmentValues.Center,
        "right" or "end" => TableRowAlignmentValues.Right,
        _ => throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Unknown table alignment '{value}'.",
            "Use alignment left, center or right.",
            candidates: ["left", "center", "right"]),
    };

    /// <summary>Default cell margins (w:tblCellMar) on all four sides, in centimeters.</summary>
    private static void ApplyCellPadding(Table table, string value)
    {
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var cm) ||
            cm is < 0 or > 5)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"cellPaddingCm must be a number of centimeters in [0, 5], got '{value}'.",
                "Example: {\"props\":{\"cellPaddingCm\":0.15}}.");
        }

        var twips = (int)Math.Round(cm * TwipsPerCm);
        EnsureTblPr(table).TableCellMarginDefault = new TableCellMarginDefault(
            new TopMargin { Type = TableWidthUnitValues.Dxa, Width = twips.ToString(CultureInfo.InvariantCulture) },
            new TableCellLeftMargin { Type = TableWidthValues.Dxa, Width = (short)twips },
            new BottomMargin { Type = TableWidthUnitValues.Dxa, Width = twips.ToString(CultureInfo.InvariantCulture) },
            new TableCellRightMargin { Type = TableWidthValues.Dxa, Width = (short)twips });
    }

    // -------------------------------------------------------------- cell set

    /// <summary>
    /// <c>{"op":"set","path":"…/tc[1]","props":{…}}</c>: cell text plus the M5
    /// deep-table props — mergeRight (w:gridSpan, 1 unmerges), mergeDown
    /// (w:vMerge restart/continue chain, 1 unmerges), shading and valign.
    /// </summary>
    private static object ApplySetCell(TableCell cell, ResolvedNode node, JsonObject props)
    {
        // Merges restructure the row, so they run before cosmetic props.
        var ordered = props.OrderBy(kv => kv.Key switch { "mergeRight" => 0, "mergeDown" => 1, _ => 2 });
        var applied = new List<string>();
        foreach (var (name, value) in ordered)
        {
            switch (name)
            {
                case "text":
                    cell.RemoveAllChildren<Paragraph>();
                    cell.AppendChild(WordFactory.Paragraph(NodeToString(value)));
                    break;

                case "mergeRight":
                    MergeRight(cell, node, RequireMergeCount(name, value));
                    break;

                case "mergeDown":
                    MergeDown(cell, node, RequireMergeCount(name, value));
                    break;

                case "shading":
                    ApplyShading(EnsureTcPr(cell), NodeToString(value), set: s => EnsureTcPr(cell).Shading = s);
                    break;

                case "valign":
                    EnsureTcPr(cell).TableCellVerticalAlignment = new TableCellVerticalAlignment
                    {
                        Val = NodeToString(value).ToLowerInvariant() switch
                        {
                            "top" => TableVerticalAlignmentValues.Top,
                            "center" or "centre" or "middle" => TableVerticalAlignmentValues.Center,
                            "bottom" => TableVerticalAlignmentValues.Bottom,
                            var other => throw new AiofficeException(
                                ErrorCodes.InvalidArgs,
                                $"Unknown valign '{other}'.",
                                "Use valign top, center or bottom.",
                                candidates: ["top", "center", "bottom"]),
                        },
                    };
                    break;

                default:
                    throw new AiofficeException(
                        ErrorCodes.UnsupportedFeature,
                        $"Property '{name}' is not supported on tc.",
                        $"Did you mean '{WordFormatting.Nearest(name, CellProps)}'? Supported cell properties: {string.Join(", ", CellProps)}. " +
                        $"For run formatting, address a paragraph inside the cell: {node.CanonicalPath}/p[1].",
                        candidates: CellProps);
            }

            applied.Add(name);
        }

        var (colspan, rowspan) = CellSpans(cell);
        return new { op = "set", path = node.CanonicalPath, type = "tc", colspan, rowspan, properties = applied };
    }

    private static int RequireMergeCount(string name, JsonNode? value)
    {
        if (int.TryParse(NodeToString(value), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) && n >= 1)
        {
            return n;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"{name} must be a whole number ≥ 1, got '{NodeToString(value)}'.",
            $"{name} is the total number of cells the merged cell covers; 1 unmerges.");
    }

    // ----------------------------------------------------------------- merges

    /// <summary>Grid columns this cell spans (w:gridSpan, default 1).</summary>
    internal static int CellSpan(TableCell cell) =>
        cell.GetFirstChild<TableCellProperties>()?.GridSpan?.Val?.Value is { } span && span > 1 ? span : 1;

    /// <summary>null (no vertical merge), "restart" (head) or "continue" (swallowed slot).</summary>
    internal static string? VerticalMergeState(TableCell cell)
    {
        var vMerge = cell.GetFirstChild<TableCellProperties>()?.VerticalMerge;
        if (vMerge is null)
        {
            return null;
        }

        return vMerge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
    }

    /// <summary>The grid column (0-based) where this cell starts.</summary>
    private static int GridStart(TableCell cell)
    {
        var start = 0;
        for (var sibling = cell.PreviousSibling<TableCell>(); sibling is not null; sibling = sibling.PreviousSibling<TableCell>())
        {
            start += CellSpan(sibling);
        }

        return start;
    }

    /// <summary>Total grid columns: w:tblGrid when present, else the widest row's span sum.</summary>
    internal static int GridColumnCount(Table table)
    {
        var grid = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;
        if (grid > 0)
        {
            return grid;
        }

        return table.ChildElements.OfType<TableRow>()
            .Select(r => r.ChildElements.OfType<TableCell>().Sum(CellSpan))
            .DefaultIfEmpty(0)
            .Max();
    }

    /// <summary>
    /// mergeRight n: the cell spans n grid columns. Growing absorbs the cells to
    /// its right (their non-empty paragraphs move into the merged cell, like
    /// Word's own merge); shrinking re-inserts empty cells so the row keeps its
    /// grid width. n must land exactly on a cell boundary.
    /// </summary>
    private static void MergeRight(TableCell cell, ResolvedNode node, int n)
    {
        var current = CellSpan(cell);
        var rowColumns = cell.Parent is TableRow row ? row.ChildElements.OfType<TableCell>().Sum(CellSpan) : current;
        if (GridStart(cell) + n > rowColumns)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"mergeRight {n} runs past the end of the row ({rowColumns} grid column(s)).",
                "Merge at most up to the last column of the row.");
        }

        if (n == current)
        {
            return;
        }

        if (n > current)
        {
            var needed = n - current;
            while (needed > 0)
            {
                var next = cell.NextSibling<TableCell>()!; // bounds-checked above
                var span = CellSpan(next);
                if (span > needed)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"mergeRight {n} would split a neighboring merged cell (it spans {span} columns).",
                        $"Merge to a cell boundary instead — mergeRight {n - needed + span} covers the whole neighbor — " +
                        "or unmerge the neighbor first ({\"props\":{\"mergeRight\":1}}).");
                }

                if (VerticalMergeState(next) is not null)
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        "mergeRight cannot absorb a vertically merged cell.",
                        "Unmerge it first ({\"props\":{\"mergeDown\":1}} on its restart cell), then merge horizontally.");
                }

                foreach (var paragraph in next.ChildElements.OfType<Paragraph>().Where(p => p.InnerText.Length > 0).ToList())
                {
                    paragraph.Remove();
                    cell.AppendChild(paragraph);
                }

                next.Remove();
                needed -= span;
            }
        }
        else
        {
            for (var i = 0; i < current - n; i++)
            {
                cell.InsertAfterSelf(new TableCell(WordFactory.Paragraph(string.Empty)));
            }
        }

        var tcPr = EnsureTcPr(cell);
        tcPr.GridSpan = n > 1 ? new GridSpan { Val = n } : null;
    }

    /// <summary>
    /// mergeDown n: a w:vMerge chain n rows tall starting at this cell. The
    /// continuation slots in the rows below (same grid column; absorbed to the
    /// same width first when needed) keep existing as cells marked
    /// vMerge-continue, their non-empty content moving up — Word's model
    /// exactly. n=1 dissolves the chain.
    /// </summary>
    private static void MergeDown(TableCell cell, ResolvedNode node, int n)
    {
        if (VerticalMergeState(cell) == "continue")
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"{node.CanonicalPath} is a continuation slot of a vertical merge; target the restart cell above it.",
                "Run get on the cells above to find the one reporting rowspan > 1.");
        }

        if (cell.Parent is not TableRow row || row.Parent is not Table table)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                "The cell is not inside a table row.",
                "Re-export the file from Word.");
        }

        var rows = table.ChildElements.OfType<TableRow>().ToList();
        var rowIndex = rows.IndexOf(row);
        if (rowIndex + n > rows.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"mergeDown {n} needs {n - 1} row(s) below, but only {rows.Count - rowIndex - 1} exist.",
                "Add rows first, or merge fewer rows.");
        }

        var start = GridStart(cell);
        var span = CellSpan(cell);

        // Dissolve the existing chain first: every op then builds from a clean slate.
        for (var r = rowIndex + 1; r < rows.Count; r++)
        {
            var slot = CellAtGridStart(rows[r], start);
            if (slot is null || VerticalMergeState(slot) != "continue" || CellSpan(slot) != span)
            {
                break;
            }

            EnsureTcPr(slot).VerticalMerge = null;
        }

        EnsureTcPr(cell).VerticalMerge = null;

        if (n == 1)
        {
            return;
        }

        for (var r = rowIndex + 1; r < rowIndex + n; r++)
        {
            var slot = CellAtGridStart(rows[r], start) ?? throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Row {r + 1} has no cell starting at grid column {start + 1}, so the merge would not be rectangular.",
                "Align the rows first (matching mergeRight on the rows below), then merge down.");

            if (VerticalMergeState(slot) is not null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Row {r + 1} already participates in another vertical merge at that column.",
                    "Unmerge it first ({\"props\":{\"mergeDown\":1}} on its restart cell).");
            }

            if (CellSpan(slot) != span)
            {
                MergeRight(slot, node, span); // rectangularize: Word merges the block, so do we
            }

            foreach (var paragraph in slot.ChildElements.OfType<Paragraph>().Where(p => p.InnerText.Length > 0).ToList())
            {
                paragraph.Remove();
                cell.AppendChild(paragraph);
            }

            if (!slot.ChildElements.OfType<Paragraph>().Any())
            {
                slot.AppendChild(new Paragraph());
            }

            EnsureTcPr(slot).VerticalMerge = new VerticalMerge(); // no w:val = continue
        }

        EnsureTcPr(cell).VerticalMerge = new VerticalMerge { Val = MergedCellValues.Restart };
    }

    /// <summary>The cell of a row that starts exactly at one grid column, or null.</summary>
    private static TableCell? CellAtGridStart(TableRow row, int start)
    {
        var cursor = 0;
        foreach (var cell in row.ChildElements.OfType<TableCell>())
        {
            if (cursor == start)
            {
                return cell;
            }

            if (cursor > start)
            {
                return null;
            }

            cursor += CellSpan(cell);
        }

        return null;
    }

    /// <summary>(colspan, rowspan) of a cell; continuation slots report rowspan 0.</summary>
    internal static (int Colspan, int Rowspan) CellSpans(TableCell cell)
    {
        var colspan = CellSpan(cell);
        switch (VerticalMergeState(cell))
        {
            case "continue":
                return (colspan, 0);

            case "restart":
            {
                var rowspan = 1;
                if (cell.Parent is TableRow row && row.Parent is Table table)
                {
                    var rows = table.ChildElements.OfType<TableRow>().ToList();
                    var start = GridStart(cell);
                    for (var r = rows.IndexOf(row) + 1; r < rows.Count; r++)
                    {
                        var slot = CellAtGridStart(rows[r], start);
                        if (slot is null || VerticalMergeState(slot) != "continue")
                        {
                            break;
                        }

                        rowspan++;
                    }
                }

                return (colspan, rowspan);
            }

            default:
                return (colspan, 1);
        }
    }

    // -------------------------------------------------------------------- get

    /// <summary>get on a table: structure plus the M5 formatting surface, reopen-verifiable.</summary>
    private static Dictionary<string, object?> TableGetShape(Table table)
    {
        var tblPr = table.GetFirstChild<TableProperties>();
        var firstRow = table.ChildElements.OfType<TableRow>().FirstOrDefault();
        var grid = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();

        return new Dictionary<string, object?>
        {
            ["rows"] = table.ChildElements.OfType<TableRow>().Count(),
            ["columns"] = GridColumnCount(table),
            ["borders"] = BordersKindName(tblPr?.TableBorders),
            ["shading"] = tblPr?.Shading?.Fill?.Value,
            ["alignment"] = TableAlignmentName(tblPr?.TableJustification),
            ["width"] = TableWidthName(tblPr?.TableWidth),
            ["cellPaddingCm"] = tblPr?.TableCellMarginDefault?.TopMargin?.Width?.Value is { } w &&
                long.TryParse(w, NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips)
                    ? TwipsToCm(twips)
                    : null,
            ["headerRow"] = firstRow?.GetFirstChild<TableRowProperties>()?.GetFirstChild<TableHeader>() is not null,
            ["rtl"] = tblPr?.BiDiVisual is { } bidi && (bidi.Val?.Value ?? OnOffOnlyValues.On) == OnOffOnlyValues.On,
            ["columnWidthsCm"] = grid is { Count: > 0 }
                ? grid.Select(c => c.Width?.Value is { } cw &&
                    long.TryParse(cw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var t)
                        ? (object)TwipsToCm(t)
                        : "auto").ToList()
                : null,
        };
    }

    /// <summary>get on a tc: text/paragraphs plus {colspan,rowspan} and cosmetics.</summary>
    private static Dictionary<string, object?> CellGetShape(TableCell cell)
    {
        var (colspan, rowspan) = CellSpans(cell);
        var tcPr = cell.GetFirstChild<TableCellProperties>();
        var properties = new Dictionary<string, object?>
        {
            ["text"] = cell.InnerText,
            ["paragraphs"] = cell.ChildElements.OfType<Paragraph>().Count(),
            ["colspan"] = colspan,
            ["rowspan"] = rowspan,
            ["shading"] = tcPr?.Shading?.Fill?.Value,
            ["valign"] = CellValignName(tcPr?.TableCellVerticalAlignment),
        };

        if (rowspan == 0)
        {
            properties["note"] = "This slot is a vMerge continuation; the restart cell above holds the merged content.";
        }

        return properties;
    }

    private static string? BordersKindName(TableBorders? borders)
    {
        if (borders is null)
        {
            return null;
        }

        bool IsLine(BorderType? b) => b?.Val?.Value is { } v && v != BorderValues.None && v != BorderValues.Nil;
        var outer = new[] { (BorderType?)borders.TopBorder, borders.LeftBorder, borders.BottomBorder, borders.RightBorder };
        var inner = new[] { (BorderType?)borders.InsideHorizontalBorder, borders.InsideVerticalBorder };

        if (outer.All(IsLine))
        {
            return inner.All(IsLine) ? "all" : "outer";
        }

        return outer.Any(IsLine) || inner.Any(IsLine) ? "custom" : "none";
    }

    private static string? TableAlignmentName(TableJustification? jc)
    {
        if (jc?.Val?.Value is not { } val)
        {
            return null;
        }

        if (val == TableRowAlignmentValues.Center)
        {
            return "center";
        }

        return val == TableRowAlignmentValues.Right ? "right" : "left";
    }

    private static string? TableWidthName(TableWidth? tblW)
    {
        if (tblW?.Width?.Value is not { } raw ||
            !long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value))
        {
            return null;
        }

        if (tblW.Type?.Value == TableWidthUnitValues.Pct)
        {
            return (value / 50.0).ToString(CultureInfo.InvariantCulture) + "%";
        }

        return tblW.Type?.Value == TableWidthUnitValues.Dxa
            ? TwipsToCm(value).ToString(CultureInfo.InvariantCulture) + "cm"
            : null;
    }

    private static string? CellValignName(TableCellVerticalAlignment? vAlign)
    {
        if (vAlign?.Val?.Value is not { } val)
        {
            return null;
        }

        if (val == TableVerticalAlignmentValues.Center)
        {
            return "center";
        }

        return val == TableVerticalAlignmentValues.Bottom ? "bottom" : "top";
    }
}
