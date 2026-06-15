using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>One built-in table look, painted with direct fills/borders (never theme table styles).</summary>
internal sealed record PptxTableLook(
    string Border,
    string HeaderFill,
    string? HeaderColor,
    string BandFill,
    string PlainFill,
    string? BodyColor);

/// <summary>
/// Native pptx tables: an a:tbl inside a p:graphicFrame, addressed as
/// /slide[i]/table[k] with docx-style /tr[r]/tc[c] segments beneath. Looks are
/// painted directly (per-cell solid fills + explicit borders) so they survive
/// any theme; merges write the full gridSpan/rowSpan + hMerge/vMerge matrix
/// PowerPoint expects.
/// </summary>
internal static class PptxTables
{
    /// <summary>The direct-paint looks add table understands (each painted per-cell, theme-independent).</summary>
    public static readonly IReadOnlyList<string> Styles = ["light", "medium", "dark"];

    /// <summary>
    /// The built-in PowerPoint table-style presets (v1.4.0, additive): these map to
    /// the workbook's stock a:tableStyleId GUIDs rather than direct per-cell paint, so
    /// PowerPoint themes the table itself. "none" clears the style id. They are accepted
    /// on the same <c>style</c> prop as the direct-paint looks above; the token decides.
    /// </summary>
    public static readonly IReadOnlyList<string> StylePresets =
        ["none", "light1", "light2", "medium1", "medium2", "medium3", "dark1", "dark2"];

    /// <summary>The built-in style toggles a preset-styled table accepts (a:tblPr flags).</summary>
    public static readonly IReadOnlyList<string> StyleOptions = ["firstRow", "lastRow", "bandRow", "firstCol"];

    /// <summary>
    /// The stock built-in table-style GUIDs PowerPoint ships, keyed by preset token. These
    /// are the well-known ids the table-style part references; "none" has no id (it clears).
    /// </summary>
    private static readonly IReadOnlyDictionary<string, string> PresetStyleIds =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["light1"] = "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}", // Light Style 1
            ["light2"] = "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}", // Light Style 2 - Accent 1
            ["medium1"] = "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}", // Medium Style 1
            ["medium2"] = "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}", // Medium Style 2 - Accent 1
            ["medium3"] = "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}", // Medium Style 3
            ["dark1"] = "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}", // Dark Style 1
            ["dark2"] = "{5940675A-B579-460E-94D1-54222C63F5DA}", // Dark Style 2
        };

    private static readonly IReadOnlyList<string> AddProps =
        ["rows", "cols", "x", "y", "w", "h", "headerRow", "style", "columnWidths", "name",
         "firstRow", "lastRow", "bandRow", "firstCol"];

    private static readonly IReadOnlyList<string> CellPropKeys =
        ["text", "bold", "color", "fontSize", "align", "fill", "mergeRight", "mergeDown",
         "valign", "marginLeft", "marginRight", "marginTop", "marginBottom", "textDirection"];

    private const int MaxRows = 100;
    private const int MaxCols = 30;

    /// <summary>Cell border width (1pt).</summary>
    private const int BorderWidthEmu = 12_700;

    /// <summary>Default row height when no table height is given (1cm).</summary>
    private const long DefaultRowHeightEmu = 360_000;

    private static readonly IReadOnlyDictionary<string, PptxTableLook> Looks =
        new Dictionary<string, PptxTableLook>(StringComparer.Ordinal)
        {
            ["light"] = new("BFBFBF", "F2F2F2", "111111", "FFFFFF", "FFFFFF", null),
            ["medium"] = new("8EAADB", "4472C4", "FFFFFF", "D9E2F3", "FFFFFF", null),
            ["dark"] = new("0B1120", "0F172A", "FFFFFF", "1F2937", "374151", "FFFFFF"),
        };

    // ----- add -------------------------------------------------------------------

    /// <summary>Adds a table graphic frame and returns its 1-based table index on the slide.</summary>
    public static int Add(SlidePart slidePart, JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown table prop '{key}'.",
                    "Table props: rows, cols, x, y, w, h, headerRow, style, columnWidths, name, " +
                    "firstRow, lastRow, bandRow, firstCol.",
                    candidates: AddProps);
            }
        }

        var rows = CountProp(props, "rows", MaxRows);
        var cols = CountProp(props, "cols", MaxCols);
        var plan = ParseStylePlan(props);
        var headerRow = props.TryGetPropertyValue("headerRow", out var headerNode) && AsBool("headerRow", headerNode);
        var options = ParseStyleOptions(props, plan.PresetToken is not null, headerRow);

        var x = Length(props, "x", Units.CmToEmu(2.5));
        var y = Length(props, "y", Units.CmToEmu(2.5));
        var columnWidths = props.TryGetPropertyValue("columnWidths", out var widthsNode)
            ? ParseColumnWidths(widthsNode, cols)
            : null;
        var w = columnWidths?.Sum() ?? Length(props, "w", Units.CmToEmu(20));
        var rowHeight = props.TryGetPropertyValue("h", out var hNode)
            ? Math.Max(Units.ParseLengthEmu("h", hNode) / rows, 1)
            : DefaultRowHeightEmu;
        var h = rowHeight * rows;

        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = PptxDoc.NextShapeId(tree);
        var name = props.TryGetPropertyValue("name", out var nameNode) && nameNode is not null
            ? J.ScalarText(nameNode)
            : Units.Inv($"Table {id}");

        var grid = new A.TableGrid();
        for (var c = 0; c < cols; c++)
        {
            grid.Append(new A.GridColumn { Width = columnWidths?[c] ?? w / cols });
        }

        var table = new A.Table(BuildTableProperties(plan.PresetId, options), grid);
        for (var r = 1; r <= rows; r++)
        {
            var row = new A.TableRow { Height = rowHeight };
            for (var c = 0; c < cols; c++)
            {
                // A built-in preset themes the table itself, so its cells stay neutral
                // (one empty styled run, no direct fill/borders); a direct-paint look
                // paints each cell explicitly so it survives any theme.
                row.Append(plan.Look is { } look ? BuildCell(look, headerRow, r) : BuildNeutralCell(headerRow, r));
            }

            table.Append(row);
        }

        tree.Append(new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = w, Cy = h }),
            new A.Graphic(new A.GraphicData(table)
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/table",
            })));

        return Tables(slidePart).Count;
    }

    /// <summary>
    /// The a:tblPr for a new table: a direct-paint look records only headerRow as the
    /// FirstRow flag (the paint carries the look); a built-in preset records its
    /// a:tableStyleId GUID plus the requested first/last/band/firstCol flags. "none"
    /// keeps the flags but writes no style id.
    /// </summary>
    private static A.TableProperties BuildTableProperties(string? presetId, StyleOptionFlags options)
    {
        var properties = new A.TableProperties();
        ApplyStyleFlags(properties, options);
        if (presetId is not null)
        {
            // The a:tableStyleId is a child element (it must follow the flag attributes,
            // which the SDK orders for us).
            properties.Append(new A.TableStyleId(presetId));
        }

        return properties;
    }

    /// <summary>Writes the first/last/band/firstCol flags onto an a:tblPr (only the true ones).</summary>
    private static void ApplyStyleFlags(A.TableProperties properties, StyleOptionFlags options)
    {
        properties.FirstRow = options.FirstRow ? true : null;
        properties.LastRow = options.LastRow ? true : null;
        properties.BandRow = options.BandRow ? true : null;
        properties.FirstColumn = options.FirstCol ? true : null;
    }

    /// <summary>One neutral cell for a preset-styled table: a styled empty run, no direct fill or borders.</summary>
    private static A.TableCell BuildNeutralCell(bool headerRow, int rowIndex)
    {
        var runProperties = new A.RunProperties { Language = "en-US" };
        if (headerRow && rowIndex == 1)
        {
            runProperties.Bold = true;
        }

        return new A.TableCell(
            new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(runProperties, new A.Text(string.Empty)))),
            new A.TableCellProperties());
    }

    /// <summary>One cell painted per the look: explicit borders + solid fill + a styled empty run.</summary>
    private static A.TableCell BuildCell(PptxTableLook look, bool headerRow, int rowIndex)
    {
        var isHeader = headerRow && rowIndex == 1;
        var bodyOrdinal = rowIndex - (headerRow ? 1 : 0);
        var fill = isHeader ? look.HeaderFill : bodyOrdinal % 2 == 1 ? look.BandFill : look.PlainFill;
        var color = isHeader ? look.HeaderColor : look.BodyColor;

        var runProperties = new A.RunProperties { Language = "en-US" };
        if (isHeader)
        {
            runProperties.Bold = true;
        }

        if (color is not null)
        {
            runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
        }

        static A.SolidFill BorderFill(string hex) => new(new A.RgbColorModelHex { Val = hex });

        return new A.TableCell(
            new A.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(runProperties, new A.Text(string.Empty)))),
            new A.TableCellProperties(
                new A.LeftBorderLineProperties(BorderFill(look.Border)) { Width = BorderWidthEmu },
                new A.RightBorderLineProperties(BorderFill(look.Border)) { Width = BorderWidthEmu },
                new A.TopBorderLineProperties(BorderFill(look.Border)) { Width = BorderWidthEmu },
                new A.BottomBorderLineProperties(BorderFill(look.Border)) { Width = BorderWidthEmu },
                new A.SolidFill(new A.RgbColorModelHex { Val = fill })));
    }

    /// <summary>
    /// The chosen styling: either a direct-paint <see cref="Look"/> (light/medium/dark)
    /// or a built-in <see cref="PresetId"/> (the a:tableStyleId GUID; null for the "none"
    /// preset, which still takes flags). Exactly one of Look/PresetToken is set.
    /// </summary>
    private sealed record StylePlan(PptxTableLook? Look, string? PresetToken, string? PresetId);

    private readonly record struct StyleOptionFlags(bool FirstRow, bool LastRow, bool BandRow, bool FirstCol);

    /// <summary>
    /// Resolves props.style into a styling plan. A direct-paint look (light/medium/dark, the
    /// default) paints each cell; a built-in preset (none/light1/.../dark2) writes an
    /// a:tableStyleId and leaves the cells neutral for PowerPoint to theme.
    /// </summary>
    private static StylePlan ParseStylePlan(JsonObject props)
    {
        if (!props.TryGetPropertyValue("style", out var node) || node is null)
        {
            return new StylePlan(Looks["light"], null, null);
        }

        var raw = J.ScalarText(node).Trim();
        var token = raw.ToLowerInvariant();
        if (Looks.TryGetValue(token, out var look))
        {
            return new StylePlan(look, null, null);
        }

        if (StylePresets.Contains(token, StringComparer.Ordinal))
        {
            // "none" keeps the flags but writes no style id; every other preset maps to a GUID.
            return new StylePlan(null, token, token == "none" ? null : PresetStyleIds[token]);
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"Table style '{raw}' is not supported.",
            "Direct-paint looks: light, medium, dark. Built-in presets (a:tableStyleId): " +
            "none, light1, light2, medium1, medium2, medium3, dark1, dark2.",
            candidates: [.. Styles, .. StylePresets]);
    }

    /// <summary>
    /// Reads the first/last/band/firstCol style toggles. They only apply to a built-in preset
    /// (direct-paint looks carry their own banding); on a direct-paint look they are rejected.
    /// With a preset and no explicit flags, headerRow seeds FirstRow so the header still reads.
    /// </summary>
    private static StyleOptionFlags ParseStyleOptions(JsonObject props, bool isPreset, bool headerRow)
    {
        var anyFlag = StyleOptions.Any(props.ContainsKey);
        if (!isPreset)
        {
            if (anyFlag)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    "firstRow/lastRow/bandRow/firstCol apply only to a built-in table preset.",
                    "Pass a preset style (e.g. {\"style\":\"medium2\",\"bandRow\":true}); the direct-paint " +
                    "looks (light/medium/dark) already band their rows.",
                    candidates: StyleOptions);
            }

            // A direct-paint look still records headerRow as the FirstRow flag (the paint
            // carries the look; the flag keeps get/structure and the existing contract honest).
            return new StyleOptionFlags(FirstRow: headerRow, LastRow: false, BandRow: false, FirstCol: false);
        }

        bool Flag(string key, bool fallback) =>
            props.TryGetPropertyValue(key, out var node) ? AsBool(key, node) : fallback;

        // Default the header band on when headerRow was asked for; the rest default off.
        return new StyleOptionFlags(
            FirstRow: Flag("firstRow", headerRow),
            LastRow: Flag("lastRow", false),
            BandRow: Flag("bandRow", false),
            FirstCol: Flag("firstCol", false));
    }

    private static int CountProp(JsonObject props, string key, int max)
    {
        if (props.TryGetPropertyValue(key, out var node) &&
            TryInteger(node, out var number) && number >= 1 && number <= max)
        {
            return number;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"add table needs '{key}' as an integer between 1 and {max}; got {node?.ToJsonString() ?? "null"}.",
            "The full shape is {\"op\":\"add\",\"path\":\"/slide[2]\",\"type\":\"table\"," +
            "\"props\":{\"rows\":3,\"cols\":4,\"x\":\"2cm\",\"y\":\"5cm\",\"w\":\"28cm\",\"headerRow\":true}}.");
    }

    /// <summary>Integers arrive as JSON numbers from hand-written ops and as strings via the CLI sugar.</summary>
    private static bool TryInteger(JsonNode? node, out int number)
    {
        number = 0;
        if (node is not JsonValue value)
        {
            return false;
        }

        double parsed;
        if (!Units.TryNumber(value, out parsed))
        {
            if (!value.TryGetValue<string>(out var raw) ||
                !double.TryParse(raw, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out parsed))
            {
                return false;
            }
        }

        if (parsed != Math.Floor(parsed))
        {
            return false;
        }

        number = (int)parsed;
        return true;
    }

    private static long Length(JsonObject props, string key, long fallback) =>
        props.TryGetPropertyValue(key, out var node) ? Units.ParseLengthEmu(key, node) : fallback;

    private static List<long> ParseColumnWidths(JsonNode? node, int cols)
    {
        if (node is not JsonArray array || array.Count != cols)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"props.columnWidths must be an array with one width per column ({cols}); got {node?.ToJsonString() ?? "null"}."),
                "Pass lengths like {\"columnWidths\":[\"8cm\",\"5cm\",\"5cm\"]} — numbers are centimeters.");
        }

        var widths = new List<long>(array.Count);
        for (var i = 0; i < array.Count; i++)
        {
            var width = Units.ParseLengthEmu(Units.Inv($"columnWidths[{i}]"), array[i]);
            if (width < 1)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    Units.Inv($"props.columnWidths[{i}] must be positive."),
                    "Every column needs a positive width, e.g. \"5cm\".");
            }

            widths.Add(width);
        }

        return widths;
    }

    // ----- enumerate / resolve -------------------------------------------------

    /// <summary>The a:tbl a graphic frame hosts; null when the element hosts no table.</summary>
    public static A.Table? TableOf(OpenXmlCompositeElement element) =>
        (element as P.GraphicFrame)?.Graphic?.GraphicData?.GetFirstChild<A.Table>();

    /// <summary>Table-hosting frames on a slide in paint order; table indices are 1-based.</summary>
    public static List<(int Index, ShapeView View, A.Table Table)> Tables(SlidePart slidePart)
    {
        var result = new List<(int, ShapeView, A.Table)>();
        foreach (var view in PptxDoc.Shapes(slidePart))
        {
            if (TableOf(view.Element) is { } table)
            {
                result.Add((result.Count + 1, view, table));
            }
        }

        return result;
    }

    /// <summary>The 1-based table index of a table-hosting frame on its slide; null for non-tables.</summary>
    public static int? IndexOf(SlidePart slidePart, OpenXmlCompositeElement element) =>
        Tables(slidePart).Where(t => ReferenceEquals(t.View.Element, element))
            .Select(t => (int?)t.Index)
            .FirstOrDefault();

    /// <summary>Resolves /slide[i]/table[k] or throws invalid_path with candidates.</summary>
    public static (int Index, ShapeView View, A.Table Table) Resolve(SlidePart slidePart, PptxAddress address)
    {
        var tables = Tables(slidePart);
        var index = address.TableIndex!.Value;
        if (index >= 1 && index <= tables.Count)
        {
            return tables[index - 1];
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            Units.Inv($"No table {index} on slide {address.SlideIndex}; it has {tables.Count} table(s)."),
            tables.Count > 0
                ? "Table indices are 1-based per slide; run 'aioffice get <file> /slide[i]' to list them."
                : "Add one first: {\"op\":\"add\",\"path\":\"" + address.CanonicalSlidePath + "\",\"type\":\"table\"," +
                  "\"props\":{\"rows\":3,\"cols\":4}}.",
            candidates: [.. tables.Take(10).Select(t => Units.Inv($"{address.CanonicalSlidePath}/table[{t.Index}]"))]);
    }

    private static A.TableRow ResolveRow(A.Table table, PptxAddress address)
    {
        var rows = table.Elements<A.TableRow>().ToList();
        var index = address.TableRowIndex!.Value;
        if (index >= 1 && index <= rows.Count)
        {
            return rows[index - 1];
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            Units.Inv($"No row {index} in {address.CanonicalTablePath}; it has {rows.Count} row(s)."),
            "Row indices are 1-based; run 'aioffice get' on the table path to list rows and cells.",
            candidates: [.. Enumerable.Range(1, Math.Min(rows.Count, 10))
                .Select(r => Units.Inv($"{address.CanonicalTablePath}/tr[{r}]"))]);
    }

    private static A.TableCell ResolveCell(A.TableRow row, PptxAddress address)
    {
        var cells = row.Elements<A.TableCell>().ToList();
        var index = address.TableCellIndex!.Value;
        if (index >= 1 && index <= cells.Count)
        {
            return cells[index - 1];
        }

        throw new AiofficeException(
            ErrorCodes.InvalidPath,
            Units.Inv($"No cell {index} in {address.CanonicalTablePath}/tr[{address.TableRowIndex}]; the row has {cells.Count} cell(s)."),
            "Cell indices are 1-based grid positions (merged-away cells still count).",
            candidates: [.. Enumerable.Range(1, Math.Min(cells.Count, 10))
                .Select(c => Units.Inv($"{address.CanonicalTablePath}/tr[{address.TableRowIndex}]/tc[{c}]"))]);
    }

    /// <summary>Paragraph-joined text of a table cell.</summary>
    public static string CellText(A.TableCell cell) => cell.TextBody is null
        ? string.Empty
        : string.Join('\n', cell.TextBody.Elements<A.Paragraph>().Select(PptxDoc.ParagraphText));

    /// <summary>All cell texts of a table in reading order (for query/snippets).</summary>
    public static string TableText(A.Table table) => string.Join(
        '\n',
        table.Elements<A.TableRow>()
            .SelectMany(r => r.Elements<A.TableCell>())
            .Select(CellText)
            .Where(t => t.Length > 0));

    /// <summary>True when the cell is covered by a merge originating elsewhere.</summary>
    public static bool IsCovered(A.TableCell cell) =>
        cell.HorizontalMerge?.Value == true || cell.VerticalMerge?.Value == true;

    // ----- set -------------------------------------------------------------------

    /// <summary>set on /slide[i]/table[k] (columnWidths) or a /tr[r]/tc[c] cell beneath it.</summary>
    public static string Set(PresentationPart presentation, PptxAddress address, JsonObject props)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var (_, view, table) = Resolve(slidePart, address);

        if (address.TableCellIndex is not null)
        {
            return SetCell(table, address, props);
        }

        if (address.TableRowIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Rows have no settable props; target a cell or the table.",
                "Set cells (/slide[i]/table[k]/tr[r]/tc[c] with text/fill/...) or the table " +
                "(columnWidths); use add/remove with type row to restructure.");
        }

        // style/firstRow/lastRow/bandRow/firstCol restyle the a:tblPr (a built-in preset);
        // they are applied together after validating the whole batch.
        var hasStyle = props.ContainsKey("style");
        var hasFlag = StyleOptions.Any(props.ContainsKey);
        foreach (var (key, value) in props)
        {
            switch (key)
            {
                case "columnWidths":
                    SetColumnWidths(table, view, value);
                    break;
                case "style" or "firstRow" or "lastRow" or "bandRow" or "firstCol":
                    break; // handled below, once
                default:
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Prop '{key}' does not apply to a table.",
                        "Table sets take columnWidths and style (a built-in preset, with firstRow/lastRow/bandRow/firstCol " +
                        "toggles); position/size/name sets target its shape path (" +
                        view.CanonicalPath(address.SlideIndex) + "); text/fill target cells.",
                        candidates: ["columnWidths", "style", .. StyleOptions]);
            }
        }

        if (hasStyle || hasFlag)
        {
            RestyleTable(table, props, hasStyle, hasFlag);
        }

        if (props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"set on '{address.Raw}' has no props.",
                "Pass {\"columnWidths\":[\"8cm\",\"5cm\"]} or {\"style\":\"medium2\"} on the table, or target a cell for text/fill.");
        }

        return address.CanonicalTablePath;
    }

    /// <summary>
    /// Restyles an existing table's a:tblPr in place: a new built-in preset rewrites its
    /// a:tableStyleId ("none" clears it), and the first/last/band/firstCol toggles update the
    /// flags. Only a built-in preset may carry the toggles — direct-paint looks are rejected
    /// (their banding is painted, not flagged). The cells keep their content.
    /// </summary>
    private static void RestyleTable(A.Table table, JsonObject props, bool hasStyle, bool hasFlag)
    {
        var properties = table.TableProperties ??= new A.TableProperties();

        string? presetToken = null;
        if (hasStyle)
        {
            var raw = props["style"] is { } node ? J.ScalarText(node).Trim() : string.Empty;
            var token = raw.ToLowerInvariant();
            if (!StylePresets.Contains(token, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    Looks.ContainsKey(token)
                        ? $"Table style '{raw}' is a direct-paint look — it can be set only when adding a table."
                        : $"Table style '{raw}' is not supported.",
                    "set style takes a built-in preset (a:tableStyleId): none, light1, light2, medium1, medium2, " +
                    "medium3, dark1, dark2. Re-add the table to switch to a direct-paint look (light/medium/dark).",
                    candidates: StylePresets);
            }

            presetToken = token;
            foreach (var id in properties.Elements<A.TableStyleId>().ToList())
            {
                id.Remove();
            }

            if (token != "none")
            {
                properties.Append(new A.TableStyleId(PresetStyleIds[token]));
            }
        }
        else if (properties.GetFirstChild<A.TableStyleId>() is null)
        {
            // Flags alone require a preset table; without an a:tableStyleId there is no
            // built-in style for them to drive.
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "firstRow/lastRow/bandRow/firstCol need a built-in table preset.",
                "Set a preset first (e.g. {\"style\":\"medium2\"}); a direct-paint table (light/medium/dark) " +
                "bands its rows by paint, not flags.");
        }

        if (hasFlag)
        {
            bool Flag(string key, bool current) =>
                props.TryGetPropertyValue(key, out var node) ? AsBool(key, node) : current;
            ApplyStyleFlags(properties, new StyleOptionFlags(
                FirstRow: Flag("firstRow", properties.FirstRow?.Value ?? false),
                LastRow: Flag("lastRow", properties.LastRow?.Value ?? false),
                BandRow: Flag("bandRow", properties.BandRow?.Value ?? false),
                FirstCol: Flag("firstCol", properties.FirstColumn?.Value ?? false)));
        }
    }

    /// <summary>Rewrites the grid column widths (and keeps the frame width in sync).</summary>
    private static void SetColumnWidths(A.Table table, ShapeView view, JsonNode? value)
    {
        var grid = table.TableGrid ?? throw CorruptTable("a:tblGrid is missing");
        var columns = grid.Elements<A.GridColumn>().ToList();
        var widths = ParseColumnWidths(value, columns.Count);
        for (var i = 0; i < columns.Count; i++)
        {
            columns[i].Width = widths[i];
        }

        if (view.Element is P.GraphicFrame { Transform.Extents: { } extents })
        {
            extents.Cx = widths.Sum();
        }
    }

    private static string SetCell(A.Table table, PptxAddress address, JsonObject props)
    {
        var row = ResolveRow(table, address);
        var cell = ResolveCell(row, address);
        var cellPath = Units.Inv($"{address.CanonicalTablePath}/tr[{address.TableRowIndex}]/tc[{address.TableCellIndex}]");

        foreach (var (key, _) in props)
        {
            if (!CellPropKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown cell prop '{key}'.",
                    "Cell props: text, bold, color, fontSize, align, fill, mergeRight, mergeDown, " +
                    "valign (top/middle/bottom), marginLeft/marginRight/marginTop/marginBottom (lengths), " +
                    "textDirection (horizontal/vertical).",
                    candidates: CellPropKeys);
            }
        }

        var mergeRight = MergeCount(props, "mergeRight");
        var mergeDown = MergeCount(props, "mergeDown");
        if (mergeRight > 0 || mergeDown > 0)
        {
            Merge(table, address.TableRowIndex!.Value, address.TableCellIndex!.Value, mergeRight, mergeDown);
        }
        else if (IsCovered(cell))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{cellPath}' is covered by a merge; it has no visible content.",
                "Target the merge's origin cell instead — run 'aioffice get' on the table path to see merges.");
        }

        foreach (var (key, value) in props)
        {
            switch (key)
            {
                case "text":
                    SetCellText(cell, value is null ? string.Empty : J.ScalarText(value));
                    break;
                case "bold":
                    ApplyCellRunProps(cell, rPr => rPr.Bold = AsBool(key, value));
                    break;
                case "color":
                    var hex = Units.ParseColorHex(key, value);
                    ApplyCellRunProps(cell, rPr => SetRunColor(rPr, hex));
                    break;
                case "fontSize":
                    ApplyCellRunProps(cell, rPr => rPr.FontSize = Units.ParseFontSizeHundredths(key, value));
                    break;
                case "align":
                    var alignment = PptxEditor.ParseAlign(value) ?? throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"Not a valid align value: {value?.ToJsonString() ?? "null"}",
                        "Use left, center, right or justify.",
                        candidates: ["left", "center", "right", "justify"]);
                    foreach (var paragraph in cell.TextBody?.Elements<A.Paragraph>() ?? [])
                    {
                        var pPr = paragraph.ParagraphProperties;
                        if (pPr is null)
                        {
                            pPr = new A.ParagraphProperties();
                            paragraph.InsertAt(pPr, 0);
                        }

                        pPr.Alignment = alignment;
                    }

                    break;
                case "fill":
                    SetCellFill(cell, Units.ParseColorHex(key, value));
                    break;
                case "valign":
                    EnsureCellProperties(cell).Anchor = ParseVAlign(value);
                    break;
                case "marginLeft":
                    EnsureCellProperties(cell).LeftMargin = ParseMarginEmu(key, value);
                    break;
                case "marginRight":
                    EnsureCellProperties(cell).RightMargin = ParseMarginEmu(key, value);
                    break;
                case "marginTop":
                    EnsureCellProperties(cell).TopMargin = ParseMarginEmu(key, value);
                    break;
                case "marginBottom":
                    EnsureCellProperties(cell).BottomMargin = ParseMarginEmu(key, value);
                    break;
                case "textDirection":
                    EnsureCellProperties(cell).Vertical = ParseTextDirection(value);
                    break;
                default:
                    break; // mergeRight/mergeDown already handled
            }
        }

        return cellPath;
    }

    /// <summary>The cell's a:tcPr, created (before any fill child, which the SDK keeps in order) when absent.</summary>
    private static A.TableCellProperties EnsureCellProperties(A.TableCell cell) =>
        cell.TableCellProperties ??= cell.AppendChild(new A.TableCellProperties());

    /// <summary>Parses valign top/middle/bottom into the a:tcPr anchor (TextAnchoringType).</summary>
    private static A.TextAnchoringTypeValues ParseVAlign(JsonNode? value)
    {
        var token = J.ScalarText(value ?? string.Empty).Trim().ToLowerInvariant();
        return token switch
        {
            "top" => A.TextAnchoringTypeValues.Top,
            "middle" or "center" or "centre" => A.TextAnchoringTypeValues.Center,
            "bottom" => A.TextAnchoringTypeValues.Bottom,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid valign value: {value?.ToJsonString() ?? "null"}",
                "Use top, middle or bottom.",
                candidates: ["top", "middle", "bottom"]),
        };
    }

    /// <summary>Parses textDirection horizontal/vertical into the a:tcPr vert (TextVertical).</summary>
    private static A.TextVerticalValues ParseTextDirection(JsonNode? value)
    {
        var token = J.ScalarText(value ?? string.Empty).Trim().ToLowerInvariant();
        return token switch
        {
            "horizontal" => A.TextVerticalValues.Horizontal,
            "vertical" or "vertical90" or "vert" => A.TextVerticalValues.Vertical,
            "vertical270" => A.TextVerticalValues.Vertical270,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Not a valid textDirection value: {value?.ToJsonString() ?? "null"}",
                "Use horizontal or vertical.",
                candidates: ["horizontal", "vertical"]),
        };
    }

    /// <summary>Parses a cell-margin length into EMU (the a:tcPr margin attributes are Int32 EMU).</summary>
    private static int ParseMarginEmu(string key, JsonNode? value)
    {
        var emu = Units.ParseLengthEmu(key, value);
        if (emu < 0 || emu > int.MaxValue)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cell margin '{key}' is out of range: {value?.ToJsonString() ?? "null"}.",
                "Use a small non-negative length like \"0.2cm\" or \"6pt\".");
        }

        return (int)emu;
    }

    private static int MergeCount(JsonObject props, string key)
    {
        if (!props.TryGetPropertyValue(key, out var node))
        {
            return 0;
        }

        if (TryInteger(node, out var number) && number >= 1)
        {
            return number;
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Prop '{key}' must be a positive integer (the number of cells to absorb); got {node?.ToJsonString() ?? "null"}.",
            "Example: {\"mergeRight\":1} merges this cell with the one to its right (gridSpan 2).");
    }

    /// <summary>Replaces the cell text, keeping the first run's formatting as the prototype.</summary>
    private static void SetCellText(A.TableCell cell, string text)
    {
        var body = cell.TextBody;
        if (body is null)
        {
            body = new A.TextBody(new A.BodyProperties(), new A.ListStyle());
            cell.InsertAt(body, 0);
        }

        var runPrototype = body.Descendants<A.RunProperties>().FirstOrDefault();
        var paragraphPrototype = body.Elements<A.Paragraph>().FirstOrDefault()?.ParagraphProperties;

        foreach (var paragraph in body.Elements<A.Paragraph>().ToList())
        {
            paragraph.Remove();
        }

        foreach (var line in text.Split('\n'))
        {
            var paragraph = new A.Paragraph();
            if (paragraphPrototype is not null)
            {
                paragraph.Append((A.ParagraphProperties)paragraphPrototype.CloneNode(true));
            }

            var run = new A.Run();
            if (runPrototype is not null)
            {
                run.Append((A.RunProperties)runPrototype.CloneNode(true));
            }

            run.Append(new A.Text(line));
            paragraph.Append(run);
            body.Append(paragraph);
        }
    }

    private static void ApplyCellRunProps(A.TableCell cell, Action<A.RunProperties> mutate)
    {
        foreach (var run in cell.TextBody?.Descendants<A.Run>() ?? [])
        {
            var runProperties = run.RunProperties;
            if (runProperties is null)
            {
                runProperties = new A.RunProperties { Language = "en-US" };
                run.InsertAt(runProperties, 0);
            }

            mutate(runProperties);
        }
    }

    private static void SetRunColor(A.RunProperties runProperties, string hex)
    {
        foreach (var fill in runProperties.ChildElements.Where(c => c is A.SolidFill or A.NoFill or A.GradientFill).ToList())
        {
            fill.Remove();
        }

        runProperties.InsertAt(new A.SolidFill(new A.RgbColorModelHex { Val = hex }), 0);
    }

    private static void SetCellFill(A.TableCell cell, string hex)
    {
        var properties = cell.TableCellProperties;
        if (properties is null)
        {
            properties = new A.TableCellProperties();
            cell.Append(properties);
        }

        foreach (var fill in properties.ChildElements
            .Where(c => c is A.NoFill or A.SolidFill or A.GradientFill or A.BlipFill or A.PatternFill or A.GroupFill)
            .ToList())
        {
            fill.Remove();
        }

        properties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = hex }));
    }

    // ----- merge -------------------------------------------------------------------

    /// <summary>
    /// Merges a block of (down+1) x (right+1) cells anchored at the 1-based
    /// (row, col): the origin gets gridSpan/rowSpan, covered cells get
    /// hMerge/vMerge (first merged row keeps rowSpan, first merged column keeps
    /// gridSpan — the exact matrix PowerPoint writes). Covered cells' text moves
    /// into the origin.
    /// </summary>
    private static void Merge(A.Table table, int row, int col, int right, int down)
    {
        var rows = table.Elements<A.TableRow>().ToList();
        var cols = table.TableGrid?.Elements<A.GridColumn>().Count() ?? 0;
        if (col + right > cols || row + down > rows.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"The merge block exceeds the table: it spans to row {row + down}/column {col + right} but the table is {rows.Count}x{cols}."),
                "Shrink mergeRight/mergeDown so the block stays inside the grid.");
        }

        var block = new List<(int R, int C, A.TableCell Cell)>();
        for (var r = row; r <= row + down; r++)
        {
            var rowCells = rows[r - 1].Elements<A.TableCell>().ToList();
            for (var c = col; c <= col + right; c++)
            {
                if (c > rowCells.Count)
                {
                    throw CorruptTable(Units.Inv($"row {r} has only {rowCells.Count} cell(s) for a {cols}-column grid"));
                }

                var cell = rowCells[c - 1];
                if ((cell.GridSpan?.Value ?? 1) > 1 || (cell.RowSpan?.Value ?? 1) > 1 || IsCovered(cell))
                {
                    throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        Units.Inv($"The merge overlaps an existing merge at row {r}, cell {c}."),
                        "Merges cannot overlap; run 'aioffice get' on the table path to see current spans.");
                }

                block.Add((r, c, cell));
            }
        }

        var origin = block[0].Cell;
        foreach (var (r, c, cell) in block)
        {
            if (r == row && c == col)
            {
                if (right > 0)
                {
                    cell.GridSpan = right + 1;
                }

                if (down > 0)
                {
                    cell.RowSpan = down + 1;
                }

                continue;
            }

            // Move any covered text into the origin before hiding the cell.
            if (CellText(cell).Trim().Length > 0 && cell.TextBody is { } coveredBody && origin.TextBody is { } originBody)
            {
                foreach (var paragraph in coveredBody.Elements<A.Paragraph>().ToList())
                {
                    if (PptxDoc.ParagraphText(paragraph).Trim().Length > 0)
                    {
                        paragraph.Remove();
                        originBody.Append(paragraph);
                    }
                }
            }

            ClearCellText(cell);

            if (r == row)
            {
                cell.HorizontalMerge = true;
                if (down > 0)
                {
                    cell.RowSpan = down + 1;
                }
            }
            else if (c == col)
            {
                cell.VerticalMerge = true;
                if (right > 0)
                {
                    cell.GridSpan = right + 1;
                }
            }
            else
            {
                cell.HorizontalMerge = true;
                cell.VerticalMerge = true;
            }
        }
    }

    /// <summary>Resets a cell's body to one empty paragraph (keeping the first run's formatting).</summary>
    private static void ClearCellText(A.TableCell cell)
    {
        if (cell.TextBody is not { } body)
        {
            return;
        }

        var runPrototype = body.Descendants<A.RunProperties>().FirstOrDefault();
        foreach (var paragraph in body.Elements<A.Paragraph>().ToList())
        {
            paragraph.Remove();
        }

        var run = new A.Run();
        if (runPrototype is not null)
        {
            run.Append((A.RunProperties)runPrototype.CloneNode(true));
        }

        run.Append(new A.Text(string.Empty));
        body.Append(new A.Paragraph(run));
    }

    // ----- rows ---------------------------------------------------------------------

    /// <summary>
    /// add type:row — /table[k] appends, /table[k]/tr[r] inserts so the new row
    /// becomes tr[r] (position "after" inserts below it). The inserted row clones
    /// the look of the row at the spot, minus text and merges.
    /// </summary>
    public static string AddRow(PresentationPart presentation, PptxAddress address, string? position)
    {
        if (address.TableCellIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add row targets the table or a row position, not '{address.Raw}'.",
                "Use {\"op\":\"add\",\"path\":\"" + address.CanonicalTablePath + "/tr[2]\",\"type\":\"row\"} — " +
                "the new row becomes tr[2]; omit /tr to append.");
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var (_, view, table) = Resolve(slidePart, address);
        var rows = table.Elements<A.TableRow>().ToList();

        var anchor = address.TableRowIndex ?? rows.Count + 1;
        var target = position?.Trim().ToLowerInvariant() switch
        {
            null or "" or "at" or "before" => anchor,
            "after" => anchor + 1,
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown position '{position}' for add row.",
                "Use \"at\"/\"before\" (new row takes the path's index) or \"after\".",
                candidates: ["at", "before", "after"]),
        };

        if (target < 1 || target > rows.Count + 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                Units.Inv($"Cannot insert a row at position {target}; the table has {rows.Count} row(s)."),
                Units.Inv($"Valid positions are 1..{rows.Count + 1} (target {address.CanonicalTablePath}/tr[{rows.Count + 1}] to append)."));
        }

        // Inserting above a vMerge continuation would split the merge block.
        if (target <= rows.Count &&
            rows[target - 1].Elements<A.TableCell>().Any(c => c.VerticalMerge?.Value == true))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Inserting at row {target} would split a vertical merge."),
                "Insert above the merge's origin or below its last row, or rebuild the table without the merge.");
        }

        var template = rows[Math.Min(target, rows.Count) - 1];
        var row = (A.TableRow)template.CloneNode(true);
        foreach (var cell in row.Elements<A.TableCell>())
        {
            cell.GridSpan = null;
            cell.RowSpan = null;
            cell.HorizontalMerge = null;
            cell.VerticalMerge = null;
            ClearCellText(cell);
        }

        if (target > rows.Count)
        {
            table.Append(row);
        }
        else
        {
            table.InsertBefore(row, rows[target - 1]);
        }

        SyncFrameHeight(view, table);
        return Units.Inv($"{address.CanonicalTablePath}/tr[{target}]");
    }

    // ----- remove -------------------------------------------------------------------

    /// <summary>remove /slide[i]/table[k] (the whole frame) or /tr[r] (one row).</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        if (address.TableCellIndex is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Cells cannot be removed; the grid stays rectangular.",
                "Clear the text ({\"op\":\"set\",...,\"props\":{\"text\":\"\"}}), merge it away, or remove the row.");
        }

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var (_, view, table) = Resolve(slidePart, address);

        if (address.TableRowIndex is null)
        {
            view.Element.Remove();
            return address.CanonicalTablePath;
        }

        var rows = table.Elements<A.TableRow>().ToList();
        var row = ResolveRow(table, address);
        if (rows.Count == 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "Cannot remove the table's only row.",
                "Remove the table itself: {\"op\":\"remove\",\"path\":\"" + address.CanonicalTablePath + "\"}.");
        }

        if (row.Elements<A.TableCell>().Any(c => (c.RowSpan?.Value ?? 1) > 1 || c.VerticalMerge?.Value == true))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                Units.Inv($"Row {address.TableRowIndex} is part of a vertical merge and cannot be removed."),
                "Rebuild the table without the merge, or remove the whole merged block's rows in PowerPoint.");
        }

        row.Remove();
        SyncFrameHeight(view, table);
        return Units.Inv($"{address.CanonicalTablePath}/tr[{address.TableRowIndex}]");
    }

    /// <summary>Keeps the frame height in sync with the sum of row heights after row edits.</summary>
    private static void SyncFrameHeight(ShapeView view, A.Table table)
    {
        if (view.Element is P.GraphicFrame { Transform.Extents: { } extents })
        {
            extents.Cy = table.Elements<A.TableRow>().Sum(r => r.Height?.Value ?? DefaultRowHeightEmu);
        }
    }

    // ----- projections ----------------------------------------------------------------

    /// <summary>The `get` projection for table, row and cell paths.</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var (index, view, table) = Resolve(slidePart, address);
        var tablePath = Units.Inv($"/slide[{address.SlideIndex}]/table[{index}]");

        if (address.TableCellIndex is not null)
        {
            var row = ResolveRow(table, address);
            var cell = ResolveCell(row, address);
            return CellProjection(cell, tablePath, address.TableRowIndex!.Value, address.TableCellIndex.Value, table);
        }

        if (address.TableRowIndex is not null)
        {
            var row = ResolveRow(table, address);
            return new
            {
                Path = Units.Inv($"{tablePath}/tr[{address.TableRowIndex}]"),
                Index = address.TableRowIndex.Value,
                HeightCm = row.Height?.Value is { } height ? Units.EmuToCm(height) : (double?)null,
                Cells = row.Elements<A.TableCell>()
                    .Select((cell, c) => CellProjection(cell, tablePath, address.TableRowIndex.Value, c + 1, table))
                    .ToList(),
            };
        }

        var rows = table.Elements<A.TableRow>().ToList();
        var columns = table.TableGrid?.Elements<A.GridColumn>().ToList() ?? [];
        var geometry = PptxDoc.Geometry(view.Element);
        return new
        {
            Path = tablePath,
            ShapePath = view.CanonicalPath(address.SlideIndex),
            Slide = address.SlideIndex,
            Id = view.Id,
            Kind = "table",
            Rows = rows.Count,
            Cols = columns.Count,
            HeaderRow = table.TableProperties?.FirstRow?.Value == true,
            Style = StyleTokenOf(table),
            StyleOptions = StyleOptionsOf(table),
            ColumnWidths = columns.Select(c => Units.EmuToCm(c.Width?.Value ?? 0)).ToList(),
            RowHeights = rows.Select(r => Units.EmuToCm(r.Height?.Value ?? 0)).ToList(),
            X = geometry is { } g1 ? Units.EmuToCm(g1.X) : (double?)null,
            Y = geometry is { } g2 ? Units.EmuToCm(g2.Y) : (double?)null,
            W = geometry is { } g3 ? Units.EmuToCm(g3.Cx) : (double?)null,
            H = geometry is { } g4 ? Units.EmuToCm(g4.Cy) : (double?)null,
            ZIndex = view.Ordinal,
            RowsDetail = rows.Select((row, r) => (object)new
            {
                Path = Units.Inv($"{tablePath}/tr[{r + 1}]"),
                Cells = row.Elements<A.TableCell>()
                    .Select((cell, c) => CellProjection(cell, tablePath, r + 1, c + 1, table))
                    .ToList(),
            }).ToList(),
        };
    }

    /// <summary>
    /// The built-in preset token a table carries (none/light1/.../dark2), or null when it
    /// is a direct-paint look (no a:tableStyleId) or carries an unrecognized GUID.
    /// </summary>
    private static string? StyleTokenOf(A.Table table)
    {
        var id = table.TableProperties?.GetFirstChild<A.TableStyleId>()?.Text;
        if (id is null)
        {
            return null;
        }

        foreach (var (token, guid) in PresetStyleIds)
        {
            if (string.Equals(id, guid, StringComparison.OrdinalIgnoreCase))
            {
                return token;
            }
        }

        return null; // a foreign/custom table style GUID: truthfully not one of our presets
    }

    /// <summary>The a:tblPr first/last/band/firstCol flags as a small object, or null when all are off.</summary>
    private static object? StyleOptionsOf(A.Table table)
    {
        var properties = table.TableProperties;
        if (properties is null)
        {
            return null;
        }

        var firstRow = properties.FirstRow?.Value == true;
        var lastRow = properties.LastRow?.Value == true;
        var bandRow = properties.BandRow?.Value == true;
        var firstCol = properties.FirstColumn?.Value == true;
        if (!firstRow && !lastRow && !bandRow && !firstCol)
        {
            return null;
        }

        return new
        {
            FirstRow = firstRow ? true : (bool?)null,
            LastRow = lastRow ? true : (bool?)null,
            BandRow = bandRow ? true : (bool?)null,
            FirstCol = firstCol ? true : (bool?)null,
        };
    }

    private static object CellProjection(A.TableCell cell, string tablePath, int row, int col, A.Table table)
    {
        var covered = IsCovered(cell);
        var tcPr = cell.TableCellProperties;
        return new
        {
            Path = Units.Inv($"{tablePath}/tr[{row}]/tc[{col}]"),
            Row = row,
            Col = col,
            Text = CellText(cell),
            GridSpan = (cell.GridSpan?.Value ?? 1) > 1 ? cell.GridSpan!.Value : (int?)null,
            RowSpan = (cell.RowSpan?.Value ?? 1) > 1 ? cell.RowSpan!.Value : (int?)null,
            Covered = covered ? true : (bool?)null,
            MergedInto = covered ? MergeOriginPath(table, tablePath, row, col) : null,
            Fill = CellFillHex(cell),
            VAlign = VAlignToken(tcPr?.Anchor?.Value),
            MarginLeft = tcPr?.LeftMargin is { } ml ? Units.EmuToCm(ml) : (double?)null,
            MarginRight = tcPr?.RightMargin is { } mr ? Units.EmuToCm(mr) : (double?)null,
            MarginTop = tcPr?.TopMargin is { } mt ? Units.EmuToCm(mt) : (double?)null,
            MarginBottom = tcPr?.BottomMargin is { } mb ? Units.EmuToCm(mb) : (double?)null,
            TextDirection = TextDirectionToken(tcPr?.Vertical?.Value),
        };
    }

    private static string? VAlignToken(A.TextAnchoringTypeValues? anchor)
    {
        if (anchor is null)
        {
            return null;
        }

        if (anchor == A.TextAnchoringTypeValues.Top)
        {
            return "top";
        }

        if (anchor == A.TextAnchoringTypeValues.Center)
        {
            return "middle";
        }

        return anchor == A.TextAnchoringTypeValues.Bottom ? "bottom" : null;
    }

    private static string? TextDirectionToken(A.TextVerticalValues? vert)
    {
        if (vert is null)
        {
            return null;
        }

        if (vert == A.TextVerticalValues.Horizontal)
        {
            return "horizontal";
        }

        if (vert == A.TextVerticalValues.Vertical)
        {
            return "vertical";
        }

        if (vert == A.TextVerticalValues.Vertical270)
        {
            return "vertical270";
        }

        return null;
    }

    /// <summary>The path of the merge origin covering the 1-based (row, col), when resolvable.</summary>
    private static string? MergeOriginPath(A.Table table, string tablePath, int row, int col)
    {
        var rows = table.Elements<A.TableRow>().ToList();
        for (var r = row; r >= 1; r--)
        {
            var cells = rows[r - 1].Elements<A.TableCell>().ToList();
            for (var c = col; c >= 1; c--)
            {
                if (c > cells.Count)
                {
                    continue;
                }

                var candidate = cells[c - 1];
                if (IsCovered(candidate))
                {
                    continue;
                }

                var spansRight = c + (candidate.GridSpan?.Value ?? 1) - 1 >= col;
                var spansDown = r + (candidate.RowSpan?.Value ?? 1) - 1 >= row;
                return spansRight && spansDown ? Units.Inv($"{tablePath}/tr[{r}]/tc[{c}]") : null;
            }
        }

        return null;
    }

    /// <summary>Solid RRGGBB fill of a cell, when set explicitly.</summary>
    public static string? CellFillHex(A.TableCell cell) => cell.TableCellProperties?
        .GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value?.ToUpperInvariant();

    private static bool AsBool(string key, JsonNode? node)
    {
        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var flag))
            {
                return flag;
            }

            if (value.TryGetValue<string>(out var text) && bool.TryParse(text, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Property '{key}' is not a boolean: {node?.ToJsonString() ?? "null"}",
            "Use true or false.");
    }

    private static AiofficeException CorruptTable(string detail) => new(
        ErrorCodes.FormatCorrupt,
        $"The table is malformed: {detail}.",
        "Re-export the file from PowerPoint/Keynote, or restore a snapshot.");
}
