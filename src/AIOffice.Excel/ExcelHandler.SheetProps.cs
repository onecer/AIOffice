using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// The M3 sheet-level set props, all ClosedXML-native:
/// <list type="bullet">
/// <item>freeze panes — <c>{op:set, path:/Sheet1, props:{freezeRows:1, freezeCols:2}}</c>, 0 clears an axis;</item>
/// <item>autofilter — <c>{op:set, path:/Sheet1/A1:D20, props:{autoFilter:true}}</c>, false clears,
/// one filter per sheet (a second range is invalid_args);</item>
/// <item>page setup — <c>{op:set, path:/Sheet1, props:{orientation:"landscape", paperSize:"A4",
/// fitToWidth:1, printArea:"A1:F40"}}</c>, get reflects everything.</item>
/// </list>
/// </summary>
public sealed partial class ExcelHandler
{
    private const int MaxSheetRows = 1048576;
    private const int MaxSheetColumns = 16384;

    private static readonly IReadOnlyList<string> Orientations = ["portrait", "landscape"];

    /// <summary>Wire name → ClosedXML paper size. Reflection back uses the same table.</summary>
    private static readonly IReadOnlyDictionary<string, XLPaperSize> PaperSizes =
        new Dictionary<string, XLPaperSize>(StringComparer.OrdinalIgnoreCase)
        {
            ["A3"] = XLPaperSize.A3Paper,
            ["A4"] = XLPaperSize.A4Paper,
            ["A5"] = XLPaperSize.A5Paper,
            ["Letter"] = XLPaperSize.LetterPaper,
            ["Legal"] = XLPaperSize.LegalPaper,
            ["Tabloid"] = XLPaperSize.TabloidPaper,
        };

    [GeneratedRegex("^[A-Z]{1,3}[0-9]{1,7}(:[A-Z]{1,3}[0-9]{1,7})?$")]
    private static partial Regex PrintAreaPattern();

    /// <summary>A row-band like "1:1" or "1:3" — the rows repeated at the top of every printed page.</summary>
    [GeneratedRegex("^([0-9]{1,7}):([0-9]{1,7})$")]
    private static partial Regex RowBandPattern();

    /// <summary>A column-band like "A:A" or "A:C" — the columns repeated at the left of every printed page.</summary>
    [GeneratedRegex("^([A-Z]{1,3}):([A-Z]{1,3})$")]
    private static partial Regex ColumnBandPattern();

    // ----- freeze panes -------------------------------------------------------

    private static void ApplyFreeze(ExcelTarget target, EditOp op, JsonNode node, string prop, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, prop, index);
        var value = IntPropValue(node, prop, index);
        var limit = prop == "freezeRows" ? MaxSheetRows - 1 : MaxSheetColumns - 1;
        if (value < 0 || value > limit)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: {prop} must be between 0 and {limit}; got {value}.",
                $"Pass the number of leading {(prop == "freezeRows" ? "rows" : "columns")} to freeze; 0 clears the freeze.");
        }

        if (prop == "freezeRows")
        {
            target.Sheet.SheetView.SplitRow = value;
        }
        else
        {
            target.Sheet.SheetView.SplitColumn = value;
        }

        applied.Add(prop);
    }

    // ----- tab color (v1.24, additive) ------------------------------------------

    /// <summary>
    /// <c>{tabColor:"4472C4"}</c> paints the worksheet tab (<c>sheetPr/tabColor</c>);
    /// <c>""</c> clears it (<see cref="XLColor.NoColor"/>). Sheet-level only — a
    /// cell/range path is invalid_args. The hex is validated through the same guard
    /// fill/color set ops use, so a bad hex fails atomically with an example.
    /// </summary>
    private static void ApplyTabColor(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "tabColor", index);
        var text = StringPropValue(node, "tabColor", index).Trim();
        if (text.Length == 0)
        {
            target.Sheet.TabColor = XLColor.NoColor;
            applied.Add("tabColorCleared");
            return;
        }

        var hex = text.StartsWith('#') ? text[1..] : text;
        target.Sheet.TabColor = ParseColor("#" + hex);
        applied.Add("tabColor");
    }

    /// <summary>
    /// The sheet tab color as an <c>RRGGBB</c> hex for get, or <c>null</c> when the
    /// tab has no color OR carries a theme/indexed color ClosedXML cannot expand to
    /// RGB (unresolvable → omitted rather than mis-reported).
    /// </summary>
    private static string? TabColorHex(XLColor? color)
    {
        if (color is null || !color.HasValue)
        {
            return null;
        }

        try
        {
            var c = color.Color; // resolves RGB + indexed; theme colors throw → null
            return $"{c.R:X2}{c.G:X2}{c.B:X2}";
        }
        catch (Exception)
        {
            return null;
        }
    }

    // ----- autofilter ----------------------------------------------------------

    private static void ApplyAutoFilter(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        if (target.Kind != ExcelTargetKind.Range)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: autoFilter targets a range, not {target.Kind.ToString().ToLowerInvariant()} '{op.Path}'.",
                "Address the data including its header row, e.g. {op:set, path:/Sheet1/A1:D20, props:{autoFilter:true}}.");
        }

        var sheet = target.Sheet;
        var requested = target.Range!.RangeAddress.ToString();

        // (1.12, additive) the criteria form — an object {column, values|criteria} or
        // an array of them — sets a real applied filter that HIDES the non-matching
        // rows. The bool form below is unchanged.
        if (node is JsonObject or JsonArray)
        {
            GuardSingleFilter(sheet, requested, index);
            applied.AddRange(ExcelAutoFilter.Apply(sheet, target.Range!, node, index));
            return;
        }

        var enable = BoolPropValue(node, "autoFilter", index);
        if (!enable)
        {
            sheet.AutoFilter.Clear();
            applied.Add("autoFilterCleared");
            return;
        }

        GuardSingleFilter(sheet, requested, index);
        target.Range!.SetAutoFilter();
        applied.Add("autoFilter");
    }

    /// <summary>Excel allows one autofilter per sheet; a second range is invalid_args with the clear recipe.</summary>
    private static void GuardSingleFilter(IXLWorksheet sheet, string? requested, int index)
    {
        if (sheet.AutoFilter.IsEnabled &&
            sheet.AutoFilter.Range?.RangeAddress.ToString() is { } existing &&
            !string.Equals(existing, requested, StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: sheet '{sheet.Name}' already has an autofilter on {existing}; Excel allows one per sheet.",
                "Clear it first with {op:set, path:" + ExcelPaths.SheetPath(sheet) + "/" + existing +
                ", props:{autoFilter:false}}, then set the new range.");
        }
    }

    // ----- page setup -----------------------------------------------------------

    private static void ApplyOrientation(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "orientation", index);
        var text = StringPropValue(node, "orientation", index);
        if (!Orientations.Contains(text, StringComparer.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: unknown orientation '{text}'.",
                "Use \"portrait\" or \"landscape\".",
                candidates: Orientations);
        }

        target.Sheet.PageSetup.PageOrientation =
            string.Equals(text, "landscape", StringComparison.OrdinalIgnoreCase)
                ? XLPageOrientation.Landscape
                : XLPageOrientation.Portrait;
        applied.Add("orientation");
    }

    private static void ApplyPaperSize(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "paperSize", index);
        var text = StringPropValue(node, "paperSize", index);
        if (!PaperSizes.TryGetValue(text, out var size))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: unknown paperSize '{text}'.",
                "Supported paper sizes: " + string.Join(", ", PaperSizes.Keys.Order(StringComparer.Ordinal)) + ".",
                candidates: [.. PaperSizes.Keys.Order(StringComparer.Ordinal)]);
        }

        target.Sheet.PageSetup.PaperSize = size;
        applied.Add("paperSize");
    }

    private static void ApplyFitToWidth(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "fitToWidth", index);
        var value = IntPropValue(node, "fitToWidth", index);
        if (value < 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: fitToWidth must be 0 or more; got {value}.",
                "Pass the number of pages the sheet width must fit on (height stays automatic); 0 clears the fit.");
        }

        target.Sheet.PageSetup.FitToPages(value, 0); // height 0 = automatic
        applied.Add("fitToWidth");
    }

    private static void ApplyPrintArea(ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "printArea", index);
        var text = StringPropValue(node, "printArea", index);
        var setup = target.Sheet.PageSetup;
        if (text.Length == 0)
        {
            setup.PrintAreas.Clear();
            applied.Add("printAreaCleared");
            return;
        }

        var normalized = text.ToUpperInvariant();
        if (!PrintAreaPattern().IsMatch(normalized))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{text}' is not a usable printArea.",
                "printArea is a plain range on the sheet itself, e.g. A1:F40 (no sheet prefix); \"\" clears it.");
        }

        setup.PrintAreas.Clear();
        setup.PrintAreas.Add(normalized);
        applied.Add("printArea");
    }

    // ----- print completeness (v1.7, additive) -----------------------------------

    /// <summary>
    /// <c>{printTitleRows:"1:1"}</c> repeats those rows at the top of every printed
    /// page (Excel's "Rows to repeat at top"); <c>""</c> clears the repeat.
    /// </summary>
    private static void ApplyPrintTitleRows(
        ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied, PostSaveWork post)
    {
        RequireSheetTarget(target, op, "printTitleRows", index);
        var text = StringPropValue(node, "printTitleRows", index);
        var setup = target.Sheet.PageSetup;
        if (text.Length == 0)
        {
            // ClosedXML 0.105 throws on every clear form, so the band is cleared raw
            // post-save (the _xlnm.Print_Titles defined name's row segment).
            post.PrintTitleClears.Add(new ExcelPrintTitles.ClearSpec(target.Sheet.Name, ClearRows: true, ClearCols: false));
            applied.Add("printTitleRowsCleared");
            return;
        }

        var match = RowBandPattern().Match(text);
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{text}' is not a usable printTitleRows band.",
                "printTitleRows is a row band like \"1:1\" (repeat the header row) or \"1:3\"; \"\" clears it.");
        }

        var first = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
        var last = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
        if (first < 1 || last < first || last > MaxSheetRows)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: printTitleRows band '{text}' must be 1-based with first ≤ last (≤ {MaxSheetRows}).",
                "Use \"1:1\" to repeat the first row, or \"1:3\" for the first three rows.");
        }

        setup.SetRowsToRepeatAtTop(first, last);
        applied.Add("printTitleRows");
    }

    /// <summary>
    /// <c>{printTitleCols:"A:A"}</c> repeats those columns at the left of every
    /// printed page (Excel's "Columns to repeat at left"); <c>""</c> clears it.
    /// </summary>
    private static void ApplyPrintTitleCols(
        ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied, PostSaveWork post)
    {
        RequireSheetTarget(target, op, "printTitleCols", index);
        var text = StringPropValue(node, "printTitleCols", index).ToUpperInvariant();
        var setup = target.Sheet.PageSetup;
        if (text.Length == 0)
        {
            // Cleared raw post-save (ClosedXML 0.105 throws on the clear forms).
            post.PrintTitleClears.Add(new ExcelPrintTitles.ClearSpec(target.Sheet.Name, ClearRows: false, ClearCols: true));
            applied.Add("printTitleColsCleared");
            return;
        }

        var match = ColumnBandPattern().Match(text);
        if (!match.Success)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{text}' is not a usable printTitleCols band.",
                "printTitleCols is a column band like \"A:A\" (repeat the first column) or \"A:C\"; \"\" clears it.");
        }

        var first = new CellRef(match.Groups[1].Value, 1).ColumnNumber;
        var last = new CellRef(match.Groups[2].Value, 1).ColumnNumber;
        if (last < first || last > MaxSheetColumns)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: printTitleCols band '{text}' must have first ≤ last (≤ column {MaxSheetColumns}).",
                "Use \"A:A\" to repeat the first column, or \"A:C\" for the first three columns.");
        }

        setup.SetColumnsToRepeatAtLeft(first, last);
        applied.Add("printTitleCols");
    }

    /// <summary>
    /// <c>{fitToPage:{fitToWidth:1,fitToHeight:0}}</c> or <c>{fitToPage:{scale:80}}</c>.
    /// Width/height are page counts (0 = automatic on that axis); scale is a whole
    /// percent (10–400). The two forms are mutually exclusive — the last one set in
    /// Excel wins, so passing both is invalid_args.
    /// </summary>
    private static void ApplyFitToPage(
        ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "fitToPage", index);
        if (node is not JsonObject spec)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: fitToPage takes an object.",
                "Use {fitToPage:{fitToWidth:1,fitToHeight:0}} (page counts; 0 = automatic) or {fitToPage:{scale:80}} (percent).");
        }

        var hasWidth = spec.TryGetPropertyValue("fitToWidth", out var widthNode) && widthNode is not null;
        var hasHeight = spec.TryGetPropertyValue("fitToHeight", out var heightNode) && heightNode is not null;
        var hasScale = spec.TryGetPropertyValue("scale", out var scaleNode) && scaleNode is not null;

        if (hasScale && (hasWidth || hasHeight))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: fitToPage takes EITHER scale OR fitToWidth/fitToHeight, not both (Excel keeps only one).",
                "Pass {fitToPage:{scale:80}} to scale by percent, or {fitToPage:{fitToWidth:1,fitToHeight:0}} to fit to pages.");
        }

        var setup = target.Sheet.PageSetup;
        if (hasScale)
        {
            var scale = IntPropValue(scaleNode!, "scale", index);
            if (scale < 10 || scale > 400)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: fitToPage scale must be 10–400 percent; got {scale}.",
                    "Excel allows a print scale between 10% and 400%.");
            }

            setup.Scale = scale; // overrides any PagesWide/PagesTall
            applied.Add("fitToPageScale");
            return;
        }

        if (!hasWidth && !hasHeight)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: fitToPage needs fitToWidth and/or fitToHeight (or scale).",
                "Use {fitToPage:{fitToWidth:1,fitToHeight:0}}; 0 leaves that axis automatic.");
        }

        var width = hasWidth ? FitPageCount(widthNode!, "fitToWidth", index) : setup.PagesWide;
        var height = hasHeight ? FitPageCount(heightNode!, "fitToHeight", index) : setup.PagesTall;
        setup.FitToPages(width, height); // overrides Scale; 0 = automatic on that axis
        applied.Add("fitToPage");
    }

    /// <summary>
    /// <c>{fitToHeight:N}</c> — the flat-prop twin of <c>fitToWidth</c> (M3): fit the
    /// sheet height to N pages (width stays automatic). 0 clears the height fit.
    /// </summary>
    private static void ApplyFitToHeight(
        ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "fitToHeight", index);
        var value = FitPageCount(node, "fitToHeight", index);
        var setup = target.Sheet.PageSetup;
        setup.FitToPages(setup.PagesWide, value); // keep any existing width fit
        applied.Add("fitToHeight");
    }

    private static int FitPageCount(JsonNode node, string prop, int index)
    {
        var value = IntPropValue(node, prop, index);
        if (value < 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: {prop} must be 0 or more; got {value}.",
                $"Pass the number of pages to fit on that axis; 0 leaves it automatic.");
        }

        return value;
    }

    /// <summary>
    /// <c>{pageBreaks:{rows:[20,40],cols:["F"]}}</c> — manual page breaks: a row
    /// break before each listed row (a horizontal break above it) and a column
    /// break before each listed column. Empty <c>rows</c>/<c>cols</c> arrays clear
    /// that axis's manual breaks. Accepts <c>rowBreaks</c>/<c>colBreaks</c> aliases.
    /// </summary>
    private static void ApplyPageBreaks(
        ExcelTarget target, EditOp op, JsonNode node, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, "pageBreaks", index);
        if (node is not JsonObject spec)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: pageBreaks takes an object {{rows, cols}}.",
                "Use {pageBreaks:{rows:[20,40],cols:[\"F\"]}}; an empty array clears that axis.");
        }

        var rowsNode = spec.TryGetPropertyValue("rows", out var r) && r is not null ? r
            : spec.TryGetPropertyValue("rowBreaks", out var rb) ? rb : null;
        var colsNode = spec.TryGetPropertyValue("cols", out var c) && c is not null ? c
            : spec.TryGetPropertyValue("colBreaks", out var cb) ? cb : null;
        if (rowsNode is null && colsNode is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: pageBreaks needs rows and/or cols.",
                "Use {pageBreaks:{rows:[20,40],cols:[\"F\"]}}; an empty array clears that axis.");
        }

        var setup = target.Sheet.PageSetup;
        if (rowsNode is not null)
        {
            ApplyRowBreaks(setup, rowsNode, index, applied);
        }

        if (colsNode is not null)
        {
            ApplyColumnBreaks(setup, colsNode, index, applied);
        }
    }

    private static void ApplyRowBreaks(IXLPageSetup setup, JsonNode node, int index, List<string> applied)
    {
        if (node is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: pageBreaks.rows must be an array of row numbers.",
                "Use {pageBreaks:{rows:[20,40]}}; [] clears the manual row breaks.");
        }

        foreach (var existing in setup.RowBreaks.ToList())
        {
            setup.RowBreaks.Remove(existing); // explicit list replaces the manual breaks on this axis
        }

        foreach (var item in array)
        {
            var row = item is null ? 0 : IntPropValue(item, "pageBreaks.rows", index);
            if (row < 1 || row > MaxSheetRows)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: pageBreaks.rows holds {row}, which is not a 1-based row (≤ {MaxSheetRows}).",
                    "Each entry is the row a page break sits above, e.g. {pageBreaks:{rows:[20,40]}}.");
            }

            setup.AddHorizontalPageBreak(row);
        }

        applied.Add(array.Count == 0 ? "rowBreaksCleared" : "rowBreaks");
    }

    private static void ApplyColumnBreaks(IXLPageSetup setup, JsonNode node, int index, List<string> applied)
    {
        if (node is not JsonArray array)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: pageBreaks.cols must be an array of column letters or numbers.",
                "Use {pageBreaks:{cols:[\"F\"]}}; [] clears the manual column breaks.");
        }

        foreach (var existing in setup.ColumnBreaks.ToList())
        {
            setup.ColumnBreaks.Remove(existing);
        }

        foreach (var item in array)
        {
            var column = ColumnBreakNumber(item, index);
            if (column < 1 || column > MaxSheetColumns)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: pageBreaks.cols holds an entry that is not a 1-based column (≤ {MaxSheetColumns}).",
                    "Each entry is the column a page break sits left of, as a letter (\"F\") or number (6).");
            }

            setup.AddVerticalPageBreak(column);
        }

        applied.Add(array.Count == 0 ? "colBreaksCleared" : "colBreaks");
    }

    /// <summary>A column break entry is a letter ("F") or a 1-based number (6).</summary>
    private static int ColumnBreakNumber(JsonNode? node, int index)
    {
        if (node is JsonValue value)
        {
            if (value.GetValueKind() == JsonValueKind.Number && value.TryGetValue<int>(out var number))
            {
                return number;
            }

            if (value.GetValueKind() == JsonValueKind.String)
            {
                var text = value.GetValue<string>().Trim().ToUpperInvariant();
                if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
                {
                    return parsed;
                }

                if (text.Length is >= 1 and <= 3 && text.All(char.IsAsciiLetterUpper))
                {
                    return new CellRef(text, 1).ColumnNumber;
                }
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: a pageBreaks.cols entry must be a column letter (\"F\") or number (6).",
            "Use {pageBreaks:{cols:[\"F\"]}} or {pageBreaks:{cols:[6]}}.");
    }

    /// <summary>One of the boolean page-setup print toggles, all ClosedXML-native.</summary>
    private static void ApplyPrintBool(
        ExcelTarget target, EditOp op, JsonNode node, string prop, int index, List<string> applied)
    {
        RequireSheetTarget(target, op, prop, index);
        var value = BoolPropValue(node, prop, index);
        var setup = target.Sheet.PageSetup;
        switch (prop)
        {
            case "centerHorizontally": setup.CenterHorizontally = value; break;
            case "centerVertically": setup.CenterVertically = value; break;
            case "printGridlines": setup.ShowGridlines = value; break;
            case "printHeadings": setup.ShowRowAndColumnHeadings = value; break;
        }

        applied.Add(prop);
    }

    /// <summary>
    /// <c>{printHeader:{left:"&F",center:"&A",right:"&D"}}</c> (or printFooter) sets
    /// the odd-page header/footer field-code strings. Excel field codes pass through
    /// verbatim: &amp;P page, &amp;N pages, &amp;D date, &amp;T time, &amp;F file,
    /// &amp;A sheet. A section sent as <c>null</c> or <c>""</c> clears that section;
    /// a section the op omits is left untouched. Authored raw post-save (ClosedXML
    /// 0.105 does not persist header/footer text).
    /// </summary>
    private static void ApplyPrintHeaderFooter(
        ExcelTarget target, EditOp op, JsonNode node, bool isHeader, int index, List<string> applied, PostSaveWork post)
    {
        var prop = isHeader ? "printHeader" : "printFooter";
        RequireSheetTarget(target, op, prop, index);
        if (node is not JsonObject spec)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: {prop} takes an object with left/center/right sections.",
                "Use {" + prop + ":{left:\"&F\",center:\"&A\",right:\"&D\"}}; &P page, &N pages, &D date, &T time, &F file, &A sheet.");
        }

        foreach (var (key, _) in spec)
        {
            if (key is not ("left" or "center" or "right"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: {prop} section '{key}' is not left, center or right.",
                    "Header/footer sections are left, center and right; pass any subset.");
            }
        }

        string? Section(string name)
        {
            if (!spec.TryGetPropertyValue(name, out var sectionNode))
            {
                return null; // omitted → leave as-is
            }

            return sectionNode is JsonValue v && v.GetValueKind() == JsonValueKind.String
                ? v.GetValue<string>()
                : sectionNode is null || sectionNode.GetValueKind() == JsonValueKind.Null
                    ? string.Empty // present-but-null → clear
                    : throw new AiofficeException(
                        ErrorCodes.InvalidArgs,
                        $"ops[{index}]: {prop}.{name} must be a string (or null to clear).",
                        "Pass a field-code string like \"Page &P of &N\".");
        }

        post.HeaderFooters.Add(new ExcelPrintHeaderFooter.Spec(
            target.Sheet.Name, isHeader, Section("left"), Section("center"), Section("right")));
        applied.Add(prop);
    }

    // ----- get reflection --------------------------------------------------------

    /// <summary>The page-setup block for sheet get (null members are omitted on the wire).</summary>
    private static object PageSetupInfo(IXLWorksheet sheet, string file)
    {
        var setup = sheet.PageSetup;
        var printAreas = setup.PrintAreas.Cast<IXLRange>()
            .Select(r => RelativeRangeText(r.RangeAddress))
            .ToList();

        // Repeat bands: ClosedXML reports 0 for "none" on each edge.
        var titleRows = setup.FirstRowToRepeatAtTop > 0
            ? string.Create(CultureInfo.InvariantCulture, $"{setup.FirstRowToRepeatAtTop}:{setup.LastRowToRepeatAtTop}")
            : null;
        var titleCols = setup.FirstColumnToRepeatAtLeft > 0
            ? $"{ExcelCharts.ColumnLetters(setup.FirstColumnToRepeatAtLeft)}:" +
              ExcelCharts.ColumnLetters(setup.LastColumnToRepeatAtLeft)
            : null;

        var rowBreaks = setup.RowBreaks.OrderBy(b => b).ToList();
        var colBreaks = setup.ColumnBreaks.OrderBy(b => b).ToList();
        object? pageBreaks = rowBreaks.Count > 0 || colBreaks.Count > 0
            ? new { rows = rowBreaks, cols = colBreaks.Select(ExcelCharts.ColumnLetters).ToList() }
            : null;

        // Header/footer text lives in raw bytes (ClosedXML 0.105 cannot read it back).
        var (printHeader, printFooter) = ExcelPrintHeaderFooter.Read(file, sheet.Name);

        return new
        {
            orientation = setup.PageOrientation.ToString().ToLowerInvariant(),
            paperSize = PaperSizes.FirstOrDefault(p => p.Value == setup.PaperSize).Key ?? setup.PaperSize.ToString(),
            fitToWidth = setup.PagesWide > 0 ? setup.PagesWide : (int?)null,
            fitToHeight = setup.PagesTall > 0 ? setup.PagesTall : (int?)null,
            // Excel's print scale defaults to 100; surface it only when the sheet is
            // NOT in fit-to-page mode and the scale deviates from the default.
            scale = setup.PagesWide == 0 && setup.PagesTall == 0 && setup.Scale != 100 ? setup.Scale : (int?)null,
            printArea = printAreas.Count > 0 ? string.Join(",", printAreas) : null,
            printTitleRows = titleRows,
            printTitleCols = titleCols,
            pageBreaks,
            centerHorizontally = setup.CenterHorizontally ? true : (bool?)null,
            centerVertically = setup.CenterVertically ? true : (bool?)null,
            printGridlines = setup.ShowGridlines ? true : (bool?)null,
            printHeadings = setup.ShowRowAndColumnHeadings ? true : (bool?)null,
            printHeader,
            printFooter,
        };
    }

    /// <summary>Print areas stringify as absolute refs ($A$1:$F$40); the wire uses plain A1:F40.</summary>
    private static string RelativeRangeText(IXLRangeAddress address) => string.Create(
        CultureInfo.InvariantCulture,
        $"{address.FirstAddress.ColumnLetter}{address.FirstAddress.RowNumber}:" +
        $"{address.LastAddress.ColumnLetter}{address.LastAddress.RowNumber}");

    // ----- prop plumbing -----------------------------------------------------------

    private static void RequireSheetTarget(ExcelTarget target, EditOp op, string prop, int index)
    {
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{prop}' targets a sheet path like /Sheet1, not '{op.Path}'.",
                "Use {op:set, path:/Sheet1, props:{" + prop + ":…}}.");
        }
    }

    private static int IntPropValue(JsonNode node, string prop, int index)
    {
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
            $"ops[{index}]: '{prop}' must be a whole number.",
            $"Pass e.g. {{\"{prop}\":1}}.");
    }

    private static bool BoolPropValue(JsonNode node, string prop, int index)
    {
        if (node is JsonValue value)
        {
            switch (value.GetValueKind())
            {
                case JsonValueKind.True:
                    return true;
                case JsonValueKind.False:
                    return false;
                case JsonValueKind.String when bool.TryParse(value.GetValue<string>(), out var parsed):
                    return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: '{prop}' must be true or false.",
            $"Pass e.g. {{\"{prop}\":true}}.");
    }

    private static string StringPropValue(JsonNode node, string prop, int index)
    {
        if (node is JsonValue value && value.GetValueKind() == JsonValueKind.String)
        {
            return value.GetValue<string>();
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: '{prop}' must be a string.",
            $"Pass e.g. {{\"{prop}\":\"A4\"}}.");
    }
}
