using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace AIOffice.Excel;

/// <summary>
/// Sparklines (M5, ClosedXML-native). Each <c>add</c> creates one sparkline in
/// its own x14 sparklineGroup (stored in the worksheet's <c>extLst</c>, so the
/// schema validator stays clean). Kinds map to the OOXML <c>ST_SparklineType</c>
/// set: <c>line</c>, <c>column</c> and <c>winLoss</c> (stored as
/// <c>stacked</c>). Addressing is <c>/Sheet1/sparkline[i]</c> — 1-based over
/// every sparkline on the sheet, groups in order; <c>read --view structure</c>
/// lists the groups.
///
/// Measured ClosedXML 0.105 quirk corrected in a post-save pass
/// (<see cref="FixUpAfterSave"/>): saved sparklineGroup elements carry an
/// <c>xr2:uid</c> attribute (2015 revision namespace) that
/// <c>OpenXmlValidator(Office2019)</c> flags as undeclared; the pass strips it.
/// </summary>
internal static partial class ExcelSparklines
{
    /// <summary>The sparkline kinds aioffice can create.</summary>
    public static readonly IReadOnlyList<string> Kinds = ["line", "column", "winLoss"];

    private static readonly IReadOnlyList<string> AddProps = ["dataRange", "kind", "color", "markers"];

    [GeneratedRegex("^#?(?:[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$")]
    private static partial Regex HexColor();

    // ----- add ---------------------------------------------------------------

    /// <summary>Validates and applies an <c>add sparkline</c> op; returns the details entry.</summary>
    public static object Add(XLWorkbook workbook, ExcelTarget target, EditOp op, int opIndex)
    {
        if (target.Kind != ExcelTargetKind.Cell)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add sparkline targets the cell it lives in, like /Sheet1/E2; " +
                "the data goes in props.dataRange.",
                "Use {op:add, type:sparkline, path:/Sheet1/E2, props:{dataRange:\"A2:D2\", kind:\"line\"}}.");
        }

        var props = op.Props ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add sparkline needs props.",
            "Pass props like {\"dataRange\":\"A2:D2\",\"kind\":\"line\",\"color\":\"376092\",\"markers\":true}.");
        GuardProps(props, opIndex);

        var cell = target.Cell!;
        if (cell.HasSparkline)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: {cell.Address} already has a sparkline.",
                "Remove it first ({op:remove, path:" + SparklinePathOf(target.Sheet, cell) + "}), then add the new one.");
        }

        var dataRangeText = OptionalString(props, "dataRange") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add sparkline needs a 'dataRange'.",
            "Pass the cells to chart, e.g. {\"dataRange\":\"A2:D2\"} (a /Sheet/A2:D2 path also works).");
        var dataRange = ResolveDataRange(workbook, target.Sheet, dataRangeText, opIndex);

        var kind = OptionalString(props, "kind") ?? "line";
        var type = kind switch
        {
            "line" => XLSparklineType.Line,
            "column" => XLSparklineType.Column,
            "winLoss" => XLSparklineType.Stacked, // OOXML spells win/loss "stacked"
            _ => throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"ops[{opIndex}]: sparkline kind '{kind}' is not supported.",
                "Supported kinds: " + string.Join(", ", Kinds) + ".",
                candidates: Kinds),
        };

        var group = target.Sheet.SparklineGroups.Add(cell, dataRange);
        group.SetType(type);

        if (OptionalString(props, "color") is { } colorText)
        {
            if (!HexColor().IsMatch(colorText))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: '{colorText}' is not a usable color.",
                    "Pass RGB hex like 376092 (a leading # and an AARRGGBB alpha form are also accepted).");
            }

            group.Style.SeriesColor = XLColor.FromHtml(colorText.StartsWith('#') ? colorText : "#" + colorText);
        }

        if (props.TryGetPropertyValue("markers", out var markersNode) && markersNode is not null &&
            markersNode.GetValue<bool>())
        {
            if (type != XLSparklineType.Line)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: markers only apply to line sparklines.",
                    "Drop the markers prop, or switch kind to line.");
            }

            group.SetShowMarkers(XLSparklineMarkers.Markers);
        }

        return new
        {
            op = "add",
            type = "sparkline",
            path = SparklinePathOf(target.Sheet, cell),
            sparklineKind = kind,
            cell = cell.Address.ToString(),
            dataRange = dataRange.RangeAddress.ToString(),
        };
    }

    /// <summary>A dataRange is a path (<c>/Sheet1/A2:D2</c>) or a bare range on the sparkline's own sheet.</summary>
    private static IXLRange ResolveDataRange(XLWorkbook workbook, IXLWorksheet sheet, string text, int opIndex)
    {
        var pathText = text.StartsWith('/') ? text : ExcelPaths.SheetPath(sheet) + "/" + text.ToUpperInvariant();
        var target = ExcelPaths.Resolve(workbook, pathText);
        return target.Kind switch
        {
            ExcelTargetKind.Range => target.Range!,
            ExcelTargetKind.Cell => target.Cell!.AsRange(),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'dataRange' must address a cell or range, not '{text}'.",
                "Pass e.g. {\"dataRange\":\"A2:D2\"} or a full path like /Data/A2:D2."),
        };
    }

    // ----- find / describe ------------------------------------------------------

    /// <summary>All sparklines on a sheet in canonical order (groups in order, then group members).</summary>
    public static List<IXLSparkline> AllOn(IXLWorksheet sheet) =>
        [.. sheet.SparklineGroups.SelectMany(group => group)];

    /// <summary>Finds <c>sparkline[i]</c> on a sheet or throws <c>invalid_path</c> with real candidates.</summary>
    public static IXLSparkline Find(ExcelTarget target)
    {
        var sparklines = AllOn(target.Sheet);
        var index = target.SparklineIndex!.Value;
        if (index > sparklines.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No sparkline[{index}] on sheet '{target.Sheet.Name}' ({sparklines.Count} sparkline(s) exist).",
                sparklines.Count > 0
                    ? "Sparkline indices are 1-based per sheet; run 'aioffice read --view structure' to list them."
                    : "This sheet has no sparklines; add one with {op:add, type:sparkline, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + "/E2, props:{dataRange:\"A2:D2\"}}.",
                candidates: sparklines.Count > 0
                    ? [.. Enumerable.Range(1, sparklines.Count).Select(i => ExcelPaths.SparklinePath(target.Sheet, i))]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return sparklines[index - 1];
    }

    /// <summary>One sparkline as agents see it (get).</summary>
    public static object Describe(IXLWorksheet sheet, IXLSparkline sparkline, int index)
    {
        var group = sparkline.SparklineGroup;
        return new
        {
            path = ExcelPaths.SparklinePath(sheet, index),
            kind = "sparkline",
            sheet = sheet.Name,
            cell = sparkline.Location.Address.ToString(),
            dataRange = sparkline.SourceData?.RangeAddress.ToString(),
            sparklineKind = KindName(group.Type),
            color = ExcelConditionalFormats.HexOf(group.Style.SeriesColor),
            markers = group.ShowMarkers.HasFlag(XLSparklineMarkers.Markers) ? true : (bool?)null,
        };
    }

    /// <summary>The sheet's sparkline groups, for read --view structure (member paths use the flat index).</summary>
    public static List<object> ListGroups(IXLWorksheet sheet)
    {
        var flatIndex = new Dictionary<IXLSparkline, int>();
        var i = 1;
        foreach (var sparkline in AllOn(sheet))
        {
            flatIndex[sparkline] = i++;
        }

        return [.. sheet.SparklineGroups.Select(group => (object)new
        {
            kind = KindName(group.Type),
            color = ExcelConditionalFormats.HexOf(group.Style.SeriesColor),
            markers = group.ShowMarkers.HasFlag(XLSparklineMarkers.Markers) ? true : (bool?)null,
            sparklines = group
                .Select(s => new
                {
                    path = ExcelPaths.SparklinePath(sheet, flatIndex[s]),
                    cell = s.Location.Address.ToString(),
                    dataRange = s.SourceData?.RangeAddress.ToString(),
                })
                .ToList(),
        })];
    }

    /// <summary>Removes one sparkline; its group goes too when it was the last member.</summary>
    public static object Remove(ExcelTarget target)
    {
        var sparkline = Find(target);
        var path = ExcelPaths.SparklinePath(target.Sheet, target.SparklineIndex!.Value);
        var cell = sparkline.Location.Address.ToString();
        var group = sparkline.SparklineGroup;
        group.Remove(sparkline);
        if (!group.Any())
        {
            target.Sheet.SparklineGroups.Remove(group);
        }

        return new { op = "remove", path, removed = "sparkline", cell };
    }

    private static string KindName(XLSparklineType type) => type switch
    {
        XLSparklineType.Column => "column",
        XLSparklineType.Stacked => "winLoss",
        _ => "line",
    };

    private static string SparklinePathOf(IXLWorksheet sheet, IXLCell cell)
    {
        var index = 0;
        foreach (var sparkline in AllOn(sheet))
        {
            index++;
            if (sparkline.Location.Address.RowNumber == cell.Address.RowNumber &&
                sparkline.Location.Address.ColumnNumber == cell.Address.ColumnNumber)
            {
                return ExcelPaths.SparklinePath(sheet, index);
            }
        }

        return ExcelPaths.SparklinePath(sheet, index + 1); // about to be added at the end
    }

    // ----- post-save fix-up -----------------------------------------------------

    private const string Revision2015Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";

    /// <summary>
    /// Strips the <c>xr2:uid</c> attribute ClosedXML stamps on saved
    /// sparklineGroup elements (the Office2019 schema validator flags it as
    /// undeclared). Cheap scan; only writes parts that actually change.
    /// </summary>
    public static void FixUpAfterSave(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return;
        }

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            if (worksheetPart.Worksheet is not { } worksheet)
            {
                continue;
            }

            var dirty = false;
            foreach (var group in worksheet.Descendants<X14.SparklineGroup>())
            {
                if (group.ExtendedAttributes.Any(a => a.NamespaceUri == Revision2015Namespace))
                {
                    group.RemoveAttribute("uid", Revision2015Namespace);
                    dirty = true;
                }
            }

            if (dirty)
            {
                worksheet.Save();
            }
        }
    }

    private static void GuardProps(JsonObject props, int opIndex)
    {
        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown sparkline prop '{key}'.",
                    "Supported sparkline props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }
    }

    private static string? OptionalString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == System.Text.Json.JsonValueKind.String
            ? value.GetValue<string>()
            : null;
}
