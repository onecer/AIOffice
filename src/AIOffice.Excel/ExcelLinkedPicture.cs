using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

namespace AIOffice.Excel;

/// <summary>
/// (1.7) The linked picture — Excel's "camera tool": a picture that mirrors a cell
/// range. A true live link (a <c>xdr:pic</c> whose blip fill carries the range
/// formula) is validator-fragile and breaks round-trip, so — per the contract's
/// honest-fallback clause — aioffice renders a STATIC SNAPSHOT of the range to a
/// PNG at edit time and embeds it as a normal picture, anchored where asked. The
/// caller raises a <c>linked_picture_static</c> warning so the agent knows it is a
/// snapshot of the values as of this edit, not a live mirror.
///
/// The snapshot is a real picture (it opens in Excel and survives round-trips), and
/// the source range / anchor / source sheet are recorded in a workbook custom
/// property registry (<see cref="RegistryProperty"/>, a JSON map keyed by picture
/// name, which ClosedXML persists). That registry is what distinguishes a linked
/// picture from a plain image: <c>/Sheet1/linkedPicture[i]</c> get/remove read it.
/// Addressing is 1-based per sheet, ordered by picture name.
/// </summary>
internal static partial class ExcelLinkedPicture
{
    /// <summary>The custom-property key the linked-picture registry is stored under.</summary>
    public const string RegistryProperty = "_aioffice_linkedPictures";

    private static readonly IReadOnlyList<string> AddProps = ["sourceRange", "anchor", "sheet", "name"];

    [GeneratedRegex("^[A-Z]{1,3}[0-9]{1,7}$")]
    private static partial Regex CellPattern();

    [GeneratedRegex("^[A-Z]{1,3}[0-9]{1,7}:[A-Z]{1,3}[0-9]{1,7}$")]
    private static partial Regex RangePattern();

    /// <summary>One linked-picture registry entry.</summary>
    internal sealed record Entry(string Name, string SourceSheet, string SourceRange, string Anchor);

    // ----- add ---------------------------------------------------------------

    /// <summary>
    /// Validates and applies an <c>add linkedPicture</c> op: renders the source
    /// range to a snapshot PNG and embeds it at the anchor on the target sheet,
    /// recording the source/anchor in the registry. Returns the details entry and a
    /// warning noting the snapshot is static. The path names the sheet the picture
    /// LANDS on; props.sheet (optional) names the sheet the source range lives on
    /// (defaults to the target sheet).
    /// </summary>
    public static (object Details, Warning Warning) Add(
        XLWorkbook workbook, ExcelTarget target, EditOp op, int opIndex)
    {
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add linkedPicture targets the sheet the picture lands on, like /Sheet1; " +
                "the mirrored range goes in props.sourceRange.",
                "Use {op:add, type:linkedPicture, path:/Sheet1, props:{sourceRange:\"A1:D10\", anchor:\"G2\"}}.");
        }

        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add linkedPicture needs props.",
                "Pass props like {\"sourceRange\":\"A1:D10\",\"anchor\":\"G2\"}.");
        }

        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown linkedPicture prop '{key}'.",
                    "Supported linkedPicture props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }

        var sourceRangeText = OptionalString(props, "sourceRange") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add linkedPicture needs the 'sourceRange' prop (the range to mirror).",
            "Pass e.g. {\"sourceRange\":\"A1:D10\",\"anchor\":\"G2\"}.");
        var sourceRange = sourceRangeText.ToUpperInvariant();
        if (!RangePattern().IsMatch(sourceRange) && !CellPattern().IsMatch(sourceRange))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{sourceRangeText}' is not a usable sourceRange.",
                "sourceRange is a plain range on the source sheet, e.g. A1:D10 (no sheet prefix).");
        }

        var anchorText = OptionalString(props, "anchor") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add linkedPicture needs the 'anchor' prop (the top-left cell it hangs from).",
            "Pass e.g. {\"anchor\":\"G2\"}.");
        var anchor = anchorText.ToUpperInvariant();
        if (!CellPattern().IsMatch(anchor))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{anchorText}' is not a usable anchor.",
                "anchor is the top-left cell the picture hangs from, e.g. G2.");
        }

        // Resolve the source sheet (defaults to the target sheet).
        var sourceSheet = target.Sheet;
        if (OptionalString(props, "sheet") is { } sheetName)
        {
            if (!workbook.TryGetWorksheet(sheetName, out sourceSheet))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"ops[{opIndex}]: no source sheet named '{sheetName}' for the linked picture.",
                    "props.sheet names the sheet the sourceRange lives on; omit it to mirror a range on the target sheet.",
                    candidates: ExcelPaths.SheetCandidates(workbook, sheetName));
            }
        }

        var grid = ReadGrid(sourceSheet, sourceRange);
        var png = ExcelSnapshotPng.Render(grid);

        IXLPicture picture;
        try
        {
            using var stream = new MemoryStream(png);
            picture = OptionalString(props, "name") is { } explicitName
                ? target.Sheet.AddPicture(stream, XLPictureFormat.Png, explicitName)
                : target.Sheet.AddPicture(stream, XLPictureFormat.Png);
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: could not place the linked picture: {exception.Message}",
                "Pick a picture name that is not already used on the sheet.",
                innerException: exception);
        }

        picture.MoveTo(target.Sheet.Cell(anchor));

        var qualifiedSource = string.Equals(sourceSheet.Name, target.Sheet.Name, StringComparison.OrdinalIgnoreCase)
            ? sourceRange
            : sourceSheet.Name + "!" + sourceRange;

        // Register the metadata under the picture name (custom property; ClosedXML
        // persists it). The raw post-save image fixups never touch custom properties.
        var registry = LoadRegistry(workbook);
        registry[picture.Name] = new Entry(picture.Name, sourceSheet.Name, sourceRange, anchor);
        StoreRegistry(workbook, registry);

        var index = LinkedIndexOf(target.Sheet, registry, picture.Name);
        var details = new
        {
            op = "add",
            type = "linkedPicture",
            path = ExcelPaths.LinkedPicturePath(target.Sheet, index),
            name = picture.Name,
            sourceRange = qualifiedSource,
            anchor,
        };
        var warning = new Warning(
            WarningCodes.LinkedPictureStatic,
            $"linkedPicture on {ExcelPaths.SheetPath(target.Sheet)} is a static snapshot of {qualifiedSource} as of " +
            "this edit; it does not live-update. Re-add it to refresh the snapshot.");
        return (details, warning);
    }

    /// <summary>Reads a range's displayed cell text into a row-major grid for the snapshot.</summary>
    private static List<IReadOnlyList<string>> ReadGrid(IXLWorksheet sheet, string range)
    {
        var xlRange = sheet.Range(range);
        var address = xlRange.RangeAddress;
        var rows = address.LastAddress.RowNumber - address.FirstAddress.RowNumber + 1;
        var columns = address.LastAddress.ColumnNumber - address.FirstAddress.ColumnNumber + 1;

        var grid = new List<IReadOnlyList<string>>(rows);
        for (var r = 0; r < rows; r++)
        {
            var row = new List<string>(columns);
            for (var c = 0; c < columns; c++)
            {
                var cell = sheet.Cell(address.FirstAddress.RowNumber + r, address.FirstAddress.ColumnNumber + c);
                row.Add(ExcelValues.SafeFormatted(cell) ?? string.Empty);
            }

            grid.Add(row);
        }

        return grid;
    }

    // ----- get / remove ------------------------------------------------------

    /// <summary>Describes the linked picture a resolved <c>linkedPicture[i]</c> path addresses.</summary>
    public static object Describe(XLWorkbook workbook, ExcelTarget target)
    {
        var (picture, entry, index) = Resolve(workbook, target);
        var qualifiedSource = string.Equals(entry.SourceSheet, target.Sheet.Name, StringComparison.OrdinalIgnoreCase)
            ? entry.SourceRange
            : entry.SourceSheet + "!" + entry.SourceRange;
        return new
        {
            path = ExcelPaths.LinkedPicturePath(target.Sheet, index),
            kind = "linkedPicture",
            sheet = target.Sheet.Name,
            name = picture.Name,
            sourceRange = qualifiedSource,
            anchor = picture.TopLeftCell.Address.ToString(),
            widthPx = picture.Width,
            heightPx = picture.Height,
            note = "static snapshot of the source range as of the last edit (not a live link)",
        };
    }

    /// <summary>
    /// Removes the linked picture a resolved path addresses: drops the registry
    /// entry now and queues the raw anchor/part deletion for the post-save pass (the
    /// same discipline plain images use — ClosedXML's own picture deletion corrupts
    /// the drawing relationships).
    /// </summary>
    public static (object Details, string Sheet, string Name) Remove(XLWorkbook workbook, ExcelTarget target)
    {
        var (picture, _, index) = Resolve(workbook, target);
        var name = picture.Name;

        var registry = LoadRegistry(workbook);
        registry.Remove(name);
        StoreRegistry(workbook, registry);

        var details = new
        {
            op = "remove",
            path = ExcelPaths.LinkedPicturePath(target.Sheet, index),
            removed = "linkedPicture",
            name,
        };
        return (details, target.Sheet.Name, name);
    }

    /// <summary>
    /// Resolves a 1-based per-sheet <c>linkedPicture[i]</c> to its picture + entry.
    /// Only pictures registered as linked pictures on the sheet count (plain images
    /// are addressed via <c>image[i]</c>).
    /// </summary>
    private static (IXLPicture Picture, Entry Entry, int Index) Resolve(XLWorkbook workbook, ExcelTarget target)
    {
        var registry = LoadRegistry(workbook);
        var linked = LinkedOnSheet(target.Sheet, registry);
        var wanted = target.LinkedPictureIndex!.Value;
        if (wanted < 1 || wanted > linked.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No linkedPicture[{wanted}] on sheet '{target.Sheet.Name}' ({linked.Count} linked picture(s) exist).",
                linked.Count > 0
                    ? "Linked-picture indices are 1-based per sheet; pick one of the candidates."
                    : "This sheet has no linked pictures; add one with {op:add, type:linkedPicture, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + ", props:{sourceRange:\"A1:D10\", anchor:\"G2\"}}.",
                candidates: linked.Count > 0
                    ? [.. Enumerable.Range(1, linked.Count).Select(i => ExcelPaths.LinkedPicturePath(target.Sheet, i))]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        var (picture, entry) = linked[wanted - 1];
        return (picture, entry, wanted);
    }

    /// <summary>The registered linked pictures on a sheet, ordered by name (stable indices).</summary>
    private static List<(IXLPicture Picture, Entry Entry)> LinkedOnSheet(
        IXLWorksheet sheet, IReadOnlyDictionary<string, Entry> registry)
    {
        var result = new List<(IXLPicture, Entry)>();
        foreach (var picture in sheet.Pictures.OrderBy(p => p.Name, StringComparer.Ordinal))
        {
            if (registry.TryGetValue(picture.Name, out var entry))
            {
                result.Add((picture, entry));
            }
        }

        return result;
    }

    private static int LinkedIndexOf(IXLWorksheet sheet, IReadOnlyDictionary<string, Entry> registry, string name)
    {
        var linked = LinkedOnSheet(sheet, registry);
        for (var i = 0; i < linked.Count; i++)
        {
            if (string.Equals(linked[i].Entry.Name, name, StringComparison.Ordinal))
            {
                return i + 1;
            }
        }

        return linked.Count;
    }

    /// <summary>The count of linked pictures on a sheet (for the sheet-get hint).</summary>
    public static int CountOnSheet(XLWorkbook workbook, IXLWorksheet sheet) =>
        LinkedOnSheet(sheet, LoadRegistry(workbook)).Count;

    // ----- registry (custom-property JSON map) -------------------------------

    private static Dictionary<string, Entry> LoadRegistry(XLWorkbook workbook)
    {
        var result = new Dictionary<string, Entry>(StringComparer.Ordinal);
        var property = workbook.CustomProperties
            .FirstOrDefault(p => string.Equals(p.Name, RegistryProperty, StringComparison.Ordinal));
        if (property?.Value?.ToString() is not { Length: > 0 } raw)
        {
            return result;
        }

        JsonNode? parsed;
        try
        {
            parsed = JsonNode.Parse(raw);
        }
        catch (JsonException)
        {
            return result; // a corrupt registry reads as empty, never crashes
        }

        if (parsed is not JsonArray array)
        {
            return result;
        }

        foreach (var node in array)
        {
            if (node is not JsonObject obj ||
                GetString(obj, "name") is not { } name ||
                GetString(obj, "sourceRange") is not { } source)
            {
                continue;
            }

            result[name] = new Entry(
                name,
                GetString(obj, "sourceSheet") ?? string.Empty,
                source,
                GetString(obj, "anchor") ?? string.Empty);
        }

        return result;
    }

    private static void StoreRegistry(XLWorkbook workbook, IReadOnlyDictionary<string, Entry> registry)
    {
        if (workbook.CustomProperties.Any(p => string.Equals(p.Name, RegistryProperty, StringComparison.Ordinal)))
        {
            workbook.CustomProperties.Delete(RegistryProperty);
        }

        if (registry.Count == 0)
        {
            return; // an empty registry leaves no property behind
        }

        var array = new JsonArray();
        foreach (var entry in registry.Values.OrderBy(e => e.Name, StringComparer.Ordinal))
        {
            array.Add(new JsonObject
            {
                ["name"] = entry.Name,
                ["sourceSheet"] = entry.SourceSheet,
                ["sourceRange"] = entry.SourceRange,
                ["anchor"] = entry.Anchor,
            });
        }

        workbook.CustomProperties.Add(RegistryProperty, array.ToJsonString(JsonDefaults.Options));
    }

    private static string? GetString(JsonObject obj, string key) =>
        obj.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static string? OptionalString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String && value.GetValue<string>().Length > 0
            ? value.GetValue<string>()
            : null;
}
