using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// The v1.8 <c>gradientFill</c> cell/range prop (additive). ClosedXML 0.105 has no
/// gradient-fill model — its <c>IXLFill</c> is a single pattern color — so a
/// gradient is authored raw on the saved bytes: a <c>&lt;fill&gt;&lt;gradientFill&gt;</c>
/// is appended to the styles part's <c>&lt;fills&gt;</c> and each target cell's
/// <c>&lt;xf&gt;</c> is rebound to that fill (cloning the cell's existing format so
/// number-format / font / border survive). Linear gradients carry a degree angle;
/// radial / path gradients use a centered <c>&lt;gradientFill type="path"&gt;</c>.
/// Every result stays OpenXmlValidator-clean.
///
/// <c>get</c> on the cell reports the gradient back via <see cref="Describe"/>,
/// read raw from the styles part (ClosedXML cannot surface it through its style
/// API). The cell is also given a solid fill of the first stop's color in the
/// ClosedXML model, so it materializes as a real <c>&lt;c&gt;</c> the post-save
/// pass can rebind, and so a cell that loses the gradient still shows a brand color.
///
/// Round-trip caveat (documented, honest): a gradient survives within the edit that
/// authors it, but a LATER edit that re-saves the workbook through ClosedXML's DOM
/// drops it back to that solid fallback color — ClosedXML rebuilds the cell styles
/// and cannot re-emit a gradient it does not model. Re-apply the gradient in the
/// same batch as any other change to keep it (the gradient definition still rides
/// in <c>&lt;fills&gt;</c>, but the cell is re-pointed to the solid fill). Charts'
/// <c>seriesColors</c> have no such caveat — they live in the chart part ClosedXML
/// preserves byte-identical.
/// </summary>
internal static class ExcelGradientFill
{
    private static readonly IReadOnlyList<string> Types = ["linear", "radial", "path"];

    /// <summary>One color stop: a 6-hex RGB and its 0..1 position along the gradient.</summary>
    internal sealed record Stop(string Color, double Position);

    /// <summary>A parsed gradient definition (the same shape drives create and readback).</summary>
    internal sealed record Definition(string Type, double? Angle, IReadOnlyList<Stop> Stops);

    /// <summary>A queued gradient-fill write for the post-save raw pass.</summary>
    internal sealed record Spec(string SheetName, IReadOnlyList<string> Cells, Definition Gradient);

    /// <summary>
    /// Parses the <c>gradientFill</c> prop value:
    /// <c>{type:"linear"|"radial"|"path", angle?:deg, stops:[{color,pos}]}</c>.
    /// <c>type</c> defaults to linear; <c>angle</c> applies only to linear (degrees,
    /// 0 = left→right). At least two stops are required, each a 6-hex RGB at a
    /// position in 0..1. Throws <c>invalid_args</c> for anything malformed.
    /// </summary>
    public static Definition Parse(JsonNode node, int opIndex)
    {
        if (node is not JsonObject obj)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: 'gradientFill' must be an object like {{\"stops\":[{{\"color\":\"2E5AAC\",\"pos\":0}},{{\"color\":\"E8743B\",\"pos\":1}}]}}.",
                "Pass gradientFill:{type:\"linear\",angle:90,stops:[{color:\"#2E5AAC\",pos:0},{color:\"#E8743B\",pos:1}]}.");
        }

        foreach (var (key, _) in obj)
        {
            if (key is not ("type" or "angle" or "stops"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown gradientFill field '{key}'.",
                    "gradientFill accepts {type, angle, stops}.",
                    candidates: ["type", "angle", "stops"]);
            }
        }

        var type = "linear";
        if (obj.TryGetPropertyValue("type", out var typeNode) && typeNode is not null)
        {
            if (typeNode is not JsonValue tv || tv.GetValueKind() != JsonValueKind.String ||
                tv.GetValue<string>() is not { } typeText || !Types.Contains(typeText, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: gradientFill type must be one of {string.Join(", ", Types)}.",
                    "Use type:\"linear\" (with an angle) or type:\"radial\"/\"path\" (centered).",
                    candidates: Types);
            }

            type = typeText;
        }

        double? angle = null;
        if (obj.TryGetPropertyValue("angle", out var angleNode) && angleNode is not null)
        {
            if (angleNode is not JsonValue av || av.GetValueKind() is not (JsonValueKind.Number) ||
                !av.TryGetValue(out double degrees))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: gradientFill angle must be a number of degrees.",
                    "e.g. angle:0 (left→right), angle:90 (top→bottom).");
            }

            if (type != "linear")
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: gradientFill angle applies only to a linear gradient (type is '{type}').",
                    "Drop angle for radial/path gradients (they are centered), or set type:\"linear\".");
            }

            angle = degrees;
        }

        if (!obj.TryGetPropertyValue("stops", out var stopsNode) || stopsNode is not JsonArray stopsArray)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: gradientFill needs a 'stops' array.",
                "Pass stops:[{color:\"#2E5AAC\",pos:0},{color:\"#E8743B\",pos:1}].");
        }

        var stops = new List<Stop>(stopsArray.Count);
        foreach (var entry in stopsArray)
        {
            stops.Add(ParseStop(entry, opIndex));
        }

        if (stops.Count < 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: gradientFill needs at least two stops.",
                "A gradient blends between stops; list a start and an end, e.g. pos 0 and pos 1.");
        }

        return new Definition(type, angle, stops);
    }

    private static Stop ParseStop(JsonNode? entry, int opIndex)
    {
        if (entry is not JsonObject stop)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: each gradientFill stop must be an object {{color, pos}}.",
                "e.g. {color:\"#2E5AAC\",pos:0}.");
        }

        foreach (var (key, _) in stop)
        {
            if (key is not ("color" or "pos"))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown gradientFill stop field '{key}'.",
                    "A stop accepts {color, pos}.",
                    candidates: ["color", "pos"]);
            }
        }

        if (!stop.TryGetPropertyValue("color", out var colorNode) || colorNode is not JsonValue cv ||
            cv.GetValueKind() != JsonValueKind.String || cv.GetValue<string>() is not { } rawColor)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a gradientFill stop needs a 6-hex 'color'.",
                "e.g. color:\"#2E5AAC\" or color:\"E8743B\".");
        }

        if (!stop.TryGetPropertyValue("pos", out var posNode) || posNode is not JsonValue pv ||
            pv.GetValueKind() != JsonValueKind.Number || !pv.TryGetValue(out double pos))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: a gradientFill stop needs a numeric 'pos' in 0..1.",
                "pos is the fractional position along the gradient: 0 = start, 1 = end.");
        }

        if (pos < 0 || pos > 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: gradientFill stop pos {pos.ToString(CultureInfo.InvariantCulture)} is outside 0..1.",
                "Positions run from 0 (start) to 1 (end).");
        }

        return new Stop(NormalizeHex(rawColor, opIndex), pos);
    }

    private static string NormalizeHex(string raw, int opIndex)
    {
        var hex = raw.Trim();
        if (hex.StartsWith('#'))
        {
            hex = hex[1..];
        }

        if (hex.Length != 6 || !hex.All(Uri.IsHexDigit))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{raw}' is not a 6-hex RGB color.",
                "Use six hex digits like 2E5AAC or #2E5AAC.");
        }

        return hex.ToUpperInvariant();
    }

    // ----- apply (post-save raw pass) ----------------------------------------

    /// <summary>
    /// Authors every queued gradient fill on the saved file. For each spec the
    /// gradient fill is appended once to the styles part; each target cell's
    /// <c>&lt;xf&gt;</c> is cloned with the new fillId and the cell rebound to it,
    /// so the cell keeps its number format / font / border. Identical gradients and
    /// identical resulting <c>&lt;xf&gt;</c>s are deduplicated.
    /// </summary>
    public static void ApplyAfterSave(string file, IReadOnlyList<Spec> specs)
    {
        if (specs.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart
            ?? throw new AiofficeException(
                ErrorCodes.InternalError,
                "The workbook has no workbook part to attach the gradient fill to.",
                "Retry the edit; if it recurs, restore a snapshot.");

        var stylesPart = workbookPart.WorkbookStylesPart
            ?? throw new AiofficeException(
                ErrorCodes.InternalError,
                "The workbook has no styles part to attach the gradient fill to.",
                "Retry the edit; if it recurs, restore a snapshot.");

        var stylesheet = stylesPart.Stylesheet
            ?? throw new AiofficeException(
                ErrorCodes.InternalError,
                "The workbook styles part has no stylesheet.",
                "Retry the edit; if it recurs, restore a snapshot.");

        var fills = stylesheet.Fills ??= new S.Fills();
        var cellFormats = stylesheet.CellFormats
            ?? throw new AiofficeException(
                ErrorCodes.InternalError,
                "The workbook styles part has no cellXfs.",
                "Retry the edit; if it recurs, restore a snapshot.");

        // Dedupe identical gradients (same definition → one fill) and identical
        // rebound xfs (same base style + same fill → one xf) across the whole batch.
        var fillCache = new Dictionary<string, uint>(StringComparer.Ordinal);
        var xfCache = new Dictionary<string, uint>(StringComparer.Ordinal);

        foreach (var spec in specs)
        {
            var worksheet = WorksheetFor(workbookPart, spec.SheetName);
            var sheetData = worksheet.GetFirstChild<S.SheetData>();
            if (sheetData is null)
            {
                continue;
            }

            var fillKey = FillKey(spec.Gradient);
            if (!fillCache.TryGetValue(fillKey, out var fillId))
            {
                fills.Append(BuildFill(spec.Gradient));
                fills.Count = (uint)fills.Count();
                fillId = (uint)(fills.Count() - 1);
                fillCache[fillKey] = fillId;
            }

            foreach (var address in spec.Cells)
            {
                var cell = FindCell(sheetData, address);
                if (cell is null)
                {
                    continue;
                }

                var baseStyleIndex = cell.StyleIndex?.Value ?? 0u;
                var xfKey = string.Create(CultureInfo.InvariantCulture, $"{baseStyleIndex}:{fillId}");
                if (!xfCache.TryGetValue(xfKey, out var styleIndex))
                {
                    var baseFormat = cellFormats.Elements<S.CellFormat>().ElementAtOrDefault((int)baseStyleIndex);
                    var newFormat = baseFormat is null
                        ? new S.CellFormat { FillId = fillId, ApplyFill = true }
                        : (S.CellFormat)baseFormat.CloneNode(true);
                    newFormat.FillId = fillId;
                    newFormat.ApplyFill = true;

                    cellFormats.Append(newFormat);
                    cellFormats.Count = (uint)cellFormats.Count();
                    styleIndex = (uint)(cellFormats.Count() - 1);
                    xfCache[xfKey] = styleIndex;
                }

                cell.StyleIndex = styleIndex;
            }
        }

        stylesheet.Save();
        SaveWorksheets(workbookPart);
    }

    private static void SaveWorksheets(WorkbookPart workbookPart)
    {
        foreach (var part in workbookPart.WorksheetParts)
        {
            part.Worksheet?.Save();
        }
    }

    /// <summary>A stable identity for a gradient definition (drives fill dedupe).</summary>
    private static string FillKey(Definition gradient)
    {
        var stops = string.Join(",", gradient.Stops.Select(s =>
            string.Create(CultureInfo.InvariantCulture, $"{s.Position:R}:{s.Color}")));
        var angle = gradient.Angle?.ToString("R", CultureInfo.InvariantCulture) ?? "-";
        return string.Create(CultureInfo.InvariantCulture, $"{gradient.Type}|{angle}|{stops}");
    }

    /// <summary>Builds the <c>&lt;fill&gt;&lt;gradientFill&gt;</c> for a definition.</summary>
    private static S.Fill BuildFill(Definition gradient)
    {
        var gradientFill = new S.GradientFill();
        if (gradient.Type == "linear")
        {
            // OOXML degree → unit-square fractions. Linear gradients use Degree on
            // the element; the default 0 runs left→right.
            gradientFill.Type = S.GradientValues.Linear;
            gradientFill.Degree = gradient.Angle ?? 0d;
        }
        else
        {
            // radial / path: a centered "path" gradient (the closest OOXML shape).
            gradientFill.Type = S.GradientValues.Path;
            gradientFill.Left = 0.5d;
            gradientFill.Right = 0.5d;
            gradientFill.Top = 0.5d;
            gradientFill.Bottom = 0.5d;
        }

        foreach (var stop in gradient.Stops)
        {
            gradientFill.Append(new S.GradientStop(new S.Color { Rgb = "FF" + stop.Color })
            {
                Position = stop.Position,
            });
        }

        return new S.Fill(gradientFill);
    }

    private static S.Cell? FindCell(S.SheetData sheetData, string address)
    {
        return sheetData.Descendants<S.Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase));
    }

    private static S.Worksheet WorksheetFor(WorkbookPart workbookPart, string sheetName)
    {
        var sheet = workbookPart.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is { } relationshipId &&
            workbookPart.GetPartById(relationshipId) is WorksheetPart { Worksheet: { } worksheet })
        {
            return worksheet;
        }

        throw new AiofficeException(
            ErrorCodes.InternalError,
            $"Sheet '{sheetName}' vanished before the gradient fill could be written.",
            "Retry the edit; if it recurs, restore a snapshot.");
    }

    // ----- readback (get) -----------------------------------------------------

    /// <summary>
    /// Reads a cell's gradient fill back from the raw styles part (ClosedXML cannot
    /// surface it). Returns the same {type, angle?, stops} shape the prop accepts,
    /// or null when the cell has no gradient fill.
    /// </summary>
    public static object? Describe(string file, string sheetName, string address)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.WorkbookStylesPart?.Stylesheet is not { } stylesheet)
        {
            return null;
        }

        var sheet = workbookPart.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is not { } relationshipId ||
            workbookPart.GetPartById(relationshipId) is not WorksheetPart { Worksheet: { } worksheet })
        {
            return null;
        }

        var cell = worksheet.Descendants<S.Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, address, StringComparison.OrdinalIgnoreCase));
        if (cell?.StyleIndex?.Value is not { } styleIndex)
        {
            return null;
        }

        var cellFormat = stylesheet.CellFormats?.Elements<S.CellFormat>().ElementAtOrDefault((int)styleIndex);
        if (cellFormat?.FillId?.Value is not { } fillId)
        {
            return null;
        }

        var fill = stylesheet.Fills?.Elements<S.Fill>().ElementAtOrDefault((int)fillId);
        if (fill?.GetFirstChild<S.GradientFill>() is not { } gradientFill)
        {
            return null;
        }

        var type = gradientFill.Type?.Value == S.GradientValues.Path ? "radial" : "linear";
        double? angle = type == "linear" ? gradientFill.Degree?.Value ?? 0d : null;
        var stops = gradientFill.Elements<S.GradientStop>()
            .Select(s => new
            {
                color = StripAlpha(s.GetFirstChild<S.Color>()?.Rgb?.Value),
                pos = s.Position?.Value ?? 0d,
            })
            .ToList();

        return new { type, angle, stops };
    }

    /// <summary>Drops a leading FF alpha byte from an 8-hex ARGB, leaving 6-hex RGB.</summary>
    private static string? StripAlpha(string? argb)
    {
        if (argb is null)
        {
            return null;
        }

        return argb.Length == 8 ? argb[2..].ToUpperInvariant() : argb.ToUpperInvariant();
    }
}
