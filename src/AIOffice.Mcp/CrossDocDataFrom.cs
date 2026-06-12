using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// The M3 cross-document data bridge, applied at the COMMAND layer (shared by
/// the CLI's edit verb and MCP <c>office_edit</c>) before ops reach a format
/// handler: a pptx <c>add chart</c> op may carry
/// <c>{"dataFrom":"metrics.xlsx!Sheet1/A1:B5"}</c> instead of literal
/// categories/series. The workbook is resolved through the workspace sandbox
/// and read via the registered xlsx handler; the range's first column becomes
/// the categories, every remaining column becomes one series, and the header
/// row carries the series names. The op is rewritten with injected literals,
/// so the pptx handler (and its literal-cache chart writer) is none the wiser.
/// </summary>
public static class CrossDocDataFrom
{
    public const string PropName = "dataFrom";

    /// <summary>Cell cap for the source range read; charts beyond this are not a thing.</summary>
    private const int MaxSourceCells = 10_000;

    private const string SpecShape =
        "dataFrom looks like \"book.xlsx!Sheet1/A1:B5\" (sheet names with spaces: \"book.xlsx!'Q3 Data'/A1:B5\"): " +
        "first column = category labels, remaining columns = one series each, header row = series names.";

    /// <summary>
    /// Rewrites every qualifying op (pptx target, add chart, props.dataFrom)
    /// into its literal categories/series form. Ops without dataFrom pass
    /// through untouched; non-pptx targets are returned as-is.
    /// </summary>
    public static IReadOnlyList<EditOp> Expand(
        IReadOnlyList<EditOp> ops,
        DocumentKind targetKind,
        Workspace workspace,
        HandlerRegistry handlers)
    {
        if (targetKind != DocumentKind.Pptx || !ops.Any(NeedsExpansion))
        {
            return ops;
        }

        var expanded = new List<EditOp>(ops.Count);
        for (var i = 0; i < ops.Count; i++)
        {
            expanded.Add(NeedsExpansion(ops[i]) ? ExpandOne(ops[i], i, workspace, handlers) : ops[i]);
        }

        return expanded;
    }

    private static bool NeedsExpansion(EditOp op) =>
        op.Op == "add" &&
        string.Equals(op.Type, "chart", StringComparison.OrdinalIgnoreCase) &&
        op.Props?.ContainsKey(PropName) == true;

    private static EditOp ExpandOne(EditOp op, int opIndex, Workspace workspace, HandlerRegistry handlers)
    {
        if (op.Props![PropName] is not JsonValue specValue || !specValue.TryGetValue<string>(out var spec))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}].props.{PropName} must be a string.",
                SpecShape);
        }

        var bang = spec.IndexOf('!', StringComparison.Ordinal);
        if (bang <= 0 || bang == spec.Length - 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}].props.{PropName} is '{spec}' — expected <workbook>!<sheet>/<range>.",
                SpecShape);
        }

        var fileText = spec[..bang];
        var addressText = spec[(bang + 1)..];

        // Sandbox-resolve the workbook like every other file-valued prop.
        var resolved = workspace.Resolve(fileText, mustExist: true);
        var handler = handlers.Resolve(resolved);
        if (handler.Kind != DocumentKind.Xlsx)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}].props.{PropName} source must be a workbook, got '{fileText}'.",
                SpecShape);
        }

        var path = addressText.StartsWith('/') ? addressText : "/" + addressText;
        var data = GetRange(handler, workspace, resolved, path, spec, opIndex);
        var (categories, series) = ToChartData(data, handler, workspace, resolved, fileText, spec, opIndex);

        // Rebuild props: everything except dataFrom survives; literals go in.
        var props = new JsonObject();
        foreach (var (key, value) in op.Props)
        {
            if (!string.Equals(key, PropName, StringComparison.Ordinal))
            {
                props[key] = value?.DeepClone();
            }
        }

        props["categories"] = categories;
        props["series"] = series;
        return op with { Props = props };
    }

    /// <summary>Reads the source range via the xlsx handler; failures are re-thrown with the dataFrom context.</summary>
    private static JsonObject GetRange(
        IFormatHandler handler, Workspace workspace, string resolved, string path, string spec, int opIndex)
    {
        var envelope = handler.Get(new CommandContext
        {
            Workspace = workspace,
            File = resolved,
            Args = new JsonObject { ["path"] = path, ["maxCells"] = MaxSourceCells },
        });

        if (!envelope.IsOk)
        {
            var error = envelope.Error!;
            throw new AiofficeException(
                error.Code,
                $"ops[{opIndex}].props.{PropName} could not read '{spec}': {error.Message}",
                error.Suggestion,
                error.Candidates);
        }

        if (JsonSerializer.SerializeToNode(envelope.Data, JsonDefaults.Options) is not JsonObject data)
        {
            throw new AiofficeException(
                ErrorCodes.InternalError,
                $"The xlsx handler returned no readable payload for '{spec}'.",
                "Run 'aioffice get <workbook> <range>' to inspect the source; report this if it persists.");
        }

        return data;
    }

    /// <summary>First column → categories, remaining columns → series, header row → names.</summary>
    private static (JsonArray Categories, JsonArray Series) ToChartData(
        JsonObject data, IFormatHandler handler, Workspace workspace, string resolved,
        string fileText, string spec, int opIndex)
    {
        var kind = data["kind"]?.GetValue<string>();
        var sheet = data["sheet"]?.GetValue<string>();
        if (kind != "range" || data["values"] is not JsonArray rows)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}].props.{PropName} target '{spec}' is a {kind ?? "non-range"}, but charts need a range.",
                SpecShape,
                candidates: NearestRangeCandidates(handler, workspace, resolved, fileText, sheet));
        }

        var columns = (data["columns"] as JsonValue)?.GetValue<int>() ?? 0;
        if (rows.Count < 2 || columns < 2)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}].props.{PropName} range '{spec}' is {rows.Count}x{columns}, " +
                "but charts need at least a header row plus one data row, and at least two columns.",
                SpecShape,
                candidates: NearestRangeCandidates(handler, workspace, resolved, fileText, sheet));
        }

        if (data["truncated"] is JsonValue tv && tv.TryGetValue<bool>(out var truncated) && truncated)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}].props.{PropName} range '{spec}' exceeds {MaxSourceCells} cells.",
                "Point dataFrom at the aggregated summary range you actually want to chart, not the raw data.");
        }

        var header = rows[0]!.AsArray();
        var series = new JsonArray();
        for (var column = 1; column < columns; column++)
        {
            var name = ScalarText(column < header.Count ? header[column] : null) is { Length: > 0 } text
                ? text
                : Inv($"Series {column}");
            var values = new JsonArray();
            for (var row = 1; row < rows.Count; row++)
            {
                values.Add(NumericValue(rows[row]!.AsArray(), row, column, spec, opIndex));
            }

            series.Add(new JsonObject { ["name"] = name, ["values"] = values });
        }

        var categories = new JsonArray();
        for (var row = 1; row < rows.Count; row++)
        {
            var cells = rows[row]!.AsArray();
            // Blank label cells become "" — the chart writer wants scalar labels, not nulls.
            categories.Add(JsonValue.Create(ScalarText(cells.Count > 0 ? cells[0] : null) ?? string.Empty));
        }

        return (categories, series);
    }

    /// <summary>One numeric series value; null stays a gap; anything non-numeric is a typed error.</summary>
    private static JsonNode? NumericValue(JsonArray cells, int row, int column, string spec, int opIndex)
    {
        var cell = column < cells.Count ? cells[column] : null;
        if (cell is null)
        {
            return null; // blank cell = literal gap, same contract as props.series values
        }

        if (cell is JsonValue value)
        {
            if (value.TryGetValue<double>(out var number))
            {
                return JsonValue.Create(number);
            }

            if (value.TryGetValue<bool>(out var flag))
            {
                return JsonValue.Create(flag ? 1d : 0d);
            }

            if (value.TryGetValue<string>(out var text) &&
                double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            {
                return JsonValue.Create(parsed);
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            Inv($"ops[{opIndex}].props.{PropName} '{spec}': data row {row}, column {column + 1} ") +
            $"is not numeric: {cell.ToJsonString(JsonDefaults.Options)}.",
            "Series columns (everything right of the first column, below the header row) must hold numbers; " +
            "move labels into the first column or the header row.");
    }

    /// <summary>Sheet-name and used-range candidates for a wrong dataFrom address (best-effort).</summary>
    private static IReadOnlyList<string>? NearestRangeCandidates(
        IFormatHandler handler, Workspace workspace, string resolved, string fileText, string? sheet)
    {
        try
        {
            if (sheet is null)
            {
                return null;
            }

            var info = handler.Get(new CommandContext
            {
                Workspace = workspace,
                File = resolved,
                Args = new JsonObject { ["path"] = "/" + (sheet.Contains(' ', StringComparison.Ordinal) ? "'" + sheet + "'" : sheet) },
            });
            if (!info.IsOk ||
                JsonSerializer.SerializeToNode(info.Data, JsonDefaults.Options) is not JsonObject sheetData ||
                sheetData["usedRange"]?.GetValue<string>() is not { Length: > 0 } usedRange)
            {
                return null;
            }

            var sheetText = sheet.Contains(' ', StringComparison.Ordinal) ? "'" + sheet + "'" : sheet;
            return [Inv($"{fileText}!{sheetText}/{usedRange}")];
        }
        catch (Exception ex) when (ex is AiofficeException or JsonException or NotSupportedException)
        {
            return null; // candidates are best-effort sugar; the main error stands alone
        }
    }

    private static string? ScalarText(JsonNode? node) => node switch
    {
        null => null,
        JsonValue v when v.TryGetValue<string>(out var s) => s,
        JsonValue v when v.TryGetValue<double>(out var d) => d.ToString(CultureInfo.InvariantCulture),
        JsonValue v when v.TryGetValue<bool>(out var b) => b ? "true" : "false",
        _ => node.ToJsonString(JsonDefaults.Options),
    };

    private static string Inv(FormattableString text) => FormattableString.Invariant(text);
}
