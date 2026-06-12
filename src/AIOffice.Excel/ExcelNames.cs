using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// Defined names (named ranges), ClosedXML-native. Addressing uses the stable
/// id form the shared DocPath grammar cannot express, so it is pre-parsed here
/// (precedent: pivot[@name=…]): <c>/name[@name=SalesData]</c> for
/// workbook-scoped names, <c>/Sheet1/name[@name=SalesData]</c> for
/// sheet-scoped ones. Formulas can reference the names afterwards — the
/// ClosedXML engine resolves them when caching values on save.
/// </summary>
internal static partial class ExcelNames
{
    private static readonly IReadOnlyList<string> Scopes = ["workbook", "sheet"];

    private static readonly IReadOnlyList<string> AddProps = ["name", "scope"];

    /// <summary><c>/name[@name=X]</c> or <c>/Sheet/name[@name=X]</c> (bare or 'quoted' name).</summary>
    [GeneratedRegex(@"^(?<sheet>/.+)?/(?i:name)\[@name=(?:'(?<quoted>(?:[^']|'')+)'|(?<bare>[^\]]+))\]$")]
    private static partial Regex NamePath();

    /// <summary>Excel defined-name rules: start letter/_/\, then letters, digits, . _ \, no spaces.</summary>
    [GeneratedRegex(@"^[\p{L}_\\][\p{L}\p{N}_.\\]*$")]
    private static partial Regex ValidName();

    [GeneratedRegex("^[A-Za-z]{1,3}[0-9]{1,7}$")]
    private static partial Regex LooksLikeCell();

    /// <summary>Detects the defined-name path form; the sheet part (when present) is the raw sheet path.</summary>
    public static bool TryParsePath(string pathText, out string? sheetPath, out string name)
    {
        sheetPath = null;
        name = string.Empty;
        var match = NamePath().Match(pathText);
        if (!match.Success)
        {
            return false;
        }

        sheetPath = match.Groups["sheet"].Success ? match.Groups["sheet"].Value : null;
        name = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value.Replace("''", "'", StringComparison.Ordinal)
            : match.Groups["bare"].Value;
        return true;
    }

    /// <summary>The canonical path aioffice emits for a defined name.</summary>
    public static string PathOf(IXLWorksheet? scopeSheet, string name) =>
        (scopeSheet is null ? string.Empty : ExcelPaths.SheetPath(scopeSheet)) +
        $"/name[@name={ExcelPaths.QuoteSheet(name)}]";

    /// <summary>
    /// Resolves a defined name. A sheet-qualified path searches only that
    /// sheet's scope; the bare form searches the workbook scope first, then
    /// every sheet. Misses throw <c>invalid_path</c> with nearest-match candidates.
    /// </summary>
    public static (IXLDefinedName Name, IXLWorksheet? ScopeSheet) Find(
        XLWorkbook workbook, string? sheetPath, string name)
    {
        if (sheetPath is not null)
        {
            var target = ExcelPaths.Resolve(workbook, sheetPath);
            if (target.Kind != ExcelTargetKind.Sheet)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"name[@name=…] must follow a sheet name: {sheetPath}",
                    "Use /SheetName/name[@name=X] for sheet-scoped names, or /name[@name=X] for workbook scope.");
            }

            if (target.Sheet.DefinedNames.TryGetValue(name, out var sheetScoped) && sheetScoped is not null)
            {
                return (sheetScoped, target.Sheet);
            }

            throw NotFound(workbook, name);
        }

        if (workbook.DefinedNames.TryGetValue(name, out var workbookScoped) && workbookScoped is not null)
        {
            return (workbookScoped, null);
        }

        foreach (var sheet in workbook.Worksheets.OrderBy(ws => ws.Position))
        {
            if (sheet.DefinedNames.TryGetValue(name, out var found) && found is not null)
            {
                return (found, sheet);
            }
        }

        throw NotFound(workbook, name);
    }

    private static AiofficeException NotFound(XLWorkbook workbook, string requested)
    {
        var candidates = All(workbook)
            .OrderBy(n => ExcelPaths.Levenshtein(requested, n.Name.Name))
            .Select(n => PathOf(n.ScopeSheet, n.Name.Name))
            .Take(5)
            .ToList();
        return new AiofficeException(
            ErrorCodes.InvalidPath,
            $"No defined name '{requested}' exists in the workbook.",
            candidates.Count > 0
                ? "Defined names are matched case-insensitively; pick one of the candidates."
                : "This workbook has no defined names; add one with " +
                  "{op:add, type:name, path:/Sheet1/A1:B5, props:{name:\"SalesData\"}}.",
            candidates: candidates.Count > 0 ? candidates : null);
    }

    private static IEnumerable<(IXLDefinedName Name, IXLWorksheet? ScopeSheet)> All(XLWorkbook workbook)
    {
        // Builtin bookkeeping names (_xlnm.Print_Area, _xlnm._FilterDatabase…)
        // are surfaced through their own features (pageSetup.printArea,
        // autoFilter), not as user-defined names.
        foreach (var name in workbook.DefinedNames.Where(n => !IsBuiltin(n)))
        {
            yield return (name, null);
        }

        foreach (var sheet in workbook.Worksheets.OrderBy(ws => ws.Position))
        {
            foreach (var name in sheet.DefinedNames.Where(n => !IsBuiltin(n)))
            {
                yield return (name, sheet);
            }
        }
    }

    private static bool IsBuiltin(IXLDefinedName name) =>
        name.Name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase);

    /// <summary>One defined name for get / structure (ranges as canonical paths).</summary>
    public static object Describe(IXLDefinedName name, IXLWorksheet? scopeSheet) => new
    {
        path = PathOf(scopeSheet, name.Name),
        kind = "name",
        name = name.Name,
        scope = scopeSheet is null ? "workbook" : "sheet",
        sheet = scopeSheet?.Name,
        refersTo = name.RefersTo,
        ranges = SafeRangePaths(name),
    };

    /// <summary>Every defined name in the workbook (workbook scope first, then per sheet).</summary>
    public static List<object> ListAll(XLWorkbook workbook) =>
        [.. All(workbook).Select(n => Describe(n.Name, n.ScopeSheet))];

    private static List<string> SafeRangePaths(IXLDefinedName name)
    {
        try
        {
            return [.. name.Ranges.Select(r => ExcelPaths.RangePath(r.Worksheet, r.RangeAddress))];
        }
        catch (Exception)
        {
            return []; // names can refer to constants or broken refs; refersTo still tells the truth
        }
    }

    /// <summary>
    /// Applies an <c>add name</c> op: the op path addresses the range the name
    /// refers to; props carry the name and its scope (workbook default).
    /// </summary>
    public static object Add(XLWorkbook workbook, ExcelTarget target, EditOp op, int index)
    {
        if (target.Kind is not (ExcelTargetKind.Range or ExcelTargetKind.Cell))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add name targets the range the name refers to, like /Sheet1/A1:B5.",
                "Use {op:add, type:name, path:/Sheet1/A1:B5, props:{name:\"SalesData\", scope:\"workbook\"}}.");
        }

        var props = op.Props ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: add name needs props.",
            "Pass props like {\"name\":\"SalesData\"} (scope defaults to workbook).");

        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: unknown name prop '{key}'.",
                    "Supported name props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }

        var name = StringProp(props, "name");
        if (string.IsNullOrEmpty(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add name needs the 'name' prop.",
                "Pass it as a string, e.g. {\"name\":\"SalesData\"}.");
        }

        if (name.Length > 255 || !ValidName().IsMatch(name) || LooksLikeCell().IsMatch(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: '{name}' is not a usable defined name.",
                @"Names start with a letter, _ or \, continue with letters, digits, . _ \ (no spaces), " +
                "must not look like a cell reference (A1), and are at most 255 characters.");
        }

        var scope = StringProp(props, "scope") ?? "workbook";
        if (!Scopes.Contains(scope, StringComparer.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: unknown name scope '{scope}'.",
                "Use scope \"workbook\" (default, visible everywhere) or \"sheet\" (visible on the range's sheet).",
                candidates: Scopes);
        }

        var names = scope == "sheet" ? target.Sheet.DefinedNames : workbook.DefinedNames;
        if (names.Contains(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: a defined name '{name}' already exists in {scope} scope.",
                "Remove it first ({op:remove, path:" + PathOf(scope == "sheet" ? target.Sheet : null, name) +
                "}) or pick a different name.");
        }

        var range = target.Range ?? target.Cell!.AsRange();
        IXLDefinedName created;
        try
        {
            created = names.Add(name, range);
        }
        catch (ArgumentException exception)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: could not create the defined name: {exception.Message}",
                "Names must be unique per scope and follow Excel's naming rules.",
                innerException: exception);
        }

        return new
        {
            op = "add",
            type = "name",
            path = PathOf(scope == "sheet" ? target.Sheet : null, name),
            name = created.Name,
            scope,
            refersTo = created.RefersTo,
        };
    }

    private static string? StringProp(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;
}
