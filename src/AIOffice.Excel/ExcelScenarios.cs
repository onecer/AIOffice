using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// (1.5) Scenario Manager. <c>{op:add, type:scenario, path:/Sheet1,
/// props:{name:"Best Case", cells:{"B1":120,"B2":0.15}, comment?}}</c> saves a
/// named scenario (a set of changing cells and the values they take) into the
/// worksheet's <c>&lt;scenarios&gt;</c> part. Scenarios are addressed for
/// <c>get</c>/<c>remove</c> as <c>/Sheet1/scenario[@name=…]</c>, and surface in
/// <c>read --view structure</c> under each sheet's <c>scenarios</c> array.
///
/// <para><c>{op:set, path:/Sheet1, props:{applyScenario:"Best Case"}}</c> writes
/// the scenario's stored values into its changing cells and recalculates, so the
/// workbook reflects that scenario.</para>
///
/// <para>ClosedXML 0.105 has no scenario model, so the <c>&lt;scenarios&gt;</c>
/// element (and its schema-ordered placement after <c>sheetProtection</c>) is
/// authored raw in a post-save pass, exactly like sheet protection and data
/// tables. Validator-clean.</para>
/// </summary>
internal static class ExcelScenarios
{
    private static readonly IReadOnlyList<string> AddProps = ["name", "cells", "comment"];

    /// <summary>A scenario queued for the post-save raw write (changing cells + their values).</summary>
    public sealed record Pending(
        string Sheet,
        string Name,
        string? Comment,
        IReadOnlyList<(string Cell, XLCellValue Value)> Cells);

    // ----- add ---------------------------------------------------------------

    /// <summary>Validates and applies an <c>add scenario</c> op; returns the details entry and queues the raw write.</summary>
    public static object Add(ExcelTarget target, EditOp op, int opIndex, List<Pending> queue)
    {
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add scenario targets a sheet path like /Sheet1; the changing cells go in props.cells.",
                "Use {op:add, type:scenario, path:/Sheet1, props:{name:\"Best Case\", cells:{\"B1\":120,\"B2\":0.15}}}.");
        }

        var props = op.Props ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add scenario needs props with a name and changing cells.",
            "Use {op:add, type:scenario, path:/Sheet1, props:{name:\"Best Case\", cells:{\"B1\":120}}}.");
        GuardProps(props, opIndex);

        var name = StringProp(props, "name");
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add scenario needs a non-empty name.",
                "Name the scenario so it can be applied/removed: props:{name:\"Best Case\", cells:{…}}.");
        }

        if (props.TryGetPropertyValue("cells", out var cellsNode) && cellsNode is JsonObject cellsObj && cellsObj.Count > 0)
        {
            var cells = ParseCells(cellsObj, target.Sheet, opIndex);
            var comment = StringProp(props, "comment");
            queue.Add(new Pending(target.Sheet.Name, name, comment, cells));
            return new
            {
                op = "add",
                type = "scenario",
                path = ScenarioPath(target.Sheet, name),
                name,
                cells = cells.Count,
            };
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add scenario needs a non-empty 'cells' object mapping changing cells to values.",
            "Use props:{name:\"Best Case\", cells:{\"B1\":120,\"B2\":0.15}} — one entry per changing cell.");
    }

    /// <summary>Parses the cells object into (cellRef, typed value) pairs, validating each address against the sheet.</summary>
    private static List<(string Cell, XLCellValue Value)> ParseCells(JsonObject cellsObj, IXLWorksheet sheet, int opIndex)
    {
        var cells = new List<(string Cell, XLCellValue Value)>(cellsObj.Count);
        foreach (var (cellRef, valueNode) in cellsObj)
        {
            IXLCell address;
            try
            {
                address = sheet.Cell(cellRef);
            }
            catch (Exception)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: scenario changing-cell '{cellRef}' is not a valid single-cell address.",
                    "Each key in 'cells' is one cell like B1; ranges are not allowed.");
            }

            var parsed = ExcelValues.Parse(valueNode);
            if (parsed.IsFormula)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: scenario changing-cell '{cellRef}' was given a formula; scenarios store constant values.",
                    "A scenario stores the value a changing cell TAKES, not a formula. Pass a number/text/boolean.");
            }

            cells.Add((address.Address.ToString()!, parsed.Value));
        }

        return cells;
    }

    // ----- post-save raw write (add) -----------------------------------------

    /// <summary>
    /// Authors the queued scenarios raw: a <c>&lt;scenarios&gt;</c> element per
    /// sheet (created or extended) holding one <c>&lt;scenario&gt;</c> per pending,
    /// each with its <c>&lt;inputCells&gt;</c>. ClosedXML left the worksheet's
    /// scenarios untouched, so this is the single writer.
    /// </summary>
    public static void ApplyAfterSave(string file, IReadOnlyList<Pending> scenarios)
    {
        if (scenarios.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart!;

        foreach (var group in scenarios.GroupBy(s => s.Sheet, StringComparer.OrdinalIgnoreCase))
        {
            if (WorksheetFor(workbookPart, group.Key) is not { } worksheet)
            {
                continue;
            }

            var scenariosElement = worksheet.Elements<S.Scenarios>().FirstOrDefault();
            if (scenariosElement is null)
            {
                scenariosElement = new S.Scenarios { Current = 0, Show = 0 };
                InsertScenarios(worksheet, scenariosElement);
            }

            foreach (var pending in group)
            {
                // A re-add of the same name replaces the earlier scenario.
                foreach (var dup in scenariosElement.Elements<S.Scenario>()
                             .Where(s => string.Equals(s.Name?.Value, pending.Name, StringComparison.OrdinalIgnoreCase))
                             .ToList())
                {
                    dup.Remove();
                }

                scenariosElement.AppendChild(BuildScenario(pending));
            }

            worksheet.Save();
        }

        workbookPart.Workbook!.Save();
    }

    private static S.Scenario BuildScenario(Pending pending)
    {
        var scenario = new S.Scenario
        {
            Name = pending.Name,
            Locked = true,
            Count = (uint)pending.Cells.Count,
        };
        if (!string.IsNullOrEmpty(pending.Comment))
        {
            scenario.Comment = pending.Comment;
        }

        foreach (var (cell, value) in pending.Cells)
        {
            scenario.AppendChild(new S.InputCells
            {
                CellReference = cell,
                Val = ScenarioValueText(value),
            });
        }

        return scenario;
    }

    /// <summary>The string form a scenario stores in its inputCells val (numbers invariant, booleans 1/0, text as-is).</summary>
    private static string ScenarioValueText(XLCellValue value) => value.Type switch
    {
        XLDataType.Number => value.GetNumber().ToString("R", CultureInfo.InvariantCulture),
        XLDataType.Boolean => value.GetBoolean() ? "1" : "0",
        XLDataType.DateTime => value.GetDateTime().ToOADate().ToString("R", CultureInfo.InvariantCulture),
        _ => value.GetText(),
    };

    /// <summary>
    /// Inserts a fresh <c>&lt;scenarios&gt;</c> in schema order: after
    /// sheetData/sheetProtection/protectedRanges, before autoFilter/mergeCells/etc.
    /// </summary>
    private static void InsertScenarios(S.Worksheet worksheet, S.Scenarios scenarios)
    {
        var successor = worksheet.ChildElements.FirstOrDefault(e =>
            e is S.AutoFilter or S.SortState or S.DataConsolidate or S.CustomSheetViews or S.MergeCells
                or S.PhoneticProperties or S.ConditionalFormatting or S.DataValidations or S.Hyperlinks
                or S.PrintOptions or S.PageMargins or S.PageSetup or S.HeaderFooter or S.RowBreaks
                or S.ColumnBreaks or S.CustomProperties or S.CellWatches or S.IgnoredErrors or S.Drawing
                or S.LegacyDrawing or S.Picture or S.OleObjects or S.Controls or S.TableParts or S.ExtensionList);
        if (successor is not null)
        {
            worksheet.InsertBefore(scenarios, successor);
            return;
        }

        DocumentFormat.OpenXml.OpenXmlElement? anchor =
            worksheet.Elements<S.ProtectedRanges>().FirstOrDefault();
        anchor ??= worksheet.Elements<S.SheetProtection>().FirstOrDefault();
        anchor ??= worksheet.Elements<S.SheetCalculationProperties>().FirstOrDefault();
        anchor ??= worksheet.Elements<S.SheetData>().FirstOrDefault();
        if (anchor is not null)
        {
            worksheet.InsertAfter(scenarios, anchor);
        }
        else
        {
            worksheet.AppendChild(scenarios);
        }
    }

    // ----- read back (structure / get / remove) ------------------------------

    /// <summary>One scenario read raw from the saved bytes.</summary>
    public sealed record Info(string Sheet, string Name, string? Comment, IReadOnlyList<(string Cell, string Val)> Cells);

    /// <summary>Reads every scenario on a sheet, in document order.</summary>
    public static List<Info> ReadOnSheet(string file, string sheetName)
    {
        using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        return ReadOnSheet(document.WorkbookPart!, sheetName);
    }

    /// <summary>Reads every scenario in the workbook (all sheets) from an open document, keyed by sheet name.</summary>
    public static ILookup<string, Info> ReadAll(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart;
        var all = new List<Info>();
        if (workbookPart?.Workbook is { } workbook)
        {
            foreach (var sheet in workbook.Descendants<S.Sheet>())
            {
                if (sheet.Name?.Value is { } name)
                {
                    all.AddRange(ReadOnSheet(workbookPart, name));
                }
            }
        }

        return all.ToLookup(s => s.Sheet, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Reads the scenarios on a sheet from an already-open workbook part.</summary>
    public static List<Info> ReadOnSheet(WorkbookPart workbookPart, string sheetName)
    {
        if (WorksheetFor(workbookPart, sheetName) is not { } worksheet ||
            worksheet.Elements<S.Scenarios>().FirstOrDefault() is not { } scenariosElement)
        {
            return [];
        }

        var result = new List<Info>();
        foreach (var scenario in scenariosElement.Elements<S.Scenario>())
        {
            var cells = scenario.Elements<S.InputCells>()
                .Where(c => c.CellReference?.Value is not null)
                .Select(c => (c.CellReference!.Value!, c.Val?.Value ?? string.Empty))
                .ToList();
            result.Add(new Info(sheetName, scenario.Name?.Value ?? string.Empty, scenario.Comment?.Value, cells));
        }

        return result;
    }

    /// <summary>Describes one scenario for <c>get</c> / <c>read --view structure</c>.</summary>
    public static object Describe(IXLWorksheet sheet, Info info) => new
    {
        path = ScenarioPath(sheet, info.Name),
        kind = "scenario",
        sheet = info.Sheet,
        name = info.Name,
        comment = info.Comment,
        cells = info.Cells.ToDictionary(c => c.Cell, c => (object)TypeValue(c.Val), StringComparer.Ordinal),
    };

    /// <summary>The structure-view shape for a scenario (path + name + cell count, no full values).</summary>
    public static object DescribeBrief(IXLWorksheet sheet, Info info) => new
    {
        path = ScenarioPath(sheet, info.Name),
        name = info.Name,
        comment = info.Comment,
        changingCells = info.Cells.Select(c => c.Cell).ToList(),
    };

    /// <summary>Types a stored scenario val for the wire (number/boolean/text).</summary>
    private static object TypeValue(string val) =>
        double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? number
            : bool.TryParse(val, out var b)
                ? b
                : val;

    /// <summary>Types a stored scenario val as a typed cell value for applying it back into a cell.</summary>
    public static XLCellValue TypeStoredValue(string val) =>
        double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            ? number
            : bool.TryParse(val, out var b)
                ? b
                : val;

    // ----- remove ------------------------------------------------------------

    /// <summary>
    /// Removes the named scenarios from the saved bytes; drops the empty
    /// <c>&lt;scenarios&gt;</c> element when the last one goes. Names are resolved
    /// case-insensitively against the on-disk file.
    /// </summary>
    public static void RemoveAfterSave(string file, IReadOnlyList<(string Sheet, string Name)> removals)
    {
        if (removals.Count == 0)
        {
            return;
        }

        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart!;

        foreach (var group in removals.GroupBy(r => r.Sheet, StringComparer.OrdinalIgnoreCase))
        {
            if (WorksheetFor(workbookPart, group.Key) is not { } worksheet ||
                worksheet.Elements<S.Scenarios>().FirstOrDefault() is not { } scenariosElement)
            {
                continue;
            }

            var wanted = group.Select(r => r.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var scenario in scenariosElement.Elements<S.Scenario>()
                         .Where(s => wanted.Contains(s.Name?.Value ?? string.Empty))
                         .ToList())
            {
                scenario.Remove();
            }

            if (!scenariosElement.Elements<S.Scenario>().Any())
            {
                scenariosElement.Remove();
            }

            worksheet.Save();
        }

        workbookPart.Workbook!.Save();
    }

    // ----- helpers -----------------------------------------------------------

    /// <summary>The canonical scenario path <c>/Sheet/scenario[@name=Name]</c> (quoted when needed).</summary>
    public static string ScenarioPath(IXLWorksheet sheet, string name) =>
        $"{ExcelPaths.SheetPath(sheet)}/scenario[@name={ExcelPaths.QuoteSheet(name)}]";

    private static S.Worksheet? WorksheetFor(WorkbookPart workbookPart, string sheetName)
    {
        var sheetElement = workbookPart.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheetElement?.Id?.Value is not { } relId ||
            workbookPart.GetPartById(relId) is not WorksheetPart worksheetPart)
        {
            return null;
        }

        return worksheetPart.Worksheet;
    }

    private static void GuardProps(JsonObject props, int opIndex)
    {
        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown scenario prop '{key}'.",
                    "Supported props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }
    }

    private static string? StringProp(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == System.Text.Json.JsonValueKind.String
            ? value.GetValue<string>()
            : null;
}
