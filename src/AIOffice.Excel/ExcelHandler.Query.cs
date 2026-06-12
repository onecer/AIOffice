using System.Globalization;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

public sealed partial class ExcelHandler
{
    private const int MaxQueryMatches = 500;

    private static readonly IReadOnlyList<string> CellAttributes =
        ["value", "formula", "sheet", "numberFormat", "bold", "italic"];

    public Envelope Query(CommandContext ctx) => Run(ctx, sw =>
    {
        var file = RequireFile(ctx, mustExist: true);
        var selectorText = ArgString(ctx, "selector") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "query needs a selector.",
            "Pass a selector like cell[value>100], cell[formula] or row[Qty>5].");

        var selector = Selector.Parse(RewriteBareAttributes(selectorText));
        ValidateOrderingPredicates(selector);

        using var workbook = OpenWorkbook(file);
        var (matches, total) = selector.Element switch
        {
            "cell" => QueryCells(workbook, selector),
            "row" => QueryRows(workbook, selector),
            "sheet" => QuerySheets(workbook, selector),
            _ => throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"xlsx selectors start with cell, row or sheet; got '{selector.Element}'.",
                "Use cell[...] for cells, row[...] for table rows, sheet[...] for sheets.",
                candidates: ["cell", "row", "sheet"]),
        };

        List<Warning>? warnings = total > matches.Count
            ? [new Warning(
                "result_truncated",
                $"{total} elements matched; returning the first {matches.Count}. Make the selector more specific.")]
            : null;

        return Envelope.Ok(
            new { selector = selector.ToCanonicalString(), total, matches },
            MetaFor(file, sw, warnings));
    });

    [GeneratedRegex(@"\[([A-Za-z_][A-Za-z0-9_.-]*)\]")]
    private static partial Regex BareAttribute();

    /// <summary>
    /// Existence sugar: <c>cell[formula]</c> means "has a formula". The core
    /// grammar requires an operator, so bare attributes are rewritten to
    /// <c>[attr=*]</c> and <c>=*</c> is interpreted as an existence test.
    /// Skipped when the selector contains quotes, to never touch quoted text.
    /// </summary>
    private static string RewriteBareAttributes(string selector) =>
        selector.Contains('\'') || selector.Contains('"')
            ? selector
            : BareAttribute().Replace(selector, "[$1=*]");

    private static bool IsExistenceTest(AttributePredicate predicate) =>
        predicate.Op == SelectorOperator.Equals && predicate.Value == "*";

    private static void ValidateOrderingPredicates(Selector selector)
    {
        foreach (var predicate in selector.Predicates.OfType<AttributePredicate>())
        {
            var ordering = predicate.Op is SelectorOperator.GreaterThan or SelectorOperator.GreaterOrEqual
                or SelectorOperator.LessThan or SelectorOperator.LessOrEqual;
            if (ordering && predicate.NumericValue is null)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"[{predicate.Attribute}{OpToken(predicate.Op)}{predicate.Value}] compares with an ordering operator but '{predicate.Value}' is not a number.",
                    "Ordering operators (> >= < <=) need numeric values, e.g. cell[value>100].");
            }
        }
    }

    private static string OpToken(SelectorOperator op) => op switch
    {
        SelectorOperator.GreaterThan => ">",
        SelectorOperator.GreaterOrEqual => ">=",
        SelectorOperator.LessThan => "<",
        SelectorOperator.LessOrEqual => "<=",
        SelectorOperator.NotEquals => "!=",
        SelectorOperator.ContainsText => "*=",
        _ => "=",
    };

    private static (List<object> Matches, int Total) QueryCells(XLWorkbook workbook, Selector selector)
    {
        var matches = new List<object>();
        var total = 0;
        foreach (var sheet in workbook.Worksheets.OrderBy(ws => ws.Position))
        {
            foreach (var cell in sheet.CellsUsed())
            {
                if (!selector.Predicates.All(p => CellMatches(sheet, cell, p)))
                {
                    continue;
                }

                total++;
                if (matches.Count >= MaxQueryMatches)
                {
                    continue;
                }

                XLCellValue value;
                try
                {
                    value = cell.Value;
                }
                catch (Exception)
                {
                    value = cell.CachedValue;
                }

                matches.Add(new
                {
                    path = ExcelPaths.CellPath(sheet, cell.Address),
                    sheet = sheet.Name,
                    address = cell.Address.ToString(),
                    value = ExcelValues.ToJson(value),
                    type = ExcelValues.TypeName(value.Type),
                    formula = cell.HasFormula ? "=" + cell.FormulaA1 : null,
                });
            }
        }

        return (matches, total);
    }

    private static bool CellMatches(IXLWorksheet sheet, IXLCell cell, SelectorPredicate predicate)
    {
        if (predicate is ContainsPredicate contains)
        {
            return ExcelValues.SafeFormatted(cell).Contains(contains.Text, StringComparison.OrdinalIgnoreCase);
        }

        var attr = (AttributePredicate)predicate;
        switch (attr.Attribute)
        {
            case "value":
            {
                XLCellValue value;
                try
                {
                    value = cell.Value;
                }
                catch (Exception)
                {
                    value = cell.CachedValue;
                }

                return IsExistenceTest(attr) ? !value.IsBlank : CompareValue(value, cell, attr);
            }

            case "formula":
            {
                if (IsExistenceTest(attr))
                {
                    return cell.HasFormula;
                }

                if (!cell.HasFormula)
                {
                    return attr.Op == SelectorOperator.NotEquals;
                }

                var formula = cell.FormulaA1;
                var wanted = attr.Value.StartsWith('=') ? attr.Value[1..] : attr.Value;
                return attr.Op switch
                {
                    SelectorOperator.Equals => string.Equals(formula, wanted, StringComparison.OrdinalIgnoreCase),
                    SelectorOperator.NotEquals => !string.Equals(formula, wanted, StringComparison.OrdinalIgnoreCase),
                    SelectorOperator.ContainsText => formula.Contains(wanted, StringComparison.OrdinalIgnoreCase),
                    _ => false,
                };
            }

            case "sheet":
                return attr.Op switch
                {
                    SelectorOperator.Equals => string.Equals(sheet.Name, attr.Value, StringComparison.OrdinalIgnoreCase),
                    SelectorOperator.NotEquals => !string.Equals(sheet.Name, attr.Value, StringComparison.OrdinalIgnoreCase),
                    SelectorOperator.ContainsText => sheet.Name.Contains(attr.Value, StringComparison.OrdinalIgnoreCase),
                    _ => false,
                };

            case "numberFormat":
            {
                var format = cell.Style.NumberFormat.Format;
                return attr.Op switch
                {
                    SelectorOperator.Equals => IsExistenceTest(attr)
                        ? !string.IsNullOrEmpty(format)
                        : string.Equals(format, attr.Value, StringComparison.Ordinal),
                    SelectorOperator.NotEquals => !string.Equals(format, attr.Value, StringComparison.Ordinal),
                    SelectorOperator.ContainsText => format.Contains(attr.Value, StringComparison.OrdinalIgnoreCase),
                    _ => false,
                };
            }

            case "bold":
                return MatchFlag(cell.Style.Font.Bold, attr);

            case "italic":
                return MatchFlag(cell.Style.Font.Italic, attr);

            default:
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown cell attribute '{attr.Attribute}' in selector.",
                    "Cell selectors support: " + string.Join(", ", CellAttributes) + ".",
                    candidates: CellAttributes);
        }
    }

    private static bool MatchFlag(bool actual, AttributePredicate attr)
    {
        if (IsExistenceTest(attr))
        {
            return actual;
        }

        var wanted = string.Equals(attr.Value, "true", StringComparison.OrdinalIgnoreCase);
        return attr.Op switch
        {
            SelectorOperator.Equals => actual == wanted,
            SelectorOperator.NotEquals => actual != wanted,
            _ => false,
        };
    }

    private static bool CompareValue(XLCellValue value, IXLCell cell, AttributePredicate attr)
    {
        if (value.IsNumber && attr.NumericValue is { } number)
        {
            var actual = value.GetNumber();
            return attr.Op switch
            {
                SelectorOperator.Equals => actual.Equals(number),
                SelectorOperator.NotEquals => !actual.Equals(number),
                SelectorOperator.GreaterThan => actual > number,
                SelectorOperator.GreaterOrEqual => actual >= number,
                SelectorOperator.LessThan => actual < number,
                SelectorOperator.LessOrEqual => actual <= number,
                SelectorOperator.ContainsText => FormattedContains(cell, attr.Value),
                _ => false,
            };
        }

        var text = value.IsText ? value.GetText() : ExcelValues.SafeFormatted(cell);
        return attr.Op switch
        {
            SelectorOperator.Equals => string.Equals(text, attr.Value, StringComparison.Ordinal),
            SelectorOperator.NotEquals => !string.Equals(text, attr.Value, StringComparison.Ordinal),
            SelectorOperator.ContainsText => text.Contains(attr.Value, StringComparison.OrdinalIgnoreCase),
            _ => false, // ordering against a non-numeric cell: no match
        };
    }

    private static bool FormattedContains(IXLCell cell, string value) =>
        ExcelValues.SafeFormatted(cell).Contains(value, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Table-row queries: attributes are table column names, e.g.
    /// <c>row[Qty>5]</c>. Tables missing a referenced column are skipped; when
    /// no table has it at all, the error lists every known column.
    /// </summary>
    private static (List<object> Matches, int Total) QueryRows(XLWorkbook workbook, Selector selector)
    {
        var attrPredicates = selector.Predicates.OfType<AttributePredicate>().ToList();
        var containsPredicates = selector.Predicates.OfType<ContainsPredicate>().ToList();
        var columnPredicates = attrPredicates
            .Where(p => p.Attribute is not ("sheet" or "table"))
            .ToList();
        var filterPredicates = attrPredicates.Except(columnPredicates).ToList();

        var matches = new List<object>();
        var total = 0;
        var seenColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var matchedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var sheet in workbook.Worksheets.OrderBy(ws => ws.Position))
        {
            foreach (var table in sheet.Tables)
            {
                if (!filterPredicates.All(p => TableFilterMatches(sheet, table, p)))
                {
                    continue;
                }

                var fieldIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var field in table.Fields)
                {
                    fieldIndex[field.Name] = field.Index;
                    seenColumns.Add(field.Name);
                }

                if (!columnPredicates.All(p => fieldIndex.ContainsKey(p.Attribute)))
                {
                    continue;
                }

                foreach (var p in columnPredicates)
                {
                    matchedColumns.Add(p.Attribute);
                }

                foreach (var row in table.DataRange.Rows())
                {
                    var rowMatches =
                        columnPredicates.All(p => RowCellMatches(row, fieldIndex[p.Attribute], p)) &&
                        containsPredicates.All(p => RowContains(row, table.Fields.Count(), p.Text));
                    if (!rowMatches)
                    {
                        continue;
                    }

                    total++;
                    if (matches.Count >= MaxQueryMatches)
                    {
                        continue;
                    }

                    var rowNumber = row.WorksheetRow().RowNumber();
                    var tableAddress = table.RangeAddress;
                    var values = new Dictionary<string, object?>(StringComparer.Ordinal);
                    foreach (var field in table.Fields)
                    {
                        values[field.Name] = ExcelValues.ToJson(row.Cell(field.Index + 1).Value);
                    }

                    matches.Add(new
                    {
                        path = string.Create(
                            CultureInfo.InvariantCulture,
                            $"{ExcelPaths.SheetPath(sheet)}/{tableAddress.FirstAddress.ColumnLetter}{rowNumber}:{tableAddress.LastAddress.ColumnLetter}{rowNumber}"),
                        sheet = sheet.Name,
                        table = table.Name,
                        row = rowNumber,
                        values,
                    });
                }
            }
        }

        var unknown = columnPredicates
            .Select(p => p.Attribute)
            .Where(a => !matchedColumns.Contains(a))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (unknown.Count > 0)
        {
            var known = seenColumns.Order(StringComparer.OrdinalIgnoreCase).ToList();
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"No table has column(s): {string.Join(", ", unknown)}.",
                known.Count > 0
                    ? "Row selectors filter table rows by column name; available columns are listed as candidates."
                    : "Row selectors need a table; create one first with edit op {op:add, type:table, path:/Sheet1/A1:C10}.",
                candidates: known.Count > 0 ? known : ["(no tables in workbook)"]);
        }

        return (matches, total);
    }

    private static bool TableFilterMatches(IXLWorksheet sheet, IXLTable table, AttributePredicate attr)
    {
        var actual = attr.Attribute == "sheet" ? sheet.Name : table.Name;
        return attr.Op switch
        {
            SelectorOperator.Equals => string.Equals(actual, attr.Value, StringComparison.OrdinalIgnoreCase),
            SelectorOperator.NotEquals => !string.Equals(actual, attr.Value, StringComparison.OrdinalIgnoreCase),
            SelectorOperator.ContainsText => actual.Contains(attr.Value, StringComparison.OrdinalIgnoreCase),
            _ => false,
        };
    }

    private static bool RowCellMatches(IXLRangeRow row, int fieldIndex, AttributePredicate attr)
    {
        var cell = row.Cell(fieldIndex + 1);
        return IsExistenceTest(attr) ? !cell.Value.IsBlank : CompareValue(cell.Value, cell, attr);
    }

    private static bool RowContains(IXLRangeRow row, int fieldCount, string text)
    {
        for (var column = 1; column <= fieldCount; column++)
        {
            if (ExcelValues.SafeFormatted(row.Cell(column)).Contains(text, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static (List<object> Matches, int Total) QuerySheets(XLWorkbook workbook, Selector selector)
    {
        var matches = new List<object>();
        foreach (var sheet in workbook.Worksheets.OrderBy(ws => ws.Position))
        {
            var ok = selector.Predicates.All(p => p switch
            {
                ContainsPredicate c => sheet.Name.Contains(c.Text, StringComparison.OrdinalIgnoreCase),
                AttributePredicate { Attribute: "name" } a => a.Op switch
                {
                    SelectorOperator.Equals => IsExistenceTest(a) ||
                        string.Equals(sheet.Name, a.Value, StringComparison.OrdinalIgnoreCase),
                    SelectorOperator.NotEquals => !string.Equals(sheet.Name, a.Value, StringComparison.OrdinalIgnoreCase),
                    SelectorOperator.ContainsText => sheet.Name.Contains(a.Value, StringComparison.OrdinalIgnoreCase),
                    _ => false,
                },
                AttributePredicate a => throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown sheet attribute '{a.Attribute}' in selector.",
                    "Sheet selectors support only [name=...] and :contains('...').",
                    candidates: ["name"]),
                _ => false,
            });

            if (ok)
            {
                matches.Add(new
                {
                    path = ExcelPaths.SheetPath(sheet),
                    sheet = sheet.Name,
                    usedRange = sheet.RangeUsed()?.RangeAddress.ToString(),
                });
            }
        }

        return (matches, matches.Count);
    }
}
