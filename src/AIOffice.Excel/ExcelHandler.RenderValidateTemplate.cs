using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace AIOffice.Excel;

public sealed partial class ExcelHandler
{
    // ----- render ------------------------------------------------------------

    public Envelope Render(CommandContext ctx) => Run(ctx, sw =>
    {
        var file = RequireFile(ctx, mustExist: true);
        var to = ArgString(ctx, "to") ?? "html";
        using var workbook = OpenWorkbook(file);
        var sections = ResolveTextSections(workbook, ArgString(ctx, "scope"));

        switch (to)
        {
            case "html":
                return Envelope.Ok(
                    new { format = "html", content = RenderHtml(sections) },
                    MetaFor(file, sw));

            case "text":
            {
                var (content, _) = BuildText(workbook, ArgString(ctx, "scope"), int.MaxValue);
                return Envelope.Ok(new { format = "text", content }, MetaFor(file, sw));
            }

            default:
                throw new AiofficeException(
                    ErrorCodes.UnsupportedFeature,
                    $"xlsx render supports html and text in v0; '{to}' lands in M1.",
                    "Render --to html and open it in a browser (or screenshot it) as a workaround.",
                    candidates: ["html", "text"]);
        }
    });

    /// <summary>
    /// Used-range tables with number formats applied (display text, not raw
    /// values). Merged cells render their value in the anchor cell; spanned
    /// cells appear empty in v0. Every cell carries
    /// <c>data-aio-path="/Sheet1/B2"</c> (and each table its sheet path) so a
    /// browser click maps back to a canonical document path.
    /// </summary>
    private static string RenderHtml(List<(IXLWorksheet Sheet, IXLRange? Range)> sections)
    {
        var sb = new StringBuilder();
        foreach (var (sheet, range) in sections)
        {
            sb.Append("<table data-sheet=\"");
            ExcelValues.AppendEscaped(sb, sheet.Name).Append("\" data-aio-path=\"");
            ExcelValues.AppendEscaped(sb, ExcelPaths.SheetPath(sheet)).Append("\">\n<caption>");
            ExcelValues.AppendEscaped(sb, sheet.Name).Append("</caption>\n");
            if (range is not null)
            {
                foreach (var row in range.Rows())
                {
                    sb.Append("<tr>");
                    foreach (var cell in row.Cells(true))
                    {
                        var bold = cell.Style.Font.Bold;
                        var italic = cell.Style.Font.Italic;
                        sb.Append("<td data-aio-path=\"");
                        ExcelValues.AppendEscaped(sb, ExcelPaths.CellPath(sheet, cell.Address));
                        sb.Append("\">");
                        if (bold)
                        {
                            sb.Append("<strong>");
                        }

                        if (italic)
                        {
                            sb.Append("<em>");
                        }

                        ExcelValues.AppendEscaped(sb, ExcelValues.SafeFormatted(cell));
                        if (italic)
                        {
                            sb.Append("</em>");
                        }

                        if (bold)
                        {
                            sb.Append("</strong>");
                        }

                        sb.Append("</td>");
                    }

                    sb.Append("</tr>\n");
                }
            }

            sb.Append("</table>\n");
        }

        return sb.ToString();
    }

    // ----- validate ----------------------------------------------------------

    public Envelope Validate(CommandContext ctx) => Run(ctx, sw =>
    {
        var file = RequireFile(ctx, mustExist: true);
        var issues = new List<object>();

        // Oracle 1: the OpenXml schema validator.
        int errorCount;
        try
        {
            using var document = SpreadsheetDocument.Open(file, isEditable: false);
            var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
            var errors = validator.Validate(document).ToList();
            errorCount = errors.Count;
            issues.AddRange(errors.Select(e => (object)new
            {
                severity = "error",
                code = e.Id,
                message = e.Description,
                part = e.Part?.Uri.ToString(),
            }));
        }
        catch (Exception exception) when (exception is not AiofficeException)
        {
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                $"The file is not a readable xlsx package: {exception.Message}",
                "Restore a snapshot ('aioffice snapshot list') or re-export the file from its source.",
                innerException: exception);
        }

        // Oracle 2: our lint — formulas that evaluate to errors.
        using var workbook = OpenWorkbook(file);
        foreach (var sheet in workbook.Worksheets)
        {
            foreach (var cell in sheet.CellsUsed().Where(c => c.HasFormula))
            {
                XLCellValue value;
                try
                {
                    value = cell.Value;
                }
                catch (Exception)
                {
                    continue;
                }

                if (!value.IsError)
                {
                    continue;
                }

                var path = ExcelPaths.CellPath(sheet, cell.Address);
                issues.Add(value.GetError() == XLError.NameNotRecognized
                    ? new
                    {
                        severity = "warning",
                        code = ErrorCodes.FormulaNotEvaluated,
                        message = $"{path}: ={cell.FormulaA1} uses a function the built-in engine cannot evaluate; Excel computes it on open.",
                        path,
                    }
                    : (object)new
                    {
                        severity = "warning",
                        code = "formula_error",
                        message = $"{path}: ={cell.FormulaA1} evaluates to {value}.",
                        path,
                    });
            }
        }

        return Envelope.Ok(
            new
            {
                valid = errorCount == 0,
                errors = errorCount,
                warnings = issues.Count - errorCount,
                issues,
            },
            MetaFor(file, sw));
    });

    // ----- template ----------------------------------------------------------

    [GeneratedRegex(@"\{\{\s*([A-Za-z0-9_.-]+)\s*\}\}")]
    private static partial Regex Placeholder();

    public Envelope Template(CommandContext ctx) => Run(ctx, sw =>
    {
        var file = RequireFile(ctx, mustExist: true);
        if (!ctx.Args.TryGetPropertyValue("data", out var dataNode) || dataNode is not JsonObject data)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "template needs a data object.",
                "Pass --data '{\"name\":\"Ann\",\"total\":42}' or --data @data.json.");
        }

        var outputArg = ArgString(ctx, "output");
        var outPath = outputArg is null ? file : ctx.Workspace.Resolve(outputArg);

        using var workbook = OpenWorkbook(file);
        var replacements = 0;
        var cellsTouched = 0;
        var unresolved = new SortedSet<string>(StringComparer.Ordinal);

        foreach (var sheet in workbook.Worksheets)
        {
            // Materialize first: writing typed values while iterating mutates the used set.
            var candidates = sheet.CellsUsed()
                .Where(c => !c.HasFormula && c.Value.IsText &&
                            c.Value.GetText().Contains("{{", StringComparison.Ordinal))
                .ToList();
            foreach (var cell in candidates)
            {
                var text = cell.Value.GetText();
                var matches = Placeholder().Matches(text);
                if (matches.Count == 0)
                {
                    continue;
                }

                // A cell that is exactly one placeholder takes the JSON value typed
                // (numbers stay numbers, formulas become formulas).
                if (matches.Count == 1 && matches[0].Length == text.Length)
                {
                    var key = matches[0].Groups[1].Value;
                    if (!data.TryGetPropertyValue(key, out var valueNode))
                    {
                        unresolved.Add(key);
                        continue;
                    }

                    WriteParsed(cell, ExcelValues.Parse(valueNode));
                    replacements++;
                    cellsTouched++;
                    continue;
                }

                var changed = false;
                var newText = Placeholder().Replace(text, match =>
                {
                    var key = match.Groups[1].Value;
                    if (data.TryGetPropertyValue(key, out var valueNode))
                    {
                        changed = true;
                        replacements++;
                        return ExcelValues.TemplateText(valueNode);
                    }

                    unresolved.Add(key);
                    return match.Value;
                });

                if (changed)
                {
                    cell.Value = newText;
                    cellsTouched++;
                }
            }
        }

        if (string.Equals(outPath, file, StringComparison.Ordinal))
        {
            _snapshots.Save(file);
        }

        var warnings = SaveWithCachedValues(workbook, outPath) ?? [];
        if (unresolved.Count > 0)
        {
            warnings.Add(new Warning(
                "template_unresolved",
                $"Placeholder(s) without data were left as-is: {string.Join(", ", unresolved)}. " +
                "Add them to --data to fill them."));
        }

        return Envelope.Ok(
            new
            {
                file = outPath,
                cells = cellsTouched,
                replacements,
                unresolved = unresolved.Count > 0 ? unresolved.ToList() : null,
            },
            MetaFor(outPath, sw, warnings));
    });
}
