using AIOffice.Core;

namespace AIOffice.Excel;

/// <summary>
/// The <see cref="IDiffer"/> half of the xlsx handler (M8). The semantic compare
/// lives in <see cref="ExcelDiff"/>; this file wires it to the verb surface and
/// guards the format/precondition checks. Diffing is read-only — neither the
/// current file nor the baseline is mutated.
/// </summary>
public sealed partial class ExcelHandler : IDiffer
{
    public DiffResult Diff(CommandContext ctx, string baselineFile)
    {
        var current = RequireFile(ctx, mustExist: true);

        if (string.IsNullOrWhiteSpace(baselineFile))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "diff needs a baseline workbook to compare against.",
                "Pass the other .xlsx as the second file, e.g. aioffice diff new.xlsx old.xlsx.");
        }

        if (!File.Exists(baselineFile))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"Baseline file not found: {baselineFile}",
                "Check the path spelling; both files must exist and be .xlsx workbooks.");
        }

        // Format guard: the baseline must be the same kind (an .xlsx). A .docx /
        // .pptx baseline is invalid_args naming the mismatch (per the contract).
        var baselineExtension = Path.GetExtension(baselineFile);
        if (!string.Equals(baselineExtension, ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"diff compares two workbooks, but the baseline is '{baselineExtension}', not .xlsx.",
                "Diff like-for-like: both files must be .xlsx. Convert the baseline first, or diff the matching format.");
        }

        FileSizeGuard.Ensure(baselineFile); // file_too_large before the expensive open

        return ExcelDiff.Compare(baselineFile, current);
    }
}
