using AIOffice.Core;

namespace AIOffice.Excel;

/// <summary>
/// The <see cref="IEmbedHost"/> half of the xlsx handler (M10). Embedded objects
/// are stored as <see cref="DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart"/>
/// children of the sheet's worksheet part, with display/source/anchor metadata in
/// a workbook custom property — see <see cref="ExcelEmbeds"/>. Listing and
/// extraction are read-only; add/remove flow through the normal edit ops.
/// </summary>
public sealed partial class ExcelHandler : IEmbedHost
{
    public IReadOnlyList<EmbeddedObject> ListEmbeds(CommandContext ctx)
    {
        var file = RequireFile(ctx, mustExist: true);
        return ExcelEmbeds.ListEmbeds(file);
    }

    public void ExtractEmbed(CommandContext ctx, string embedPath, string destPath)
    {
        var file = RequireFile(ctx, mustExist: true);
        using var workbook = OpenWorkbook(file);
        var target = ExcelPaths.Resolve(workbook, embedPath);
        if (target.Kind != ExcelTargetKind.Embed)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"'{embedPath}' is not an embedded object path.",
                "Address an embed like /Sheet1/embed[1]; run 'aioffice read --view embeds' to list them.");
        }

        ExcelEmbeds.Extract(file, target, destPath);
    }
}
