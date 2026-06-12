using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// The M5 markdown/csv bridge wiring, shared by the CLI (<c>aioffice create
/// --from</c>, <c>aioffice read --view markdown|csv</c>) and the MCP tools
/// (<c>office_create.from</c>, <c>office_read.view</c>). Routing lives here
/// exactly once: the source EXTENSION picks the importer, the import matrix is
/// enforced before any handler runs, and the source path goes through the
/// workspace sandbox first — same as every other file-path-valued argument.
/// </summary>
public static class Bridge
{
    /// <summary>The create --from compatibility matrix, verbatim for suggestions.</summary>
    public const string ImportMatrix = ".md/.markdown → .docx, .csv/.tsv → .xlsx";

    private static readonly string[] MarkdownExtensions = [".md", ".markdown"];
    private static readonly string[] CsvExtensions = [".csv", ".tsv"];

    /// <summary>Read views that only exist on one format (view → owning kind).</summary>
    private static readonly IReadOnlyDictionary<string, DocumentKind> BridgeViews =
        new Dictionary<string, DocumentKind>(StringComparer.Ordinal)
        {
            ["markdown"] = DocumentKind.Docx,
            ["csv"] = DocumentKind.Xlsx,
        };

    private static readonly IReadOnlyDictionary<DocumentKind, string> ViewsByKind =
        new Dictionary<DocumentKind, string>
        {
            [DocumentKind.Docx] = "text, outline, stats, structure, revisions, comments, styles, markdown",
            [DocumentKind.Xlsx] = "outline, text, stats, structure, csv, comments",
            [DocumentKind.Pptx] = "outline, text, stats, structure, comments",
        };

    /// <summary>
    /// Routes <c>create --from &lt;source&gt;</c>: validates the source/target
    /// pair against the import matrix, sandbox-resolves the source, then hands
    /// off to the format's <see cref="IFormatHandler.CreateFrom"/>.
    /// </summary>
    /// <param name="handler">Handler already resolved for the TARGET file.</param>
    /// <param name="ctx">Command context; <c>ctx.File</c> is the resolved target.</param>
    /// <param name="sourcePath">User-supplied source path (resolved here).</param>
    public static Envelope CreateFrom(IFormatHandler handler, CommandContext ctx, string sourcePath)
    {
        if (string.IsNullOrWhiteSpace(sourcePath))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "--from needs a source file path.",
                $"Pass the source document, e.g. create report.docx --from notes.md. Matrix: {ImportMatrix}.");
        }

        var extension = Path.GetExtension(sourcePath);
        var required = RequiredKind(extension);
        if (required is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"create --from does not understand '{(extension.Length == 0 ? sourcePath : extension)}' sources.",
                $"Supported imports: {ImportMatrix}.",
                candidates: [.. MarkdownExtensions, .. CsvExtensions]);
        }

        if (handler.Kind != required)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"A '{extension}' source imports into a .{required.Value.ToString().ToLowerInvariant()} target, " +
                $"not .{handler.Kind.ToString().ToLowerInvariant()}.",
                $"Match the pair from the import matrix: {ImportMatrix}.");
        }

        // SECURITY: the source is a file-path-valued argument — sandbox-resolve
        // it at the command layer before any handler code touches it.
        var resolvedSource = ctx.Workspace.Resolve(sourcePath, mustExist: true);
        return handler.CreateFrom(ctx, resolvedSource);
    }

    /// <summary>
    /// Guards the cross-format bridge views: asking for <c>markdown</c> on a
    /// workbook (or <c>csv</c> on anything but one) is an unsupported FEATURE
    /// of that format, reported with the views it does have. Unknown view names
    /// stay the handlers' own <c>invalid_args</c> with candidates.
    /// </summary>
    public static void GuardBridgeView(DocumentKind kind, string? view)
    {
        if (view is null || !BridgeViews.TryGetValue(view, out var owner) || owner == kind)
        {
            return;
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"View '{view}' is a {owner.ToString().ToLowerInvariant()}-only view; " +
            $"this file is {kind.ToString().ToLowerInvariant()}.",
            $"Valid {kind.ToString().ToLowerInvariant()} views: {ViewsByKind[kind]}.",
            candidates: [.. ViewsByKind[kind].Split(", ")]);
    }

    private static DocumentKind? RequiredKind(string extension)
    {
        if (MarkdownExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            return DocumentKind.Docx;
        }

        if (CsvExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
        {
            return DocumentKind.Xlsx;
        }

        return null;
    }
}
