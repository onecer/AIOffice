using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

/// <summary>
/// The .docx format handler: implements the whole aioffice verb surface on
/// DocumentFormat.OpenXml. Edits are atomic (applied to an in-memory copy,
/// written back only on success) and every successful in-place write snapshots
/// the pre-image first.
/// </summary>
public sealed partial class WordHandler : IFormatHandler
{
    private readonly SnapshotStore _snapshots;

    /// <param name="snapshots">Override for tests; defaults to the user-level snapshot ring.</param>
    public WordHandler(SnapshotStore? snapshots = null) => _snapshots = snapshots ?? new SnapshotStore();

    public DocumentKind Kind => DocumentKind.Docx;

    // ------------------------------------------------------------------ create

    public Envelope Create(CommandContext ctx)
    {
        var file = RequireFile(ctx);
        if (File.Exists(file))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"File already exists: {file}",
                "Pick a new file name, or edit the existing document with 'aioffice edit'.");
        }

        var title = StringArg(ctx.Args, "title");

        byte[] bytes;
        using (var ms = new MemoryStream())
        {
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body());
                WordFactory.AddDefaultStylesPart(main);

                var body = main.Document.Body!;
                if (title is { Length: > 0 })
                {
                    body.AppendChild(WordFactory.Paragraph(title, styleId: "Heading1"));
                }

                body.AppendChild(new Paragraph());
            }

            bytes = ms.ToArray();
        }

        var dir = Path.GetDirectoryName(file);
        if (dir is { Length: > 0 })
        {
            Directory.CreateDirectory(dir);
        }

        File.WriteAllBytes(file, bytes);

        return Envelope.Ok(
            new { created = file, kind = "docx", title },
            MetaFor(file, Rev.OfBytes(bytes)));
    }

    // ---------------------------------------------------------------- validate

    public Envelope Validate(CommandContext ctx)
    {
        var file = RequireFile(ctx, mustExist: true);
        var bytes = File.ReadAllBytes(file);

        using var ms = new MemoryStream(bytes, writable: false);
        using var doc = OpenPackage(ms, file, editable: false);

        var validator = new OpenXmlValidator();
        var issues = validator
            .Validate(doc)
            .Select(e => new
            {
                severity = e.ErrorType.ToString(),
                id = e.Id,
                description = e.Description,
                part = e.Part?.Uri.ToString(),
                xpath = e.Path?.XPath,
            })
            .ToList();

        return Envelope.Ok(
            new { valid = issues.Count == 0, count = issues.Count, issues },
            MetaFor(file, Rev.OfBytes(bytes)));
    }

    // ----------------------------------------------------------------- helpers

    /// <summary>Sandbox-resolved target file; the CLI resolves, the handler re-checks.</summary>
    private static string RequireFile(CommandContext ctx, bool mustExist = false)
    {
        if (string.IsNullOrWhiteSpace(ctx.File))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "This command requires a target file.",
                "Pass the document path, e.g. 'aioffice read report.docx'.");
        }

        var file = ctx.Workspace.Resolve(ctx.File, mustExist);
        if (mustExist && !File.Exists(file))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"File not found: {ctx.File}",
                "Check the path spelling, or run 'aioffice create' to make a new document.");
        }

        FileSizeGuard.Ensure(file); // file_too_large before any expensive open
        return file;
    }

    /// <summary>Opens a package, mapping zip/xml failures to <c>format_corrupt</c>.</summary>
    private static WordprocessingDocument OpenPackage(Stream stream, string file, bool editable)
    {
        try
        {
            return WordprocessingDocument.Open(stream, editable);
        }
        catch (Exception ex) when (
            ex is OpenXmlPackageException or InvalidDataException or System.Xml.XmlException or FileFormatException)
        {
            // FileFormatException: System.IO.Packaging's own "not a package"
            // signal (surfaced by huge non-zip files since the M3 size-cap
            // default became unlimited) — same honest format_corrupt mapping.
            throw new AiofficeException(
                ErrorCodes.FormatCorrupt,
                $"Not a readable .docx package: {file} ({ex.Message})",
                "The file is corrupt or not OOXML. Re-export it from Word, or restore a snapshot with 'aioffice snapshot restore'.",
                innerException: ex);
        }
    }

    /// <summary>Loads the file into an expandable in-memory copy and opens it.</summary>
    private static (WordprocessingDocument Doc, MemoryStream Stream, byte[] OriginalBytes) OpenCopy(string file, bool editable)
    {
        var bytes = File.ReadAllBytes(file);
        var ms = new MemoryStream();
        ms.Write(bytes);
        ms.Position = 0;
        return (OpenPackage(ms, file, editable), ms, bytes);
    }

    private static Body GetBody(WordprocessingDocument doc, string file)
    {
        var body = doc.MainDocumentPart?.Document?.Body;
        return body ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            $"The document has no body: {file}",
            "The main document part is missing or empty. Re-export the file from Word.");
    }

    private static Meta MetaFor(string file, string rev, IReadOnlyList<Warning>? warnings = null) =>
        new() { File = file, Rev = rev, Warnings = warnings is { Count: > 0 } ? warnings : null };

    // ------------------------------------------------------------ args parsing

    private static string? StringArg(JsonObject args, string name)
    {
        var node = args[name];
        if (node is null)
        {
            return null;
        }

        if (node is JsonValue v)
        {
            if (v.TryGetValue<string>(out var s))
            {
                return s;
            }

            if (v.TryGetValue<bool>(out var b))
            {
                return b ? "true" : "false";
            }

            if (v.TryGetValue<double>(out var d))
            {
                return d.ToString(CultureInfo.InvariantCulture);
            }
        }

        return node.ToJsonString();
    }

    private static bool BoolArg(JsonObject args, string name)
    {
        var s = StringArg(args, name);
        return s is not null && (s.Equals("true", StringComparison.OrdinalIgnoreCase) || s == "1");
    }

    private static int? IntArg(JsonObject args, string name)
    {
        var s = StringArg(args, name);
        if (s is null)
        {
            return null;
        }

        if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"--{name} must be an integer, got '{s}'.",
                $"Pass a whole number, e.g. --{name} 10.");
        }

        return n;
    }

    /// <summary>A w:t with xml:space=preserve when whitespace is significant.</summary>
    internal static Text NewText(string value)
    {
        var t = new Text(value);
        if (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])))
        {
            t.Space = SpaceProcessingModeValues.Preserve;
        }

        return t;
    }
}
