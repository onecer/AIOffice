using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// Embedded objects on a worksheet (M10): any file — a source .xlsx attached to
/// a report, a .pdf, a .zip — embedded as an <see cref="EmbeddedPackagePart"/>
/// child of the sheet's <c>WorksheetPart</c>. ClosedXML cannot author these, so
/// the create/extract/remove raw work runs in a post-save pass over the file
/// ClosedXML saved (MEASURED for ClosedXML 0.105: an EmbeddedPackagePart and the
/// custom document property below both survive its saves byte-identical, exactly
/// like charts).
///
/// The worksheet OLE form (<c>x:oleObject</c> + a paired VML control shape) is
/// the alternative the shared contract allows but is validator-fragile; an
/// embedded package part on the worksheet is validator-clean and round-trips the
/// payload bytes exactly, so that is what aioffice writes.
///
/// Metadata the part itself cannot carry — the display name, the original source
/// file name, and the anchor cell — lives in a workbook custom document property
/// (<see cref="RegistryProperty"/>), a JSON array keyed by the absolute part URI
/// (stable across ClosedXML saves). The part's content type is the media type and
/// the part stream length is the size, so those are never duplicated.
///
/// Addressing: <c>/Sheet1/embed[i]</c>, 1-based per sheet, ordered by part URI.
/// </summary>
internal static class ExcelEmbeds
{
    /// <summary>The custom-property key the embed registry is stored under.</summary>
    public const string RegistryProperty = "_aioffice_embeds";

    private static readonly IReadOnlyList<string> AddProps = ["src", "name", "anchor", "icon"];

    // ----- op-time validation & capture --------------------------------------

    /// <summary>One validated embed-add captured at op time; the raw write is deferred to the post-save pass.</summary>
    public sealed record AddSpec(string SheetName, byte[] Payload, string MediaType, string Name, string? Anchor);

    /// <summary>One validated embed removal (resolved part URI at op time).</summary>
    public sealed record RemoveSpec(string SheetName, string PartUri, string Name);

    /// <summary>
    /// Collects embed ops during an edit batch so indices in a multi-op batch
    /// stay consistent (mirrors <see cref="ChartOpBatch"/>). Projected per-sheet
    /// counts move as adds/removes are queued.
    /// </summary>
    public sealed class Batch
    {
        private readonly string _file;
        private Dictionary<string, int>? _projected;

        public Batch(string file) => _file = file;

        /// <summary>Add/remove specs in batch order.</summary>
        public List<object> Ops { get; } = [];

        public bool IsEmpty => Ops.Count == 0;

        /// <summary>Queues an add and returns the embed's projected 1-based index on its sheet.</summary>
        public int Add(AddSpec spec)
        {
            var counts = Projected();
            var next = counts.GetValueOrDefault(spec.SheetName) + 1;
            counts[spec.SheetName] = next;
            Ops.Add(spec);
            return next;
        }

        /// <summary>Queues a removal of the 1-based index on a sheet, validating against the projected state.</summary>
        public RemoveSpec Remove(string file, string sheetName, string sheetPath, int index, int opIndex)
        {
            var onSheet = ListOnSheet(file, sheetName);

            // Account for embeds already removed earlier in this batch.
            var removedUris = Ops.OfType<RemoveSpec>()
                .Where(r => string.Equals(r.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                .Select(r => r.PartUri)
                .ToHashSet(StringComparer.Ordinal);
            var live = onSheet.Where(e => !removedUris.Contains(e.PartUri)).ToList();

            if (index < 1 || index > live.Count)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidPath,
                    $"ops[{opIndex}]: no embed[{index}] on sheet '{sheetName}' ({live.Count} embed(s) exist).",
                    live.Count > 0
                        ? "Embed indices are 1-based per sheet; run 'aioffice read --view embeds' to list them."
                        : "This sheet has no embeds; add one with {op:add, type:embed, path:" + sheetPath +
                          ", props:{src:\"report.xlsx\"}}.",
                    candidates: live.Count > 0
                        ? [.. Enumerable.Range(1, live.Count).Select(i => $"{sheetPath}/embed[{i}]")]
                        : [sheetPath]);
            }

            var target = live[index - 1];
            var counts = Projected();
            counts[sheetName] = counts.GetValueOrDefault(sheetName) - 1;
            var spec = new RemoveSpec(sheetName, target.PartUri, target.Name);
            Ops.Add(spec);
            return spec;
        }

        private Dictionary<string, int> Projected()
        {
            if (_projected is null)
            {
                _projected = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var group in ReadAll(_file).GroupBy(e => e.SheetName, StringComparer.OrdinalIgnoreCase))
                {
                    _projected[group.Key] = group.Count();
                }
            }

            return _projected;
        }
    }

    /// <summary>
    /// Validates an <c>add embed</c> op against the live workbook and captures
    /// the payload bytes + metadata. The src path is ALWAYS sandbox-resolved
    /// before a byte is read.
    /// </summary>
    public static AddSpec ParseAdd(Workspace workspace, ExcelTarget target, EditOp op, int opIndex)
    {
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add embed targets a sheet path like /Sheet1; the file goes in props.src.",
                "Use {op:add, type:embed, path:/Sheet1, props:{src:\"report.xlsx\", name:\"Q3 source\"}}.");
        }

        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add embed needs props.",
                "Pass props like {\"src\":\"report.xlsx\"}.");
        }

        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown embed prop '{key}'.",
                    "Supported embed props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }

        var src = OptionalString(props, "src");
        if (string.IsNullOrWhiteSpace(src))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add embed needs the 'src' prop (a file inside the workspace).",
                "Pass e.g. {\"src\":\"data/report.xlsx\"}.");
        }

        // Sandbox first: a src that escapes the workspace must die here, never be read.
        var resolved = workspace.Resolve(src, mustExist: true);
        if (!File.Exists(resolved))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"ops[{opIndex}]: embed src is not a file: {src}",
                "Point src at a file inside the workspace.");
        }

        var payload = File.ReadAllBytes(resolved);
        var mediaType = MediaTypeSniffer.Sniff(resolved);
        var name = OptionalString(props, "name");
        if (string.IsNullOrWhiteSpace(name))
        {
            name = Path.GetFileName(resolved);
        }

        string? anchor = null;
        var anchorText = OptionalString(props, "anchor");
        if (!string.IsNullOrWhiteSpace(anchorText))
        {
            anchor = anchorText.ToUpperInvariant();
            if (!CellPattern.IsMatch(anchor))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: '{anchorText}' is not a usable anchor.",
                    "anchor is the cell the embed is associated with, e.g. E2; omit it to leave the embed unanchored.");
            }
        }

        return new AddSpec(target.Sheet.Name, payload, mediaType, name!.Trim(), anchor);
    }

    // ----- read-back ---------------------------------------------------------

    /// <summary>One embed as read back from the package + registry (sheet-scoped, 1-based per sheet).</summary>
    public sealed record Info(
        string SheetName,
        int Index,
        string Path,
        string PartUri,
        string Name,
        string MediaType,
        long Size,
        string? Anchor,
        string? Source);

    /// <summary>Every embed in the workbook, in (sheet position, part-URI) order, 1-based per sheet.</summary>
    public static List<Info> ReadAll(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        return ReadAll(document);
    }

    /// <summary>Every embed in an already-open document.</summary>
    public static List<Info> ReadAll(SpreadsheetDocument document)
    {
        var result = new List<Info>();
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return result;
        }

        var registry = LoadRegistry(document);
        foreach (var sheet in workbookPart.Workbook.Descendants<S.Sheet>())
        {
            if (sheet.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var sheetName = sheet.Name?.Value ?? string.Empty;
            var index = 0;
            foreach (var part in worksheetPart.EmbeddedPackageParts.OrderBy(p => p.Uri.ToString(), StringComparer.Ordinal))
            {
                index++;
                var uri = part.Uri.ToString();
                var entry = registry.GetValueOrDefault(uri);
                result.Add(new Info(
                    SheetName: sheetName,
                    Index: index,
                    Path: EmbedPath(sheetName, index),
                    PartUri: uri,
                    Name: entry?.Name is { Length: > 0 } registered
                        ? registered
                        : DefaultName(uri),
                    MediaType: part.ContentType,
                    Size: PartSize(part),
                    Anchor: entry?.Anchor,
                    Source: entry?.Source));
            }
        }

        return result;
    }

    /// <summary>Every embed on one sheet, 1-based in part-URI order.</summary>
    public static List<Info> ListOnSheet(string file, string sheetName) =>
        ReadAll(file)
            .Where(e => string.Equals(e.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            .ToList();

    /// <summary>The contract projection (path/name/mediaType/size/container) for every embed.</summary>
    public static IReadOnlyList<EmbeddedObject> ListEmbeds(string file) =>
        [.. ReadAll(file).Select(e => new EmbeddedObject(
            Path: e.Path,
            Name: e.Name,
            MediaType: e.MediaType,
            Size: e.Size,
            Container: $"/{ExcelPaths.QuoteSheet(e.SheetName)}"))];

    /// <summary>Resolves an addressed embed; throws <c>invalid_path</c> with candidates on a miss.</summary>
    public static Info Resolve(string file, ExcelTarget target)
    {
        var onSheet = ListOnSheet(file, target.Sheet.Name);
        var wanted = target.EmbedIndex!.Value;
        var info = onSheet.FirstOrDefault(e => e.Index == wanted);
        if (info is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No embed[{wanted}] on sheet '{target.Sheet.Name}' ({onSheet.Count} embed(s) exist).",
                onSheet.Count > 0
                    ? "Embed indices are 1-based per sheet; run 'aioffice read --view embeds' to list them."
                    : "This sheet has no embeds; add one with {op:add, type:embed, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + ", props:{src:\"report.xlsx\"}}.",
                candidates: onSheet.Count > 0
                    ? [.. onSheet.Select(e => e.Path)]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return info;
    }

    /// <summary>The metadata an addressed embed exposes through <c>get</c> — never the bytes.</summary>
    public static object Describe(Info info) => new
    {
        path = info.Path,
        kind = "embed",
        sheet = info.SheetName,
        name = info.Name,
        mediaType = info.MediaType,
        size = info.Size,
        anchor = info.Anchor,
        source = info.Source,
    };

    // ----- extract -----------------------------------------------------------

    /// <summary>
    /// Writes the addressed embed's payload to <paramref name="destPath"/>
    /// (already sandbox-resolved). Does NOT modify the source document; the
    /// returned canonical path lets the op echo a stable target.
    /// </summary>
    public static string Extract(string file, ExcelTarget target, string destPath)
    {
        var info = Resolve(file, target);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var part = FindPart(document, info.SheetName, info.PartUri)
            ?? throw new AiofficeException(
                ErrorCodes.InternalError,
                $"The embed at {info.Path} resolved but its package part is missing.",
                "Retry; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");

        var directory = Path.GetDirectoryName(destPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using (var source = part.GetStream())
        using (var dest = File.Create(destPath))
        {
            source.CopyTo(dest);
        }

        return info.Path;
    }

    // ----- post-save apply ---------------------------------------------------

    /// <summary>
    /// Applies queued embed adds/removes to the file ClosedXML just saved. All
    /// semantic validation already happened at op time, so this pass is
    /// mechanical (mirrors <see cref="ExcelCharts.Apply"/>).
    /// </summary>
    public static void Apply(string file, IReadOnlyList<object> ops)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var registry = LoadRegistry(document);

        foreach (var op in ops)
        {
            switch (op)
            {
                case AddSpec add:
                {
                    var worksheetPart = SheetPartOrThrow(document, add.SheetName);
                    var part = worksheetPart.AddNewPart<EmbeddedPackagePart>(add.MediaType);
                    using (var stream = new MemoryStream(add.Payload, writable: false))
                    {
                        part.FeedData(stream);
                    }

                    registry[part.Uri.ToString()] = new Entry(add.Name, add.Anchor, add.SheetName, add.Name);
                    break;
                }

                case RemoveSpec remove:
                {
                    var worksheetPart = SheetPartOrThrow(document, remove.SheetName);
                    var part = worksheetPart.EmbeddedPackageParts
                        .FirstOrDefault(p => string.Equals(p.Uri.ToString(), remove.PartUri, StringComparison.Ordinal));
                    if (part is not null)
                    {
                        worksheetPart.DeletePart(part);
                    }

                    registry.Remove(remove.PartUri);
                    break;
                }
            }
        }

        StoreRegistry(document, registry);
    }

    // ----- registry (custom-property JSON map) -------------------------------

    /// <summary>One registry entry: the metadata the part cannot carry itself.</summary>
    private sealed record Entry(string Name, string? Anchor, string? Sheet, string? Source);

    private static Dictionary<string, Entry> LoadRegistry(SpreadsheetDocument document)
    {
        var result = new Dictionary<string, Entry>(StringComparer.Ordinal);
        var raw = ReadCustomProperty(document, RegistryProperty);
        if (string.IsNullOrEmpty(raw))
        {
            return result;
        }

        JsonNode? parsed;
        try
        {
            parsed = JsonNode.Parse(raw);
        }
        catch (JsonException)
        {
            return result; // a corrupt registry never crashes a read; it reads as empty
        }

        if (parsed is not JsonArray array)
        {
            return result;
        }

        foreach (var node in array)
        {
            if (node is not JsonObject obj || GetString(obj, "uri") is not { } uri)
            {
                continue;
            }

            result[uri] = new Entry(
                Name: GetString(obj, "name") ?? DefaultName(uri),
                Anchor: GetString(obj, "anchor"),
                Sheet: GetString(obj, "sheet"),
                Source: GetString(obj, "source"));
        }

        return result;
    }

    private static void StoreRegistry(SpreadsheetDocument document, Dictionary<string, Entry> registry)
    {
        var array = new JsonArray();
        foreach (var (uri, entry) in registry.OrderBy(kv => kv.Key, StringComparer.Ordinal))
        {
            var obj = new JsonObject { ["uri"] = uri, ["name"] = entry.Name };
            if (entry.Anchor is not null)
            {
                obj["anchor"] = entry.Anchor;
            }

            if (entry.Sheet is not null)
            {
                obj["sheet"] = entry.Sheet;
            }

            if (entry.Source is not null)
            {
                obj["source"] = entry.Source;
            }

            array.Add(obj);
        }

        WriteCustomProperty(document, RegistryProperty, array.ToJsonString(JsonDefaults.Options));
    }

    // ----- raw helpers -------------------------------------------------------

    private static readonly System.Text.RegularExpressions.Regex CellPattern =
        new("^[A-Z]{1,3}[0-9]{1,7}$", System.Text.RegularExpressions.RegexOptions.Compiled);

    private static WorksheetPart SheetPartOrThrow(SpreadsheetDocument document, string sheetName)
    {
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is { } relationshipId &&
            workbookPart!.GetPartById(relationshipId) is WorksheetPart worksheetPart)
        {
            return worksheetPart;
        }

        throw new AiofficeException(
            ErrorCodes.InternalError,
            $"Sheet '{sheetName}' disappeared between validation and the embed write pass.",
            "Retry the edit; if it recurs, restore a snapshot with 'aioffice snapshot restore'.");
    }

    private static EmbeddedPackagePart? FindPart(SpreadsheetDocument document, string sheetName, string partUri)
    {
        var workbookPart = document.WorkbookPart;
        var sheet = workbookPart?.Workbook?.Descendants<S.Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet?.Id?.Value is not { } relationshipId ||
            workbookPart!.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
        {
            return null;
        }

        return worksheetPart.EmbeddedPackageParts
            .FirstOrDefault(p => string.Equals(p.Uri.ToString(), partUri, StringComparison.Ordinal));
    }

    private static long PartSize(EmbeddedPackagePart part)
    {
        using var stream = part.GetStream();
        return stream.Length;
    }

    private static string EmbedPath(string sheetName, int index) =>
        string.Create(CultureInfo.InvariantCulture, $"/{ExcelPaths.QuoteSheet(sheetName)}/embed[{index}]");

    private static string DefaultName(string partUri) => Path.GetFileName(partUri);

    private static string? OptionalString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static string? GetString(JsonObject obj, string key) =>
        obj.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String &&
        value.GetValue<string>() is { Length: > 0 } text
            ? text
            : null;

    // The custom document property is read/written raw so it lands in the same
    // docProps/custom.xml ClosedXML uses; ClosedXML preserves it across saves.
    private static string? ReadCustomProperty(SpreadsheetDocument document, string name)
    {
        var part = document.CustomFilePropertiesPart;
        if (part?.Properties is null)
        {
            return null;
        }

        foreach (var property in part.Properties
                     .Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
        {
            if (string.Equals(property.Name?.Value, name, StringComparison.Ordinal))
            {
                return property.VTLPWSTR?.Text ?? property.VTBString?.Text;
            }
        }

        return null;
    }

    private static void WriteCustomProperty(SpreadsheetDocument document, string name, string value)
    {
        var part = document.CustomFilePropertiesPart ?? document.AddCustomFilePropertiesPart();
        part.Properties ??= new DocumentFormat.OpenXml.CustomProperties.Properties();
        var properties = part.Properties;

        var existing = properties
            .Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>()
            .FirstOrDefault(p => string.Equals(p.Name?.Value, name, StringComparison.Ordinal));
        existing?.Remove();

        var nextId = properties
            .Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>()
            .Select(p => p.PropertyId?.Value ?? 1)
            .DefaultIfEmpty(1)
            .Max() + 1;

        properties.Append(new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty
        {
            FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
            PropertyId = Math.Max(2, nextId),
            Name = name,
            VTLPWSTR = new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR(value),
        });

        // Renumber sequentially so ids stay 2..N+1 (Office requires pid >= 2,
        // contiguous) after a delete + append.
        var pid = 2;
        foreach (var property in properties
                     .Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
        {
            property.PropertyId = pid++;
        }

        properties.Save();
    }
}

/// <summary>Sniffs a content type from a file's header bytes (extension is a fallback only).</summary>
internal static class MediaTypeSniffer
{
    /// <summary>The generic binary content type when no signature matches.</summary>
    public const string OctetStream = "application/octet-stream";

    /// <summary>Content type of an embedded source: header-sniffed, then by extension, then octet-stream.</summary>
    public static string Sniff(string file)
    {
        byte[] header;
        using (var stream = File.OpenRead(file))
        {
            header = new byte[Math.Min(8, stream.Length)];
            _ = stream.Read(header, 0, header.Length);
        }

        // ZIP-family (xlsx/docx/pptx/odt/jar/zip) all start "PK\x03\x04"; the OOXML
        // ones are distinguished by extension since the magic is identical.
        if (header.Length >= 4 && header[0] == 0x50 && header[1] == 0x4B && header[2] == 0x03 && header[3] == 0x04)
        {
            return ByExtension(file) ?? "application/zip";
        }

        if (header.Length >= 5 && header[0] == 0x25 && header[1] == 0x50 && header[2] == 0x44 && header[3] == 0x46 &&
            header[4] == 0x2D)
        {
            return "application/pdf"; // "%PDF-"
        }

        if (header.Length >= 8 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
        {
            return "image/png";
        }

        if (header.Length >= 3 && header[0] == 0xFF && header[1] == 0xD8 && header[2] == 0xFF)
        {
            return "image/jpeg";
        }

        if (header.Length >= 4 && header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0)
        {
            // Legacy OLE compound file (old .doc/.xls/.ppt): sniff the extension.
            return ByExtension(file) ?? OctetStream;
        }

        return ByExtension(file) ?? OctetStream;
    }

    private static string? ByExtension(string file) =>
        Path.GetExtension(file).ToLowerInvariant() switch
        {
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".pdf" => "application/pdf",
            ".zip" => "application/zip",
            ".csv" => "text/csv",
            ".txt" => "text/plain",
            ".json" => "application/json",
            ".xml" => "application/xml",
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".doc" => "application/msword",
            ".xls" => "application/vnd.ms-excel",
            ".ppt" => "application/vnd.ms-powerpoint",
            _ => null,
        };
}
