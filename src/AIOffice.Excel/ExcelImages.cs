using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace AIOffice.Excel;

/// <summary>
/// The xlsx image layer. Adds go through ClosedXML (measured: new picture
/// anchors merge correctly into existing drawings parts). Removals do NOT:
/// ClosedXML 0.105's <c>Pictures.Delete</c> + save deletes the media part and
/// renumbers the drawing relationships but leaves every picture anchor in the
/// drawing XML, producing dangling <c>r:embed</c> references. So removals
/// leave the ClosedXML model untouched and a raw post-save pass
/// (<see cref="RemoveAfterSave"/>) deletes the anchor and its now-orphaned
/// image part instead. PNG and JPEG only, decided by sniffing the file header
/// — never by extension. The source path is ALWAYS resolved through the Core
/// <see cref="Workspace"/> sandbox.
/// </summary>
internal static partial class ExcelImages
{
    /// <summary>Image formats aioffice embeds (header-sniffed).</summary>
    public static readonly IReadOnlyList<string> Formats = ["png", "jpeg"];

    private static readonly IReadOnlyList<string> AddProps = ["src", "anchor", "widthPx", "heightPx", "name"];

    [GeneratedRegex("^([A-Z]{1,3})([0-9]{1,7})$")]
    private static partial Regex CellPattern();

    // ----- add ---------------------------------------------------------------

    /// <summary>
    /// Validates and applies an <c>add image</c> op on the in-memory workbook.
    /// Returns the details entry for the envelope. <paramref name="pendingRemoved"/>
    /// holds picture names queued for raw removal earlier in the same batch.
    /// </summary>
    public static object Add(
        Workspace workspace, ExcelTarget target, EditOp op, int opIndex, IReadOnlySet<string> pendingRemoved)
    {
        if (target.Kind != ExcelTargetKind.Sheet)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add image targets a sheet path like /Sheet1; the position goes in props.anchor.",
                "Use {op:add, type:image, path:/Sheet1, props:{src:\"logo.png\", anchor:\"E2\"}}.");
        }

        var props = op.Props;
        if (props is null || props.Count == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: add image needs props.",
                "Pass props like {\"src\":\"logo.png\",\"anchor\":\"E2\"}.");
        }

        foreach (var (key, _) in props)
        {
            if (!AddProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{opIndex}]: unknown image prop '{key}'.",
                    "Supported image props: " + string.Join(", ", AddProps) + ".",
                    candidates: AddProps);
            }
        }

        var src = OptionalString(props, "src") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add image needs the 'src' prop (path of a PNG or JPEG inside the workspace).",
            "Pass e.g. {\"src\":\"assets/logo.png\",\"anchor\":\"E2\"}.");

        // Sandbox first: a src that escapes the workspace must die here, never be read.
        var resolved = workspace.Resolve(src, mustExist: true);
        if (!File.Exists(resolved))
        {
            throw new AiofficeException(
                ErrorCodes.FileNotFound,
                $"ops[{opIndex}]: image src is not a file: {src}",
                "Point src at a .png or .jpeg file inside the workspace.");
        }

        var bytes = File.ReadAllBytes(resolved);
        var format = Sniff(bytes, src, opIndex);

        var anchorText = OptionalString(props, "anchor") ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: add image needs the 'anchor' prop (top-left cell).",
            "Pass e.g. {\"anchor\":\"E2\"}.");
        var anchor = anchorText.ToUpperInvariant();
        if (!CellPattern().IsMatch(anchor))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: '{anchorText}' is not a usable anchor.",
                "anchor is the top-left cell the image hangs from, e.g. E2.");
        }

        var width = OptionalInt(props, "widthPx", opIndex);
        var height = OptionalInt(props, "heightPx", opIndex);
        if (width is < 1 || height is < 1)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: widthPx and heightPx must be at least 1.",
                "Omit them to keep the image's natural pixel size.");
        }

        var name = OptionalString(props, "name");
        if (name is not null && pendingRemoved.Contains(name))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: image '{name}' is being removed in this batch; its name frees up only after the save.",
                "Run the remove and the re-add as two separate edit batches, or pick another name.");
        }

        IXLPicture picture;
        try
        {
            using var stream = new MemoryStream(bytes);
            picture = name is null
                ? target.Sheet.AddPicture(stream, format)
                : target.Sheet.AddPicture(stream, format, name);
        }
        catch (Exception exception) when (exception is ArgumentException or InvalidOperationException)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{opIndex}]: could not embed the image: {exception.Message}",
                "Check that the file is a real PNG/JPEG and any explicit name is unused on the sheet.",
                innerException: exception);
        }

        picture.MoveTo(target.Sheet.Cell(anchor));
        Resize(picture, width, height);

        var index = target.Sheet.Pictures.Count(p => !pendingRemoved.Contains(p.Name));
        return new
        {
            op = "add",
            type = "image",
            path = ExcelPaths.ImagePath(target.Sheet, index),
            name = picture.Name,
            format = FormatName(picture.Format),
            anchor,
            widthPx = picture.Width,
            heightPx = picture.Height,
        };
    }

    /// <summary>Explicit sizes win; a single dimension keeps the aspect ratio; neither keeps natural size.</summary>
    private static void Resize(IXLPicture picture, int? width, int? height)
    {
        switch (width, height)
        {
            case ({ } w, { } h):
                picture.WithSize(w, h);
                break;
            case ({ } w, null):
                picture.WithSize(w, Math.Max(1, (int)Math.Round(
                    (double)picture.OriginalHeight * w / picture.OriginalWidth)));
                break;
            case (null, { } h):
                picture.WithSize(Math.Max(1, (int)Math.Round(
                    (double)picture.OriginalWidth * h / picture.OriginalHeight)), h);
                break;
            default:
                break; // natural size
        }
    }

    /// <summary>PNG/JPEG by magic bytes; anything else is a typed unsupported_feature.</summary>
    private static XLPictureFormat Sniff(byte[] bytes, string src, int opIndex)
    {
        if (bytes.Length >= 8 &&
            bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47 &&
            bytes[4] == 0x0D && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A)
        {
            return XLPictureFormat.Png;
        }

        if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
        {
            return XLPictureFormat.Jpeg;
        }

        throw new AiofficeException(
            ErrorCodes.UnsupportedFeature,
            $"ops[{opIndex}]: '{src}' is not a PNG or JPEG (sniffed by file header, not extension).",
            "Only png and jpeg images can be embedded; convert the file first (e.g. with sips or ImageMagick).",
            candidates: Formats);
    }

    // ----- find / describe -----------------------------------------------------

    /// <summary>
    /// Finds the picture a resolved target addresses (1-based sheet order).
    /// Throws <c>invalid_path</c> with the sheet's actual image paths as candidates.
    /// </summary>
    public static IXLPicture Find(ExcelTarget target) =>
        FindForEdit(target, new HashSet<string>(StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// <see cref="Find"/> against the batch-projected state: pictures whose
    /// raw removal is already queued no longer count.
    /// </summary>
    public static IXLPicture FindForEdit(ExcelTarget target, IReadOnlySet<string> pendingRemoved)
    {
        var pictures = target.Sheet.Pictures.Where(p => !pendingRemoved.Contains(p.Name)).ToList();
        var index = target.ImageIndex!.Value;
        if (index > pictures.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No image[{index}] on sheet '{target.Sheet.Name}' ({pictures.Count} image(s) exist).",
                pictures.Count > 0
                    ? "Image indices are 1-based per sheet; run 'aioffice read --view structure' to list them."
                    : "This sheet has no images; add one with {op:add, type:image, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + ", props:{src:\"logo.png\", anchor:\"E2\"}}.",
                candidates: pictures.Count > 0
                    ? [.. Enumerable.Range(1, pictures.Count).Select(i => ExcelPaths.ImagePath(target.Sheet, i))]
                    : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return pictures[index - 1];
    }

    // ----- post-save raw removal -------------------------------------------------

    /// <summary>
    /// Deletes queued picture anchors (matched by the picture name in
    /// <c>xdr:cNvPr</c>) plus their image parts when nothing references them
    /// any more, and drops a drawings part that ends up empty — mirroring the
    /// chart removal clean-up.
    /// </summary>
    public static void RemoveAfterSave(string file, IReadOnlyList<(string Sheet, string Name)> removals)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart?.Workbook is null)
        {
            return;
        }

        foreach (var group in removals.GroupBy(r => r.Sheet, StringComparer.OrdinalIgnoreCase))
        {
            var sheet = workbookPart.Workbook.Descendants<S.Sheet>()
                .FirstOrDefault(s => string.Equals(s.Name?.Value, group.Key, StringComparison.OrdinalIgnoreCase));
            if (sheet?.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart ||
                worksheetPart.DrawingsPart is not { } drawings ||
                drawings.WorksheetDrawing is not { } root)
            {
                continue;
            }

            var wanted = group.Select(g => g.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var anchor in root.ChildElements
                         .OfType<OpenXmlCompositeElement>()
                         .Where(a => a is Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor)
                         .ToList())
            {
                if (anchor.Descendants<Xdr.Picture>().FirstOrDefault() is not { } picture)
                {
                    continue; // a chart or shape anchor — never ours to touch here
                }

                var pictureName = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
                if (pictureName is null || !wanted.Contains(pictureName))
                {
                    continue;
                }

                var embedId = picture.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                    .FirstOrDefault()?.Embed?.Value;
                anchor.Remove();
                if (embedId is not null &&
                    !root.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Any(b => b.Embed?.Value == embedId) &&
                    drawings.Parts.Any(p => p.RelationshipId == embedId))
                {
                    drawings.DeletePart(embedId);
                }
            }

            var anchorsLeft = root.ChildElements
                .Any(c => c is Xdr.TwoCellAnchor or Xdr.OneCellAnchor or Xdr.AbsoluteAnchor);
            if (anchorsLeft)
            {
                root.Save();
                continue;
            }

            var drawingsId = worksheetPart.GetIdOfPart(drawings);
            var drawingElements = worksheetPart.Worksheet?.Elements<S.Drawing>()
                .Where(d => d.Id?.Value == drawingsId).ToList() ?? [];
            foreach (var drawing in drawingElements)
            {
                drawing.Remove();
            }

            worksheetPart.DeletePart(drawings);
            worksheetPart.Worksheet?.Save();
        }
    }

    /// <summary>One image as agents see it (get and read --view structure).</summary>
    public static object Describe(IXLWorksheet sheet, IXLPicture picture, int index) => new
    {
        path = ExcelPaths.ImagePath(sheet, index),
        kind = "image",
        sheet = sheet.Name,
        name = picture.Name,
        format = FormatName(picture.Format),
        anchor = picture.TopLeftCell.Address.ToString(),
        widthPx = picture.Width,
        heightPx = picture.Height,
    };

    private static string FormatName(XLPictureFormat format) => format switch
    {
        XLPictureFormat.Png => "png",
        XLPictureFormat.Jpeg => "jpeg",
        _ => format.ToString().ToLowerInvariant(),
    };

    private static string? OptionalString(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    private static int? OptionalInt(JsonObject props, string key, int opIndex)
    {
        if (!props.TryGetPropertyValue(key, out var node) || node is null)
        {
            return null;
        }

        if (node is JsonValue value)
        {
            if (value.GetValueKind() == JsonValueKind.Number && value.TryGetValue<int>(out var number))
            {
                return number;
            }

            if (value.GetValueKind() == JsonValueKind.String &&
                int.TryParse(value.GetValue<string>(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
            {
                return parsed;
            }
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{opIndex}]: '{key}' must be a whole number of pixels.",
            $"Pass e.g. {{\"{key}\":240}}.");
    }
}
