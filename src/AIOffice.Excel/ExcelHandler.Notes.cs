using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;

namespace AIOffice.Excel;

/// <summary>
/// Cell notes (Excel's classic comments), ClosedXML-native. The documented
/// addressing choice: notes are ALWAYS addressed through their cell — add with
/// <c>{op:add, type:note, path:/Sheet1/B2, props:{text:"…", author?:"…"}}</c>,
/// remove with <c>{op:remove, path:/Sheet1/B2, props:{target:"note"}}</c>, and
/// <c>get</c> on the cell reflects the note. There is no <c>/Sheet1/note[i]</c>
/// index form (cells are the stable address; note indices would shift).
/// <c>read --view structure</c> lists every noted cell per sheet.
/// </summary>
public sealed partial class ExcelHandler
{
    private static readonly IReadOnlyList<string> NoteProps = ["text", "author"];

    private static readonly IReadOnlyList<string> RemoveTargets = ["note"];

    private static object AddNote(ExcelTarget target, EditOp op, int index)
    {
        if (target.Kind != ExcelTargetKind.Cell)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add note targets a cell path like /Sheet1/B2, not '{op.Path}'.",
                "Use {op:add, type:note, path:/Sheet1/B2, props:{text:\"check this\"}}.");
        }

        var props = op.Props ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: add note needs props.",
            "Pass props like {\"text\":\"check this\"} (author is optional).");

        foreach (var (key, _) in props)
        {
            if (!NoteProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: unknown note prop '{key}'.",
                    "Supported note props: " + string.Join(", ", NoteProps) + ".",
                    candidates: NoteProps);
            }
        }

        var text = StringNoteProp(props, "text");
        if (string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add note needs a non-empty 'text' prop.",
                "Pass it as a string, e.g. {\"text\":\"check this\"}.");
        }

        var cell = target.Cell!;
        var path = ExcelPaths.CellPath(target.Sheet, cell.Address);
        if (cell.HasComment)
        {
            if (ExcelComments.IsShadow(cell.GetComment()))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: {cell.Address} carries a threaded comment, which excludes a plain note.",
                    "Reply to the thread instead ({op:add, type:reply, path:/Sheet/comment[@id=…]}); " +
                    "run 'aioffice read --view comments' to find its path.");
            }

            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: {cell.Address} already has a note.",
                "Remove it first ({op:remove, path:" + path + ", props:{target:\"note\"}}), then add the new one.");
        }

        var comment = cell.CreateComment();
        comment.AddText(text);
        var author = StringNoteProp(props, "author");
        if (!string.IsNullOrEmpty(author))
        {
            comment.Author = author;
        }

        return new
        {
            op = "add",
            type = "note",
            path,
            author = string.IsNullOrEmpty(comment.Author) ? null : comment.Author,
        };
    }

    /// <summary>Handles <c>remove</c> ops carrying a <c>target</c> prop (currently: "note").</summary>
    private static object RemoveCellTarget(ExcelTarget target, EditOp op, JsonNode targetNode, int index)
    {
        var requested = targetNode is JsonValue value && value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;
        if (!string.Equals(requested, "note", StringComparison.Ordinal))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: unknown remove target '{requested ?? targetNode.ToJsonString()}'.",
                "The only remove target is \"note\" ({op:remove, path:/Sheet1/B2, props:{target:\"note\"}}); " +
                "omit props to clear the cell's contents instead.",
                candidates: RemoveTargets);
        }

        if (target.Kind != ExcelTargetKind.Cell)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: remove note targets a cell path like /Sheet1/B2, not '{op.Path}'.",
                "Notes are addressed through their cell; run 'aioffice read --view structure' to list noted cells.");
        }

        var cell = target.Cell!;
        var path = ExcelPaths.CellPath(target.Sheet, cell.Address);
        if (!cell.HasComment)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: {cell.Address} has no note to remove.",
                "Run 'aioffice read --view structure' to list noted cells.");
        }

        if (ExcelComments.IsShadow(cell.GetComment()))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: {cell.Address} carries a threaded comment, not a note.",
                "Remove the whole thread via its comment path ({op:remove, path:/Sheet/comment[@id=…]}); " +
                "run 'aioffice read --view comments' to find it.");
        }

        cell.Clear(XLClearOptions.Comments);
        return new { op = "remove", path, removed = "note" };
    }

    /// <summary>
    /// The note block for cell get; null (omitted on the wire) when the cell
    /// has none. A threaded comment's legacy shadow is NOT a note — it shows
    /// under <c>comment</c> instead (M5 note-vs-comment distinction).
    /// </summary>
    private static object? NoteInfo(IXLCell cell) =>
        cell.HasComment && !ExcelComments.IsShadow(cell.GetComment())
            ? new
            {
                text = cell.GetComment().Text,
                author = string.IsNullOrEmpty(cell.GetComment().Author) ? null : cell.GetComment().Author,
            }
            : null;

    /// <summary>Every noted cell on a sheet, for read --view structure (comment shadows excluded).</summary>
    private static List<object> NoteList(IXLWorksheet sheet) =>
        [.. sheet
            .CellsUsed(XLCellsUsedOptions.All)
            .Where(c => c.HasComment && !ExcelComments.IsShadow(c.GetComment()))
            .Select(c => new
            {
                cell = c.Address.ToString(),
                path = ExcelPaths.CellPath(sheet, c.Address),
                author = string.IsNullOrEmpty(c.GetComment().Author) ? null : c.GetComment().Author,
                text = c.GetComment().Text,
            })];

    private static string? StringNoteProp(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;
}
