using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;

namespace AIOffice.Excel;

public sealed partial class ExcelHandler
{
    /// <summary>
    /// <c>{op:add, type:comment, path:/Sheet1/B2, props:{text:"…", author?:"…"}}</c> —
    /// starts a THREADED comment (M5) on a cell. One thread per cell; replies
    /// target the thread (<c>comment[@id=…]</c>), and the legacy note shadow is
    /// mirrored through ClosedXML for old Excel.
    /// </summary>
    private static object AddComment(CommandContext ctx, ExcelTarget target, EditOp op, int index, PostSaveWork post)
    {
        if (target.Kind != ExcelTargetKind.Cell)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add comment targets a cell path like /Sheet1/B2, not '{op.Path}'.",
                "Use {op:add, type:comment, path:/Sheet1/B2, props:{text:\"…\"}}; " +
                "to answer an existing thread use type:reply on its comment[@id=…] path.");
        }

        var (text, author) = CommentTextAndAuthor(op, index, "comment");
        var cell = target.Cell!;
        var cellRef = cell.Address.ToString()!;
        var model = CommentModelFor(ctx, post);

        if (model.FindByCell(target.Sheet.Name, cellRef) is { } existing)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: {cellRef} already has a comment thread.",
                "Reply to it ({op:add, type:reply, path:" + ExcelComments.PathOf(target.Sheet, existing) +
                ", props:{text:\"…\"}}) or remove the thread first.");
        }

        if (cell.HasComment && !ExcelComments.IsShadow(cell.GetComment()))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: {cellRef} has a plain note; a cell cannot carry both a note and a comment thread.",
                "Remove the note first ({op:remove, path:" + ExcelPaths.CellPath(target.Sheet, cell.Address) +
                ", props:{target:\"note\"}}), then add the comment.");
        }

        var person = model.PersonFor(author);
        var thread = new ExcelComments.Thread
        {
            SheetName = target.Sheet.Name,
            CellRef = cellRef,
            Root = new ExcelComments.Message(ExcelComments.NewId(), person.Id, text, DateTime.UtcNow),
        };
        model.Threads.Add(thread);
        model.Dirty = true;

        ExcelComments.WriteShadow(cell, thread);

        return new
        {
            op = "add",
            type = "comment",
            path = ExcelComments.PathOf(target.Sheet, thread),
            cell = ExcelPaths.CellPath(target.Sheet, cell.Address),
            author = person.DisplayName,
        };
    }

    /// <summary>
    /// <c>{op:add, type:reply, path:/Sheet1/comment[@id=…], props:{text:"…"}}</c> —
    /// appends to an existing thread (the shadow note is rebuilt to mirror it).
    /// </summary>
    private static object AddReply(CommandContext ctx, ExcelTarget target, EditOp op, int index, PostSaveWork post)
    {
        if (target.Kind != ExcelTargetKind.Comment)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add reply targets a comment path like /Sheet1/comment[@id=GUID], not '{op.Path}'.",
                "Run 'aioffice read --view comments' to list thread paths; " +
                "to start a new thread use type:comment on the cell.");
        }

        var (text, author) = CommentTextAndAuthor(op, index, "reply");
        var model = CommentModelFor(ctx, post);
        var thread = FindThread(model, target, index);
        var person = model.PersonFor(author);
        thread.Replies.Add(new ExcelComments.Message(ExcelComments.NewId(), person.Id, text, DateTime.UtcNow));
        model.Dirty = true;

        ExcelComments.WriteShadow(target.Sheet.Cell(thread.CellRef), thread);

        return new
        {
            op = "add",
            type = "reply",
            path = ExcelComments.PathOf(target.Sheet, thread),
            cell = $"{ExcelPaths.SheetPath(target.Sheet)}/{thread.CellRef}",
            author = person.DisplayName,
            replies = thread.Replies.Count,
        };
    }

    /// <summary><c>{op:remove, path:/Sheet1/comment[@id=…]}</c> — deletes the WHOLE thread (replies included).</summary>
    private static object RemoveComment(CommandContext ctx, ExcelTarget target, int index, PostSaveWork post)
    {
        var model = CommentModelFor(ctx, post);
        var thread = FindThread(model, target, index);
        model.Threads.Remove(thread);
        model.Dirty = true;

        var cell = target.Sheet.Cell(thread.CellRef);
        if (ExcelComments.HasShadow(cell))
        {
            cell.Clear(ClosedXML.Excel.XLClearOptions.Comments);
        }

        return new
        {
            op = "remove",
            path = ExcelComments.PathOf(target.Sheet, thread),
            removed = "comment",
            cell = $"{ExcelPaths.SheetPath(target.Sheet)}/{thread.CellRef}",
            replies = thread.Replies.Count,
        };
    }

    private static ExcelComments.Thread FindThread(ExcelComments.Model model, ExcelTarget target, int index)
    {
        var thread = model.FindById(target.Sheet.Name, target.CommentId!);
        if (thread is null)
        {
            var candidates = ExcelComments.CandidatesOn(target.Sheet, model);
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"ops[{index}]: no comment thread with id '{target.CommentId}' on sheet '{target.Sheet.Name}'.",
                candidates.Count > 0
                    ? "Run 'aioffice read --view comments' to list thread paths; pick one of the candidates."
                    : "This sheet has no comment threads; start one with {op:add, type:comment, path:" +
                      ExcelPaths.SheetPath(target.Sheet) + "/B2, props:{text:\"…\"}}.",
                candidates: candidates.Count > 0 ? candidates : [ExcelPaths.SheetPath(target.Sheet)]);
        }

        return thread;
    }

    private static (string Text, string? Author) CommentTextAndAuthor(EditOp op, int index, string what)
    {
        var props = op.Props ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"ops[{index}]: add {what} needs props.",
            "Pass props like {\"text\":\"looks wrong\"} (author is optional).");

        foreach (var (key, _) in props)
        {
            if (!ExcelComments.CommentProps.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"ops[{index}]: unknown {what} prop '{key}'.",
                    "Supported props: " + string.Join(", ", ExcelComments.CommentProps) + ".",
                    candidates: ExcelComments.CommentProps);
            }
        }

        var text = StringProp(props, "text");
        if (string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"ops[{index}]: add {what} needs a non-empty 'text' prop.",
                "Pass it as a string, e.g. {\"text\":\"looks wrong\"}.");
        }

        return (text, StringProp(props, "author"));
    }

    private static string? StringProp(JsonObject props, string key) =>
        props.TryGetPropertyValue(key, out var node) && node is JsonValue value &&
        value.GetValueKind() == JsonValueKind.String
            ? value.GetValue<string>()
            : null;

    // ----- read --view comments ---------------------------------------------------

    /// <summary>Every comment thread in the workbook (cell, author, text, replies) with canonical paths.</summary>
    private Envelope ReadComments(
        ClosedXML.Excel.XLWorkbook workbook, string file, System.Diagnostics.Stopwatch sw)
    {
        var model = ExcelComments.Load(file);
        var threads = new List<object>();
        foreach (var sheet in workbook.Worksheets.OrderBy(ws => ws.Position))
        {
            threads.AddRange(model.Threads
                .Where(t => string.Equals(t.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase))
                .Select(t => ExcelComments.Describe(sheet, t, model)));
        }

        return Envelope.Ok(
            new { view = "comments", count = threads.Count, threads },
            MetaFor(file, sw));
    }
}
