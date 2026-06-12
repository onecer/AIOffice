using System.Globalization;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using TC = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace AIOffice.Excel;

/// <summary>
/// Modern THREADED comments (M5) — distinct from M4 plain notes. Storage is
/// the 2018 threadedComments format: one <c>threadedComments</c> part per
/// sheet plus a workbook-level <c>persons</c> part (authors deduplicated by
/// display name). Authored on raw OpenXml in a post-save pass, because
/// ClosedXML cannot see these parts — MEASURED for ClosedXML 0.105: once
/// present, both parts survive its saves byte-identical.
///
/// Old-Excel fallback: every thread also writes the standard legacy note
/// shadow through ClosedXML (author <c>tc={thread id}</c>, the MS-XLSX
/// pairing convention), so pre-2019 Excel shows the conversation as a note
/// while modern Excel hides the shadow behind the real thread. Note surfaces
/// (get, structure) treat shadows as comments, never as notes.
///
/// Addressing: threads are created on a cell
/// (<c>{op:add, type:comment, path:/Sheet1/B2}</c>) and afterwards addressed
/// by the stable id form <c>/Sheet1/comment[@id=GUID]</c> (replies target it,
/// remove deletes the whole thread). <c>read --view comments</c> lists every
/// thread; <c>get</c> on the cell reports the thread under <c>comment</c>.
/// </summary>
internal static class ExcelComments
{
    public static readonly IReadOnlyList<string> CommentProps = ["text", "author"];

    /// <summary>The author used when an op does not name one.</summary>
    public const string DefaultAuthor = "AIOffice";

    private const string ShadowAuthorPrefix = "tc=";

    private const string ShadowPreamble =
        "[Threaded comment]\n\n" +
        "Your version of Excel allows you to read this threaded comment; however, any edits to it " +
        "will get removed if the file is opened in a newer version of Excel. " +
        "Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n";

    // ----- model ---------------------------------------------------------------

    /// <summary>One message (the thread root and replies share this shape).</summary>
    public sealed record Message(string Id, string PersonId, string Text, DateTime At);

    /// <summary>One comment thread anchored at a cell.</summary>
    public sealed class Thread
    {
        public required string SheetName { get; init; }

        public required string CellRef { get; set; }

        public required Message Root { get; set; }

        public List<Message> Replies { get; } = [];
    }

    /// <summary>A deduplicated author from the persons part.</summary>
    public sealed record Person(string Id, string DisplayName);

    /// <summary>
    /// The whole file's threaded-comment state, loaded once per edit batch and
    /// written back after the ClosedXML save.
    /// </summary>
    public sealed class Model
    {
        public List<Thread> Threads { get; } = [];

        public List<Person> Persons { get; } = [];

        /// <summary>Set by every mutating op; gates the post-save write.</summary>
        public bool Dirty { get; set; }

        public Thread? FindByCell(string sheetName, string cellRef) =>
            Threads.FirstOrDefault(t =>
                string.Equals(t.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(t.CellRef, cellRef, StringComparison.OrdinalIgnoreCase));

        public Thread? FindById(string sheetName, string bareId) =>
            Threads.FirstOrDefault(t =>
                string.Equals(t.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(BareId(t.Root.Id), bareId, StringComparison.OrdinalIgnoreCase));

        public Person PersonFor(string? displayName)
        {
            var name = string.IsNullOrEmpty(displayName) ? DefaultAuthor : displayName;
            var existing = Persons.FirstOrDefault(p => string.Equals(p.DisplayName, name, StringComparison.Ordinal));
            if (existing is not null)
            {
                return existing;
            }

            var person = new Person(NewId(), name);
            Persons.Add(person);
            return person;
        }

        public string DisplayNameOf(string personId) =>
            Persons.FirstOrDefault(p => string.Equals(p.Id, personId, StringComparison.OrdinalIgnoreCase))
                ?.DisplayName ?? DefaultAuthor;
    }

    /// <summary>Braced uppercase GUID, the on-disk id form.</summary>
    public static string NewId() => "{" + Guid.NewGuid().ToString("D").ToUpperInvariant() + "}";

    /// <summary>The wire id form (no braces) used in <c>comment[@id=…]</c> paths.</summary>
    public static string BareId(string storedId) => storedId.Trim('{', '}');

    // ----- load ------------------------------------------------------------------

    /// <summary>Reads the file's threaded comments and persons (empty model when none exist).</summary>
    public static Model Load(string file)
    {
        var model = new Model();
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var workbookPart = document.WorkbookPart;
        if (workbookPart is null)
        {
            return model;
        }

        foreach (var person in workbookPart.GetPartsOfType<WorkbookPersonPart>()
                     .SelectMany(p => p.PersonList?.Elements<TC.Person>() ?? []))
        {
            if (person.Id?.Value is { } id)
            {
                model.Persons.Add(new Person(id, person.DisplayName?.Value ?? DefaultAuthor));
            }
        }

        foreach (var sheetElement in workbookPart.Workbook?.Descendants<S.Sheet>() ?? [])
        {
            if (sheetElement.Name?.Value is not { } sheetName ||
                sheetElement.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var byId = new Dictionary<string, Thread>(StringComparer.OrdinalIgnoreCase);
            foreach (var comment in worksheetPart.GetPartsOfType<WorksheetThreadedCommentsPart>()
                         .SelectMany(p => p.ThreadedComments?.Elements<TC.ThreadedComment>() ?? []))
            {
                if (comment.Id?.Value is not { } id || comment.Ref?.Value is not { } cellRef)
                {
                    continue;
                }

                var message = new Message(
                    id,
                    comment.PersonId?.Value ?? string.Empty,
                    comment.GetFirstChild<TC.ThreadedCommentText>()?.Text ?? string.Empty,
                    comment.DT?.Value ?? default);

                if (comment.ParentId?.Value is { } parentId)
                {
                    if (byId.TryGetValue(parentId, out var parent))
                    {
                        parent.Replies.Add(message);
                    }

                    continue; // orphaned replies are dropped defensively
                }

                var thread = new Thread { SheetName = sheetName, CellRef = cellRef, Root = message };
                byId[id] = thread;
                model.Threads.Add(thread);
            }
        }

        return model;
    }

    // ----- shadow notes ------------------------------------------------------------

    /// <summary>True when a ClosedXML comment is the legacy shadow of a thread, not a real note.</summary>
    public static bool IsShadow(IXLComment comment) =>
        comment.Author?.StartsWith(ShadowAuthorPrefix, StringComparison.Ordinal) == true;

    /// <summary>True when the cell carries a threaded comment's legacy shadow.</summary>
    public static bool HasShadow(IXLCell cell) => cell.HasComment && IsShadow(cell.GetComment());

    /// <summary>Rewrites the cell's legacy shadow note to mirror the thread's current state.</summary>
    public static void WriteShadow(IXLCell cell, Thread thread)
    {
        if (cell.HasComment)
        {
            cell.Clear(XLClearOptions.Comments);
        }

        var shadow = cell.CreateComment();
        shadow.AddText(ShadowText(thread));
        shadow.Author = ShadowAuthorPrefix + thread.Root.Id;
    }

    private static string ShadowText(Thread thread)
    {
        var text = ShadowPreamble + thread.Root.Text;
        foreach (var reply in thread.Replies)
        {
            text += "\nReply:\n" + reply.Text;
        }

        return text;
    }

    // ----- describe ------------------------------------------------------------------

    /// <summary>The canonical thread path: <c>/Sheet1/comment[@id=GUID]</c>.</summary>
    public static string PathOf(IXLWorksheet sheet, Thread thread) =>
        $"{ExcelPaths.SheetPath(sheet)}/comment[@id={BareId(thread.Root.Id)}]";

    /// <summary>One thread as agents see it (get, read --view comments).</summary>
    public static object Describe(IXLWorksheet sheet, Thread thread, Model model) => new
    {
        path = PathOf(sheet, thread),
        kind = "comment",
        sheet = sheet.Name,
        cell = thread.CellRef,
        cellPath = $"{ExcelPaths.SheetPath(sheet)}/{thread.CellRef}",
        author = model.DisplayNameOf(thread.Root.PersonId),
        text = thread.Root.Text,
        at = AtText(thread.Root.At),
        replies = thread.Replies
            .Select(reply => new
            {
                author = model.DisplayNameOf(reply.PersonId),
                text = reply.Text,
                at = AtText(reply.At),
            })
            .ToList(),
    };

    private static string? AtText(DateTime at) =>
        at == default ? null : at.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);

    /// <summary>The sheet's thread paths, nearest first is not meaningful here — file order.</summary>
    public static IReadOnlyList<string> CandidatesOn(IXLWorksheet sheet, Model model) =>
        [.. model.Threads
            .Where(t => string.Equals(t.SheetName, sheet.Name, StringComparison.OrdinalIgnoreCase))
            .Select(t => PathOf(sheet, t))];

    // ----- write-back -------------------------------------------------------------------

    /// <summary>
    /// Rewrites the threadedComments and persons parts from the model after the
    /// ClosedXML save (which preserved any existing ones byte-identical).
    /// </summary>
    public static void WriteAfterSave(string file, Model model)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: true);
        var workbookPart = document.WorkbookPart ?? throw new AiofficeException(
            ErrorCodes.InternalError,
            "Threaded-comment write failed: the package has no workbook part.",
            "The edit was rolled back; re-run it. If this persists, report a bug with the workbook.");

        WritePersons(workbookPart, model);

        foreach (var sheetElement in workbookPart.Workbook?.Descendants<S.Sheet>().ToList() ?? [])
        {
            if (sheetElement.Name?.Value is not { } sheetName ||
                sheetElement.Id?.Value is not { } relationshipId ||
                workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var threads = model.Threads
                .Where(t => string.Equals(t.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                .ToList();
            var existing = worksheetPart.GetPartsOfType<WorksheetThreadedCommentsPart>().FirstOrDefault();

            if (threads.Count == 0)
            {
                if (existing is not null)
                {
                    worksheetPart.DeletePart(existing);
                }

                continue;
            }

            var part = existing ?? worksheetPart.AddNewPart<WorksheetThreadedCommentsPart>();
            var root = new TC.ThreadedComments();
            foreach (var thread in threads)
            {
                root.Append(ToElement(thread.Root, thread.CellRef, parentId: null));
                foreach (var reply in thread.Replies)
                {
                    root.Append(ToElement(reply, thread.CellRef, parentId: thread.Root.Id));
                }
            }

            part.ThreadedComments = root;
            part.ThreadedComments.Save();
        }
    }

    private static void WritePersons(WorkbookPart workbookPart, Model model)
    {
        var existing = workbookPart.GetPartsOfType<WorkbookPersonPart>().FirstOrDefault();
        if (model.Threads.Count == 0)
        {
            if (existing is not null)
            {
                workbookPart.DeletePart(existing);
            }

            return;
        }

        var referenced = model.Threads
            .SelectMany(t => t.Replies.Select(r => r.PersonId).Prepend(t.Root.PersonId))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var part = existing ?? workbookPart.AddNewPart<WorkbookPersonPart>();
        var list = new TC.PersonList();
        foreach (var person in model.Persons.Where(p => referenced.Contains(p.Id)))
        {
            list.Append(new TC.Person
            {
                DisplayName = person.DisplayName,
                Id = person.Id,
                UserId = person.DisplayName,
                ProviderId = "None",
            });
        }

        part.PersonList = list;
        part.PersonList.Save();
    }

    private static TC.ThreadedComment ToElement(Message message, string cellRef, string? parentId)
    {
        var element = new TC.ThreadedComment(new TC.ThreadedCommentText(message.Text))
        {
            Ref = cellRef,
            Id = message.Id,
            PersonId = message.PersonId,
        };
        if (message.At != default)
        {
            element.DT = new DateTimeValue(message.At);
        }

        if (parentId is not null)
        {
            element.ParentId = parentId;
        }

        return element;
    }
}
