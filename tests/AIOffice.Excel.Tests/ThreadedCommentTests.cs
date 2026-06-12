using System.IO.Compression;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using TC = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M5 threaded comments (the modern 2018 format) — distinct from M4 notes.
/// Oracles: the raw threadedComments + persons parts (read with the OpenXml
/// SDK), the legacy shadow note pairing (<c>tc={id}</c> author, the MS-XLSX
/// convention), the schema validator, and part-survival across later saves.
/// </summary>
public sealed partial class ThreadedCommentTests : ExcelTestBase
{
    [GeneratedRegex(@"comment\[@id=(?<id>[^\]]+)\]")]
    private static partial Regex CommentIdInPath();

    private static EditOp CommentOp(string path, string text, string? author = null)
    {
        var props = new JsonObject { ["text"] = text };
        if (author is not null)
        {
            props["author"] = author;
        }

        return new EditOp { Op = "add", Path = path, Type = "comment", Props = props };
    }

    private static EditOp ReplyOp(string commentPath, string text, string? author = null)
    {
        var props = new JsonObject { ["text"] = text };
        if (author is not null)
        {
            props["author"] = author;
        }

        return new EditOp { Op = "add", Path = commentPath, Type = "reply", Props = props };
    }

    private string AddThread(string file, string cellPath, string text, string? author = null)
    {
        var envelope = EditOps(file, CommentOp(cellPath, text, author));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return Json(envelope)["data"]!["ops"]![0]!["path"]!.GetValue<string>();
    }

    private static List<TC.ThreadedComment> RawThreads(string file, int sheetIndex = 0)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var part = document.WorkbookPart!.WorksheetParts.ElementAt(sheetIndex)
            .GetPartsOfType<WorksheetThreadedCommentsPart>()
            .FirstOrDefault();
        return part?.ThreadedComments?.Elements<TC.ThreadedComment>()
            .Select(c => (TC.ThreadedComment)c.CloneNode(true))
            .ToList() ?? [];
    }

    private static List<TC.Person> RawPersons(string file)
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var part = document.WorkbookPart!.GetPartsOfType<WorkbookPersonPart>().FirstOrDefault();
        return part?.PersonList?.Elements<TC.Person>().Select(p => (TC.Person)p.CloneNode(true)).ToList() ?? [];
    }

    [Fact]
    public void Add_comment_writes_threaded_parts_persons_and_a_legacy_shadow()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/B2", ("value", 42))).IsOk);

        var path = AddThread(file, "/Sheet1/B2", "looks wrong", "Reviewer");
        Assert.Matches(CommentIdInPath(), path);

        var threads = RawThreads(file);
        var root = Assert.Single(threads);
        Assert.Equal("B2", root.Ref!.Value);
        Assert.Equal("looks wrong", root.GetFirstChild<TC.ThreadedCommentText>()!.Text);
        Assert.Null(root.ParentId);
        Assert.NotNull(root.DT);

        var person = Assert.Single(RawPersons(file));
        Assert.Equal("Reviewer", person.DisplayName!.Value);
        Assert.Equal(person.Id!.Value, root.PersonId!.Value);

        // The legacy shadow: an old-Excel note paired by the tc={id} author.
        using (var workbook = new XLWorkbook(file))
        {
            var b2 = workbook.Worksheet("Sheet1").Cell("B2");
            Assert.True(b2.HasComment);
            Assert.Equal("tc=" + root.Id!.Value, b2.GetComment().Author);
            Assert.Contains("looks wrong", b2.GetComment().Text, StringComparison.Ordinal);
            Assert.StartsWith("[Threaded comment]", b2.GetComment().Text, StringComparison.Ordinal);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Get_on_the_cell_reports_a_comment_not_a_note()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/B2", ("value", 42))).IsOk);
        var path = AddThread(file, "/Sheet1/B2", "check the total", "Reviewer");

        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Equal(42.0, cell["value"]!.GetValue<double>());
        Assert.Null(cell["note"]); // the shadow is NOT a note
        Assert.Equal("check the total", cell["comment"]!["text"]!.GetValue<string>());
        Assert.Equal("Reviewer", cell["comment"]!["author"]!.GetValue<string>());
        Assert.Equal(path, cell["comment"]!["path"]!.GetValue<string>());

        // …and structure's notes list skips shadows too.
        var structure = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        Assert.Empty(structure["sheets"]![0]!["notes"]!.AsArray());

        // get on the comment path itself works.
        var thread = OkData(Handler.Get(Ctx(file, ("path", path))));
        Assert.Equal("comment", thread["kind"]!.GetValue<string>());
        Assert.Equal("B2", thread["cell"]!.GetValue<string>());
    }

    [Fact]
    public void Replies_extend_the_thread_and_dedupe_persons()
    {
        var file = CreateWorkbook();
        var path = AddThread(file, "/Sheet1/C3", "is this right?", "Ann");

        Assert.True(EditOps(file, ReplyOp(path, "yes, verified", "Bob")).IsOk);
        var again = EditOps(file, ReplyOp(path, "thanks!", "Ann")); // same author as the root
        Assert.True(again.IsOk, again.ToJson());
        Assert.Equal(2, Json(again)["data"]!["ops"]![0]!["replies"]!.GetValue<int>());

        var threads = RawThreads(file);
        Assert.Equal(3, threads.Count); // root + 2 replies, flat in the part
        var root = threads.Single(t => t.ParentId is null);
        Assert.All(threads.Where(t => t.ParentId is not null), r => Assert.Equal(root.Id!.Value, r.ParentId!.Value));

        var persons = RawPersons(file);
        Assert.Equal(2, persons.Count); // Ann deduplicated
        Assert.Contains(persons, p => p.DisplayName!.Value == "Ann");
        Assert.Contains(persons, p => p.DisplayName!.Value == "Bob");

        // The shadow mirrors the whole conversation for old Excel.
        using (var workbook = new XLWorkbook(file))
        {
            var shadow = workbook.Worksheet("Sheet1").Cell("C3").GetComment().Text;
            Assert.Contains("is this right?", shadow, StringComparison.Ordinal);
            Assert.Contains("Reply:\nyes, verified", shadow, StringComparison.Ordinal);
            Assert.Contains("Reply:\nthanks!", shadow, StringComparison.Ordinal);
        }

        var view = OkData(Handler.Read(Ctx(file, ("view", "comments"))));
        Assert.Equal(1, view["count"]!.GetValue<int>());
        var listed = view["threads"]![0]!;
        Assert.Equal("C3", listed["cell"]!.GetValue<string>());
        Assert.Equal("Ann", listed["author"]!.GetValue<string>());
        var replies = listed["replies"]!.AsArray();
        Assert.Equal(2, replies.Count);
        Assert.Equal("Bob", replies[0]!["author"]!.GetValue<string>());
        Assert.Equal("yes, verified", replies[0]!["text"]!.GetValue<string>());

        AssertValidatorClean(file);
    }

    [Fact]
    public void Remove_deletes_the_whole_thread_and_cleans_the_parts()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/B2", ("value", "keep me"))).IsOk);
        var path = AddThread(file, "/Sheet1/B2", "temp thread");
        Assert.True(EditOps(file, ReplyOp(path, "and a reply")).IsOk);

        var removed = EditOps(file, RemoveOp(path));
        Assert.True(removed.IsOk, removed.ToJson());
        var detail = Json(removed)["data"]!["ops"]![0]!;
        Assert.Equal("comment", detail["removed"]!.GetValue<string>());
        Assert.Equal(1, detail["replies"]!.GetValue<int>());

        Assert.Empty(RawThreads(file));
        Assert.Empty(RawPersons(file)); // last thread gone -> parts removed
        using (var zip = ZipFile.OpenRead(file))
        {
            Assert.DoesNotContain(zip.Entries, e => e.FullName.Contains("threadedcomment", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(zip.Entries, e => e.FullName.Contains("persons", StringComparison.OrdinalIgnoreCase));
        }

        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Equal("keep me", cell["value"]!.GetValue<string>());
        Assert.Null(cell["comment"]);
        Assert.Null(cell["note"]); // the shadow went with the thread

        AssertValidatorClean(file);
    }

    [Fact]
    public void Threads_survive_later_unrelated_edits_byte_identical()
    {
        var file = CreateWorkbook();
        var path = AddThread(file, "/Sheet1/B2", "sticky thread", "Reviewer");

        static Dictionary<string, byte[]> ThreadParts(string f)
        {
            using var zip = ZipFile.OpenRead(f);
            return zip.Entries
                .Where(e => e.FullName.Contains("threadedcomment", StringComparison.OrdinalIgnoreCase) ||
                            e.FullName.Contains("persons", StringComparison.OrdinalIgnoreCase))
                .ToDictionary(e => e.FullName, e =>
                {
                    using var s = e.Open();
                    using var m = new MemoryStream();
                    s.CopyTo(m);
                    return m.ToArray();
                }, StringComparer.Ordinal);
        }

        var before = ThreadParts(file);
        Assert.Equal(2, before.Count); // threadedComments + persons

        // A later, unrelated edit batch must not disturb the parts.
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "later"))).IsOk);

        var after = ThreadParts(file);
        Assert.Equal(before.Keys.Order(StringComparer.Ordinal), after.Keys.Order(StringComparer.Ordinal));
        foreach (var (name, bytes) in before)
        {
            Assert.True(bytes.AsSpan().SequenceEqual(after[name]), name + " changed across an unrelated edit");
        }

        var thread = OkData(Handler.Get(Ctx(file, ("path", path))));
        Assert.Equal("sticky thread", thread["text"]!.GetValue<string>());
        AssertValidatorClean(file);
    }

    [Fact]
    public void Notes_and_comments_exclude_each_other_on_one_cell()
    {
        var file = CreateWorkbook();
        var path = AddThread(file, "/Sheet1/B2", "thread here");

        // note onto a comment cell
        var noteOnComment = EditOps(file, new EditOp
        {
            Op = "add", Path = "/Sheet1/B2", Type = "note", Props = new JsonObject { ["text"] = "a note" },
        });
        Assert.False(noteOnComment.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, noteOnComment.Error!.Code);
        Assert.Contains("threaded comment", noteOnComment.Error.Message, StringComparison.Ordinal);

        // remove "note" on a comment cell points at the thread path
        var removeNote = EditOps(file, new EditOp
        {
            Op = "remove", Path = "/Sheet1/B2", Props = new JsonObject { ["target"] = "note" },
        });
        Assert.False(removeNote.IsOk);
        Assert.Contains("threaded comment", removeNote.Error!.Message, StringComparison.Ordinal);

        // a second thread on the same cell suggests replying instead
        var secondThread = EditOps(file, CommentOp("/Sheet1/B2", "another?"));
        Assert.False(secondThread.IsOk);
        Assert.Contains("reply", secondThread.Error!.Suggestion, StringComparison.Ordinal);
        Assert.Contains(path.Split('/')[^1], secondThread.Error.Suggestion, StringComparison.Ordinal);

        // comment onto a noted cell
        Assert.True(EditOps(file, new EditOp
        {
            Op = "add", Path = "/Sheet1/D4", Type = "note", Props = new JsonObject { ["text"] = "plain note" },
        }).IsOk);
        var commentOnNote = EditOps(file, CommentOp("/Sheet1/D4", "thread?"));
        Assert.False(commentOnNote.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, commentOnNote.Error!.Code);
        Assert.Contains("note", commentOnNote.Error.Message, StringComparison.Ordinal);

        // the plain note still reads as a note
        var noted = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D4"))));
        Assert.Equal("plain note", noted["note"]!["text"]!.GetValue<string>());
        Assert.Null(noted["comment"]);
    }

    [Fact]
    public void Missing_thread_ids_fail_with_candidates()
    {
        var file = CreateWorkbook();
        var path = AddThread(file, "/Sheet1/B2", "only thread");

        var bogus = "/Sheet1/comment[@id=00000000-0000-0000-0000-000000000000]";
        var reply = EditOps(file, ReplyOp(bogus, "into the void"));
        Assert.False(reply.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, reply.Error!.Code);
        Assert.Contains(path, reply.Error.Candidates!);

        var got = Handler.Get(Ctx(file, ("path", bogus)));
        Assert.False(got.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, got.Error!.Code);
        Assert.Contains(path, got.Error.Candidates!);

        // reply must target a comment path, not a cell
        var replyOnCell = EditOps(file, ReplyOp("/Sheet1/B2", "wrong address"));
        Assert.False(replyOnCell.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, replyOnCell.Error!.Code);
    }

    [Fact]
    public void Default_author_and_multi_sheet_listing_work()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, AddOp("/Other", "sheet")).IsOk);
        AddThread(file, "/Sheet1/A1", "no author given");
        AddThread(file, "/Other/B2", "second sheet", "Ann");

        var view = OkData(Handler.Read(Ctx(file, ("view", "comments"))));
        Assert.Equal(2, view["count"]!.GetValue<int>());
        var threads = view["threads"]!.AsArray();
        Assert.Equal("AIOffice", threads[0]!["author"]!.GetValue<string>()); // documented default author
        Assert.Equal("Sheet1", threads[0]!["sheet"]!.GetValue<string>());
        Assert.Equal("Other", threads[1]!["sheet"]!.GetValue<string>());

        AssertValidatorClean(file);
    }

    /// <summary>
    /// The threaded-comment corollary of the round-trip law. MEASURED for
    /// ClosedXML 0.105.0: with threadedComments + persons parts present, a
    /// no-edit open/save keeps BOTH byte-identical (and the legacy shadow's
    /// comments part too, per the M4 note corollary).
    /// </summary>
    [Fact]
    public void NoEdit_resave_keeps_threaded_parts_byte_identical()
    {
        var file = CreateWorkbook();
        var path = AddThread(file, "/Sheet1/B2", "round trip me", "Reviewer");
        Assert.True(EditOps(file, ReplyOp(path, "still here")).IsOk);

        var resaved = Path.Combine(Dir, "resaved.xlsx");
        File.Copy(file, resaved);
        using (var workbook = new XLWorkbook(resaved))
        {
            workbook.Save(); // no edits at all
        }

        static Dictionary<string, byte[]> Parts(string f)
        {
            using var zip = ZipFile.OpenRead(f);
            return zip.Entries.ToDictionary(e => e.FullName, e =>
            {
                using var s = e.Open();
                using var m = new MemoryStream();
                s.CopyTo(m);
                return m.ToArray();
            }, StringComparer.Ordinal);
        }

        var before = Parts(file);
        var after = Parts(resaved);
        Assert.Equal(before.Keys.Order(StringComparer.Ordinal), after.Keys.Order(StringComparer.Ordinal));
        Assert.Contains(before.Keys, k => k.Contains("threadedcomment", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(before.Keys, k => k.Contains("persons", StringComparison.OrdinalIgnoreCase));

        var changed = before.Keys
            .Where(name => !before[name].AsSpan().SequenceEqual(after[name]))
            .ToHashSet(StringComparer.Ordinal);
        Assert.Equal(["xl/worksheets/sheet1.xml"], changed); // exactly the documented part

        AssertValidatorClean(resaved);
    }
}
