using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// M4 cell notes. The documented addressing: notes live on their CELL — add
/// via {op:add, type:note, path:/Sheet1/B2}, remove via {op:remove,
/// path:/Sheet1/B2, props:{target:"note"}}, reflect via get on the cell and
/// read --view structure (noted-cell list). No /Sheet1/note[i] index form.
/// </summary>
public sealed class NoteTests : ExcelTestBase
{
    private static EditOp NoteOp(string path, string text, string? author = null)
    {
        var props = new JsonObject { ["text"] = text };
        if (author is not null)
        {
            props["author"] = author;
        }

        return new EditOp { Op = "add", Path = path, Type = "note", Props = props };
    }

    private static EditOp RemoveNoteOp(string path) => new()
    {
        Op = "remove",
        Path = path,
        Props = new JsonObject { ["target"] = "note" },
    };

    [Fact]
    public void Add_note_reflects_in_get_and_survives_a_reopen()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/B2", ("value", 42))).IsOk);

        var envelope = EditOps(file, NoteOp("/Sheet1/B2", "check this", "Reviewer"));
        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal("/Sheet1/B2", detail["path"]!.GetValue<string>());
        Assert.Equal("Reviewer", detail["author"]!.GetValue<string>());

        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Equal(42.0, cell["value"]!.GetValue<double>()); // the value is untouched
        Assert.Equal("check this", cell["note"]!["text"]!.GetValue<string>());
        Assert.Equal("Reviewer", cell["note"]!["author"]!.GetValue<string>());

        // Reopen-verify with ClosedXML directly (not our own get).
        using (var workbook = new XLWorkbook(file))
        {
            var b2 = workbook.Worksheet("Sheet1").Cell("B2");
            Assert.True(b2.HasComment);
            Assert.Equal("check this", b2.GetComment().Text);
            Assert.Equal("Reviewer", b2.GetComment().Author);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Structure_lists_noted_cells()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "data")),
            NoteOp("/Sheet1/A1", "first"),
            NoteOp("/Sheet1/C3", "second")).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        var notes = data["sheets"]![0]!["notes"]!.AsArray();

        Assert.Equal(2, notes.Count);
        var cells = notes.Select(n => n!["cell"]!.GetValue<string>()).ToList();
        Assert.Contains("A1", cells);
        Assert.Contains("C3", cells);
        var first = notes.Single(n => n!["cell"]!.GetValue<string>() == "A1")!;
        Assert.Equal("first", first["text"]!.GetValue<string>());
        Assert.Equal("/Sheet1/A1", first["path"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_note_keeps_the_cell_value()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/B2", ("value", "keep me")),
            NoteOp("/Sheet1/B2", "temp note")).IsOk);

        var envelope = EditOps(file, RemoveNoteOp("/Sheet1/B2"));
        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("note", Json(envelope)["data"]!["ops"]![0]!["removed"]!.GetValue<string>());

        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Equal("keep me", cell["value"]!.GetValue<string>());
        Assert.Null(cell["note"]);

        using (var workbook = new XLWorkbook(file))
        {
            Assert.False(workbook.Worksheet("Sheet1").Cell("B2").HasComment);
        }

        AssertValidatorClean(file);
    }

    [Fact]
    public void Adding_a_second_note_to_the_same_cell_is_rejected()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, NoteOp("/Sheet1/B2", "one")).IsOk);

        var envelope = EditOps(file, NoteOp("/Sheet1/B2", "two"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("already has a note", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("target:\"note\"", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Removing_a_missing_note_is_rejected()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/B2", ("value", 1))).IsOk);

        var envelope = EditOps(file, RemoveNoteOp("/Sheet1/B2"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("no note", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_remove_target_is_rejected_with_candidates()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/B2", ("value", 1))).IsOk);

        var envelope = EditOps(file, new EditOp
        {
            Op = "remove",
            Path = "/Sheet1/B2",
            Props = new JsonObject { ["target"] = "comment" },
        });

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("note", envelope.Error.Candidates!);
    }

    [Fact]
    public void Note_on_a_non_cell_path_is_rejected()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, NoteOp("/Sheet1", "where would this go"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("cell path", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Missing_or_empty_text_is_rejected()
    {
        var file = CreateWorkbook();

        var missing = EditOps(file, new EditOp
        {
            Op = "add",
            Path = "/Sheet1/B2",
            Type = "note",
            Props = new JsonObject { ["author"] = "x" },
        });
        Assert.False(missing.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, missing.Error!.Code);
        Assert.Contains("text", missing.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Note_without_author_omits_the_author_field()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, NoteOp("/Sheet1/D4", "anonymous")).IsOk);

        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/D4"))));
        Assert.Equal("anonymous", cell["note"]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void Edits_after_a_note_keep_the_note_intact()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, NoteOp("/Sheet1/B2", "sticky")).IsOk);

        // A later, unrelated edit batch must not lose the note on resave.
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "later"))).IsOk);

        using var workbook = new XLWorkbook(file);
        Assert.True(workbook.Worksheet("Sheet1").Cell("B2").HasComment);
        Assert.Equal("sticky", workbook.Worksheet("Sheet1").Cell("B2").GetComment().Text);
        AssertValidatorClean(file);
    }
}
