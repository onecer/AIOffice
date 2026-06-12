using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>M2 speaker notes: /slide[i]/notes get/set/add/remove + notes master wiring.</summary>
public sealed class NotesTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void CreateTwoSlideDeck()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Cover"))));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("title", "Body"))),
        ]));
    }

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private JsonObject GetNotes(int slide) =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", $"/slide[{slide}]/notes"))));

    [Fact]
    public void SetNotes_CreatesPartWithNotesMasterWiring()
    {
        CreateTwoSlideDeck();
        var data = Edit(TestEnv.Op("set", "/slide[2]/notes", props: TestEnv.Props(("text", "Remember the demo"))));
        Assert.Equal("/slide[2]/notes", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var presentation = doc.PresentationPart!;

            // The deck gains exactly one notes master, registered in p:notesMasterIdLst with a live r:id.
            var notesMaster = presentation.NotesMasterPart;
            Assert.NotNull(notesMaster);
            Assert.NotNull(notesMaster!.NotesMaster!.ColorMap);
            Assert.NotNull(notesMaster.ThemePart?.Theme);
            var masterId = Assert.Single(
                presentation.Presentation!.NotesMasterIdList!.Elements<P.NotesMasterId>());
            Assert.Same(notesMaster, presentation.GetPartById(masterId.Id!.Value!));

            // Slide 2 (and only slide 2) has a notes part referencing slide and notes master.
            var slides = doc.PresentationPart!.Presentation!.SlideIdList!.Elements<P.SlideId>()
                .Select(id => (SlidePart)presentation.GetPartById(id.RelationshipId!.Value!))
                .ToList();
            Assert.Null(slides[0].NotesSlidePart);
            var notesPart = slides[1].NotesSlidePart;
            Assert.NotNull(notesPart);
            Assert.Contains(notesPart!.Parts, p => p.OpenXmlPart is NotesMasterPart);
            Assert.Contains(notesPart.Parts, p => p.OpenXmlPart is SlidePart);
        }

        var notes = GetNotes(2);
        Assert.True(notes["exists"]!.GetValue<bool>());
        Assert.Equal("Remember the demo", notes["text"]!.GetValue<string>());
        Assert.Equal("Remember the demo", notes["paragraphs"]!.AsArray().Single()!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetNotes_ReplacesExistingText()
    {
        CreateTwoSlideDeck();
        Edit(TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("text", "first draft"))));
        Edit(TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("text", "final\ntwo lines"))));

        var notes = GetNotes(1);
        Assert.Equal("final\ntwo lines", notes["text"]!.GetValue<string>());
        Assert.Equal(2, notes["paragraphs"]!.AsArray().Count);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddNotes_AppendsParagraphs()
    {
        CreateTwoSlideDeck();
        Edit(TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("text", "opening"))));
        Edit(TestEnv.Op("add", "/slide[1]/notes", props: TestEnv.Props(("text", "follow-up"))));

        var notes = GetNotes(1);
        Assert.Equal("opening\nfollow-up", notes["text"]!.GetValue<string>());
        Assert.Equal(2, notes["paragraphs"]!.AsArray().Count);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddNotes_OnSlideWithoutNotes_CreatesThePart()
    {
        CreateTwoSlideDeck();
        Edit(TestEnv.Op("add", "/slide[2]/notes", props: TestEnv.Props(("text", "only paragraph"))));

        var notes = GetNotes(2);
        Assert.True(notes["exists"]!.GetValue<bool>());
        Assert.Equal("only paragraph", notes["text"]!.GetValue<string>());
        Assert.Single(notes["paragraphs"]!.AsArray());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveNotes_ClearsAndIsIdempotent()
    {
        CreateTwoSlideDeck();
        Edit(TestEnv.Op("set", "/slide[2]/notes", props: TestEnv.Props(("text", "throwaway"))));
        Edit(TestEnv.Op("remove", "/slide[2]/notes"));

        var notes = GetNotes(2);
        Assert.False(notes["exists"]!.GetValue<bool>());
        Assert.Equal(string.Empty, notes["text"]!.GetValue<string>());
        Assert.Empty(notes["paragraphs"]!.AsArray());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.All(doc.PresentationPart!.SlideParts, slide => Assert.Null(slide.NotesSlidePart));
        }

        // Clearing already-clear notes stays a success (clear semantics, not invalid_path).
        Edit(TestEnv.Op("remove", "/slide[2]/notes"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void GetNotes_WithoutNotesPart_ReturnsEmpty()
    {
        CreateTwoSlideDeck();
        var notes = GetNotes(1);

        Assert.Equal("/slide[1]/notes", notes["path"]!.GetValue<string>());
        Assert.False(notes["exists"]!.GetValue<bool>());
        Assert.Equal(string.Empty, notes["text"]!.GetValue<string>());
        Assert.Empty(notes["paragraphs"]!.AsArray());
    }

    [Fact]
    public void Outline_IncludesNotesSnippet()
    {
        CreateTwoSlideDeck();
        Edit(TestEnv.Op("set", "/slide[2]/notes", props: TestEnv.Props(
            ("text", new string('x', 100)))));

        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        var slides = outline["slides"]!.AsArray();

        Assert.Null(slides[0]!["notes"]); // no notes -> no field (nulls are dropped on the wire)
        var snippet = slides[1]!["notes"]!.GetValue<string>();
        Assert.Equal(80, snippet.Length);
        Assert.EndsWith("...", snippet, StringComparison.Ordinal);
    }

    [Fact]
    public void TextView_IncludesNotesSection_RenderTextDoesNot()
    {
        CreateTwoSlideDeck();
        Edit(TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("text", "presenter cue"))));

        var read = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))))["text"]!.GetValue<string>();
        Assert.Contains("[notes]\npresenter cue", read, StringComparison.Ordinal);

        var rendered = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "text"))))["text"]!.GetValue<string>();
        Assert.DoesNotContain("presenter cue", rendered, StringComparison.Ordinal);
    }

    [Fact]
    public void NotesParagraphAddressing_IsTypedUnsupportedFeature()
    {
        CreateTwoSlideDeck();
        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/notes/p[1]")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("/slide[i]/notes", error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void MoveNotes_IsInvalidArgs()
    {
        CreateTwoSlideDeck();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("move", "/slide[1]/notes", position: "2")]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void SetNotes_WithoutText_IsInvalidArgs()
    {
        CreateTwoSlideDeck();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("bold", JsonValue.Create(true))))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Equal(new[] { "text" }, error.Candidates!);
    }

    [Fact]
    public void NotesOnOutOfRangeSlide_IsInvalidPathWithCandidates()
    {
        CreateTwoSlideDeck();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[9]/notes", props: TestEnv.Props(("text", "ghost")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.NotNull(error.Candidates);
        Assert.NotEmpty(error.Candidates!);
    }

    [Fact]
    public void SecondSlideNotes_ReusesTheSingleNotesMaster()
    {
        CreateTwoSlideDeck();
        Edit(
            TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("text", "one"))),
            TestEnv.Op("set", "/slide[2]/notes", props: TestEnv.Props(("text", "two"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var presentation = doc.PresentationPart!;
        Assert.Single(presentation.Presentation!.NotesMasterIdList!.Elements<P.NotesMasterId>());
        Assert.Equal(2, presentation.SlideParts.Count(s => s.NotesSlidePart is not null));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
