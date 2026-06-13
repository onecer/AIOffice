using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>M8 pptx diff: slide/shape/size/section/background changes, deterministic order.</summary>
public sealed class DiffTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    /// <summary>Creates a fresh deck file in the workspace and returns its name.</summary>
    private string CreateDeck(string name, string? title = null)
    {
        var ctx = title is null ? _ws.Ctx(name) : _ws.Ctx(name, ("title", title));
        TestEnv.AssertOk(_handler.Create(ctx));
        return name;
    }

    private JsonObject Edit(string name, params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx(name), ops));

    /// <summary>Diffs current vs baseline (current is ctx.File, baseline is the OLD deck).</summary>
    private DiffResult Diff(string current, string baseline) =>
        _handler.Diff(_ws.Ctx(current), _ws.PathOf(baseline));

    private static DiffChange One(DiffResult result, string kind) =>
        Assert.Single(result.Changes, c => c.Kind == kind);

    private static DiffChange OnePath(DiffResult result, string path) =>
        Assert.Single(result.Changes, c => c.Path == path);

    private static DiffChange OneDetail(DiffResult result, string detail) =>
        Assert.Single(result.Changes, c => c.Detail == detail);

    /// <summary>Copies a workspace file to a new name (so the copy shares slide ids).</summary>
    private string Copy(string from, string to)
    {
        File.Copy(_ws.PathOf(from), _ws.PathOf(to));
        return to;
    }

    // ---- identical ----------------------------------------------------------

    [Fact]
    public void IdenticalDecks_HaveNoChanges()
    {
        CreateDeck("base.pptx", "Title");
        Edit("base.pptx", TestEnv.Op("add", "/slide[1]", type: "slide", position: "after",
            props: TestEnv.Props(("title", JsonValue.Create("Two")))));
        var current = Copy("base.pptx", "current.pptx");

        var result = Diff(current, "base.pptx");
        Assert.Empty(result.Changes);
        Assert.Equal(new DiffSummary(0, 0, 0, 0), result.Summary);
    }

    // ---- shape text ---------------------------------------------------------

    [Fact]
    public void ShapeTextChange_IsModifiedWithBeforeAfter()
    {
        CreateDeck("base.pptx");
        var added = Edit("base.pptx", TestEnv.Op("add", "/slide[1]", type: "shape",
            props: TestEnv.Props(("text", JsonValue.Create("old text")))));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();
        var current = Copy("base.pptx", "current.pptx");
        Edit(current, TestEnv.Op("set", shapePath, props: TestEnv.Props(("text", JsonValue.Create("new text")))));

        var result = Diff(current, "base.pptx");
        var change = One(result, "modified");
        Assert.Equal(shapePath, change.Path);
        Assert.Equal("text", change.Detail);
        Assert.Equal("old text", change.Before);
        Assert.Equal("new text", change.After);
        Assert.Equal(1, result.Summary.Modified);
    }

    // ---- shape position -----------------------------------------------------

    [Fact]
    public void ShapePositionChange_IsModifiedMovedOnSlide()
    {
        CreateDeck("base.pptx");
        var added = Edit("base.pptx", TestEnv.Op("add", "/slide[1]", type: "shape",
            props: TestEnv.Props(("text", JsonValue.Create("same")), ("x", JsonValue.Create(2)), ("y", JsonValue.Create(2)))));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();
        var current = Copy("base.pptx", "current.pptx");
        Edit(current, TestEnv.Op("set", shapePath, props: TestEnv.Props(("x", JsonValue.Create(10)))));

        var change = One(Diff(current, "base.pptx"), "modified");
        Assert.Equal(shapePath, change.Path);
        Assert.Equal("moved on slide", change.Detail);
    }

    // ---- shape add / remove -------------------------------------------------

    [Fact]
    public void ShapeAdded_IsAddedChange()
    {
        CreateDeck("base.pptx");
        var current = Copy("base.pptx", "current.pptx");
        var added = Edit(current, TestEnv.Op("add", "/slide[1]", type: "shape",
            props: TestEnv.Props(("text", JsonValue.Create("brand new")))));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();

        var change = One(Diff(current, "base.pptx"), "added");
        Assert.Equal(shapePath, change.Path);
        Assert.Equal("shape added", change.Detail);
    }

    [Fact]
    public void ShapeRemoved_IsRemovedChange()
    {
        CreateDeck("base.pptx");
        var added = Edit("base.pptx", TestEnv.Op("add", "/slide[1]", type: "shape",
            props: TestEnv.Props(("text", JsonValue.Create("goner")))));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();
        var current = Copy("base.pptx", "current.pptx");
        Edit(current, TestEnv.Op("remove", shapePath));

        var change = One(Diff(current, "base.pptx"), "removed");
        Assert.Equal(shapePath, change.Path);
        Assert.Equal("shape removed", change.Detail);
    }

    // ---- slide add / remove / move -----------------------------------------

    [Fact]
    public void SlideAdded_IsAddedChange()
    {
        CreateDeck("base.pptx", "Intro");
        var current = Copy("base.pptx", "current.pptx");
        Edit(current, TestEnv.Op("add", "/slide[1]", type: "slide", position: "after",
            props: TestEnv.Props(("title", JsonValue.Create("Appendix")))));

        var result = Diff(current, "base.pptx");
        var change = One(result, "added");
        Assert.Equal("/slide[2]", change.Path);
        Assert.Equal("slide added", change.Detail);
        Assert.Equal(1, result.Summary.Added);
    }

    [Fact]
    public void SlideRemoved_IsRemovedChange()
    {
        CreateDeck("base.pptx", "Intro");
        Edit("base.pptx", TestEnv.Op("add", "/slide[1]", type: "slide", position: "after",
            props: TestEnv.Props(("title", JsonValue.Create("Body")))));
        var current = Copy("base.pptx", "current.pptx");
        Edit(current, TestEnv.Op("remove", "/slide[2]"));

        var change = One(Diff(current, "base.pptx"), "removed");
        Assert.Equal("slide removed", change.Detail);
        Assert.Equal("/slide[2]", change.Path);
    }

    [Fact]
    public void SlideReordered_IsMoved()
    {
        CreateDeck("base.pptx", "One");
        Edit("base.pptx", TestEnv.Op("add", "/slide[1]", type: "slide", position: "after",
            props: TestEnv.Props(("title", JsonValue.Create("Two")))));
        Edit("base.pptx", TestEnv.Op("add", "/slide[2]", type: "slide", position: "after",
            props: TestEnv.Props(("title", JsonValue.Create("Three")))));
        var current = Copy("base.pptx", "current.pptx");
        // Move slide 3 to the front: 3,1,2.
        Edit(current, TestEnv.Op("move", "/slide[3]", position: "1"));

        var result = Diff(current, "base.pptx");
        // Exactly one slide is flagged as moved (the LIS keeps the other two stable).
        Assert.Single(result.Changes, c => c.Kind == "moved");
        Assert.Equal(1, result.Summary.Moved);
        Assert.Equal(0, result.Summary.Added);
        Assert.Equal(0, result.Summary.Removed);
    }

    // ---- slide size ---------------------------------------------------------

    [Fact]
    public void SlideSizeChange_IsModifiedOnRoot()
    {
        CreateDeck("base.pptx");
        var current = Copy("base.pptx", "current.pptx");
        Edit(current, TestEnv.Op("set", "/", props: TestEnv.Props(("slideSize", JsonValue.Create("4:3")))));

        var change = One(Diff(current, "base.pptx"), "modified");
        Assert.Equal("/", change.Path);
        Assert.Equal("slide size", change.Detail);
        Assert.Equal("4:3", change.After);
    }

    // ---- master background --------------------------------------------------

    [Fact]
    public void MasterBackgroundChange_IsModifiedOnMaster()
    {
        CreateDeck("base.pptx");
        var current = Copy("base.pptx", "current.pptx");
        Edit(current, TestEnv.Op("set", "/master[1]", props: TestEnv.Props(("background", JsonValue.Create("0F172A")))));

        var change = OnePath(Diff(current, "base.pptx"), "/master[1]");
        Assert.Equal("modified", change.Kind);
        Assert.Equal("master background", change.Detail);
        Assert.Equal("0F172A", change.After);
    }

    // ---- sections -----------------------------------------------------------

    [Fact]
    public void SectionRename_IsModifiedSectionName()
    {
        CreateDeck("base.pptx");
        Edit("base.pptx", TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(("name", JsonValue.Create("Intro")))));
        var current = Copy("base.pptx", "current.pptx");
        Edit(current, TestEnv.Op("set", "/section[1]", props: TestEnv.Props(("name", JsonValue.Create("Overview")))));

        var change = OnePath(Diff(current, "base.pptx"), "/section[1]");
        Assert.Equal("modified", change.Kind);
        Assert.Equal("Intro", change.Before);
        Assert.Equal("Overview", change.After);
    }

    // ---- determinism --------------------------------------------------------

    [Fact]
    public void Output_IsSortedByPathThenKind_AndStableAcrossRuns()
    {
        CreateDeck("base.pptx", "Intro");
        var s1 = Edit("base.pptx", TestEnv.Op("add", "/slide[1]", type: "shape",
            props: TestEnv.Props(("text", JsonValue.Create("first")))))["results"]![0]!["target"]!.GetValue<string>();
        Edit("base.pptx", TestEnv.Op("add", "/slide[1]", type: "slide", position: "after",
            props: TestEnv.Props(("title", JsonValue.Create("Two")))));
        var current = Copy("base.pptx", "current.pptx");

        // A mix of changes: edit a shape, change size, add a slide, rename nothing.
        Edit(current, TestEnv.Op("set", s1, props: TestEnv.Props(("text", JsonValue.Create("FIRST")))));
        Edit(current, TestEnv.Op("set", "/", props: TestEnv.Props(("slideSize", JsonValue.Create("4:3")))));
        Edit(current, TestEnv.Op("add", "/slide[2]", type: "slide", position: "after",
            props: TestEnv.Props(("title", JsonValue.Create("Three")))));

        var first = Diff(current, "base.pptx");
        var second = Diff(current, "base.pptx");

        // Deterministic: two runs produce the identical ordered list.
        Assert.Equal(
            first.Changes.Select(c => (c.Path, c.Kind)).ToList(),
            second.Changes.Select(c => (c.Path, c.Kind)).ToList());

        // Sorted by canonical path (ordinal) then kind.
        var paths = first.Changes.Select(c => c.Path).ToList();
        var sorted = paths.OrderBy(p => p, StringComparer.Ordinal).ToList();
        Assert.Equal(sorted, paths);
    }

    // ---- format mismatch ----------------------------------------------------

    [Fact]
    public void Baseline_OfWrongFormat_IsInvalidArgs()
    {
        CreateDeck("base.pptx");
        var current = Copy("base.pptx", "current.pptx");
        File.WriteAllText(_ws.PathOf("baseline.docx"), "not really a docx");

        var ex = Assert.Throws<AiofficeException>(() =>
            _handler.Diff(_ws.Ctx(current), _ws.PathOf("baseline.docx")));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.False(string.IsNullOrWhiteSpace(ex.Suggestion));
    }

    // ---- hyperlink facet ----------------------------------------------------

    [Fact]
    public void HyperlinkChange_IsModifiedHyperlink()
    {
        CreateDeck("base.pptx");
        var added = Edit("base.pptx", TestEnv.Op("add", "/slide[1]", type: "shape",
            props: TestEnv.Props(("text", JsonValue.Create("Link")))));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();
        var current = Copy("base.pptx", "current.pptx");
        Edit(current, TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create("https://example.com")))));

        var change = OneDetail(Diff(current, "base.pptx"), "hyperlink");
        Assert.Equal("modified", change.Kind);
        Assert.Equal("(none)", change.Before);
        Assert.Equal("https://example.com", change.After);
    }
}
