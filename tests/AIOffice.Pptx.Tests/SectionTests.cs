using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// M6 slide sections — the standard PowerPoint p14:sectionLst in
/// presentation.xml's extLst: add (with afterSlide ranges), rename, remove
/// (slides survive), and outline grouping.
/// </summary>
public sealed class SectionTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    /// <summary>Creates deck.pptx and grows it to the given slide count.</summary>
    private void CreateWith(int slides)
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Slide 1"))));
        for (var i = 2; i <= slides; i++)
        {
            Edit(TestEnv.Op("add", $"/slide[{i}]", type: "slide", props: TestEnv.Props(("title", $"Slide {i}"))));
        }
    }

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    // ---- add -----------------------------------------------------------------

    // Regression: section add (and slide-size set) target the "/" root, which
    // the CLI/MCP route through EditOp.ParseBatch -> DocPath.Parse. Before M6
    // made "/" a real path form, a non-replace op on "/" was rejected at that
    // gate, so sections were reachable only by the tests that build EditOps
    // directly. This drives the production gate end to end.
    [Fact]
    public void Root_section_and_slideSize_ops_survive_the_ParseBatch_gate()
    {
        CreateWith(2);

        var sectionOps = EditOp.ParseBatch(
            "[{\"op\":\"add\",\"path\":\"/\",\"type\":\"section\",\"props\":{\"name\":\"Intro\"}}]");
        Assert.True(_handler.Edit(_ws.Ctx("deck.pptx"), sectionOps).IsOk);

        var sizeOps = EditOp.ParseBatch("[{\"op\":\"set\",\"path\":\"/\",\"props\":{\"slideSize\":\"4:3\"}}]");
        Assert.True(_handler.Edit(_ws.Ctx("deck.pptx"), sizeOps).IsOk);

        var root = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/"))));
        Assert.Equal("4:3", root["slideSize"]!.GetValue<string>());
        Assert.Equal(1, root["sectionCount"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddSection_NoAfterSlide_ClaimsAllSlides_ReopenVerified()
    {
        CreateWith(3);
        var data = Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(("name", "All"))));
        Assert.Equal("/section[1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var list = doc.PresentationPart!.Presentation!.PresentationExtensionList!
                .Descendants<P14.SectionList>().Single();
            var section = list.Elements<P14.Section>().Single();
            Assert.Equal("All", section.Name!.Value);
            Assert.Equal(3, section.SectionSlideIdList!.Elements<P14.SectionSlideIdListEntry>().Count());
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/section[1]"))));
        Assert.Equal(3, detail["slideCount"]!.GetValue<int>());
        Assert.Equal(new[] { "/slide[1]", "/slide[2]", "/slide[3]" }, detail["slides"]!.AsArray().Select(n => n!.GetValue<string>()).ToArray());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddTwoSections_PartitionTheSlidesByRange()
    {
        CreateWith(4);
        // Intro before slide 1 claims unsectioned slides up to the next section.
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(
            ("name", "Body"), ("afterSlide", JsonValue.Create(2)))));
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(
            ("name", "Intro"), ("afterSlide", JsonValue.Create(0)))));

        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        var sections = outline["sections"]!.AsArray();
        Assert.Equal(2, sections.Count);
        // Sections sort by first slide: Intro (slides 1-2) then Body (slides 3-4).
        Assert.Equal("Intro", sections[0]!["name"]!.GetValue<string>());
        Assert.Equal(new[] { "/slide[1]", "/slide[2]" }, sections[0]!["slides"]!.AsArray().Select(n => n!.GetValue<string>()).ToArray());
        Assert.Equal("Body", sections[1]!["name"]!.GetValue<string>());
        Assert.Equal(new[] { "/slide[3]", "/slide[4]" }, sections[1]!["slides"]!.AsArray().Select(n => n!.GetValue<string>()).ToArray());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Outline_TagsEachSlideWithItsSection()
    {
        CreateWith(2);
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(("name", "Cover"))));

        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        Assert.Equal("Cover", outline["slides"]!.AsArray()[0]!["section"]!.GetValue<string>());
    }

    [Fact]
    public void Structure_ListsSections()
    {
        CreateWith(2);
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(("name", "Deck"))));

        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var sections = structure["sections"]!.AsArray();
        Assert.Single(sections);
        Assert.Equal("/section[1]", sections[0]!["path"]!.GetValue<string>());
        Assert.Equal("Deck", sections[0]!["name"]!.GetValue<string>());
    }

    [Fact]
    public void AddSection_WithoutName_IsInvalidArgs()
    {
        CreateWith(1);
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(("afterSlide", JsonValue.Create(0)))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void AddSection_AfterSlideOutOfRange_IsInvalidArgs()
    {
        CreateWith(2);
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(
                ("name", "X"), ("afterSlide", JsonValue.Create(9)))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- rename --------------------------------------------------------------

    [Fact]
    public void RenameSection_UpdatesTheName_ReopenVerified()
    {
        CreateWith(2);
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(("name", "Old"))));
        Edit(TestEnv.Op("set", "/section[1]", props: TestEnv.Props(("name", "New"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/section[1]"))));
        Assert.Equal("New", detail["name"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RenameSection_EmptyName_IsInvalidArgs()
    {
        CreateWith(1);
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(("name", "X"))));
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/section[1]", props: TestEnv.Props(("name", "  "))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void SetUnknownSectionProp_IsInvalidArgs()
    {
        CreateWith(1);
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(("name", "X"))));
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/section[1]", props: TestEnv.Props(("afterSlide", JsonValue.Create(1)))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- remove --------------------------------------------------------------

    [Fact]
    public void RemoveSection_KeepsItsSlides_AndDropsTheEmptyList()
    {
        CreateWith(2);
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(("name", "Temp"))));
        Edit(TestEnv.Op("remove", "/section[1]"));

        // Slides survive.
        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        Assert.Equal(2, outline["slides"]!.AsArray().Count);
        Assert.Null(outline["sections"]);

        // The empty p14:sectionLst (and its ext) were pruned.
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Empty(doc.PresentationPart!.Presentation!.Descendants<P14.SectionList>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveMiddleSection_RelinksIndices()
    {
        CreateWith(3);
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(
            ("name", "A"), ("afterSlide", JsonValue.Create(0)))));
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(
            ("name", "B"), ("afterSlide", JsonValue.Create(1)))));
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(
            ("name", "C"), ("afterSlide", JsonValue.Create(2)))));

        Edit(TestEnv.Op("remove", "/section[2]"));

        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var names = structure["sections"]!.AsArray().Select(s => s!["name"]!.GetValue<string>()).ToArray();
        Assert.Equal(new[] { "A", "C" }, names);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveSection_OutOfRange_IsInvalidPathWithCandidates()
    {
        CreateWith(1);
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(("name", "Only"))));
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/section[5]")]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.Contains("/section[1]", error.Candidates!);
    }

    [Fact]
    public void GetSection_OnDeckWithoutSections_IsInvalidPath()
    {
        CreateWith(1);
        TestEnv.AssertFail(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/section[1]"))), ErrorCodes.InvalidPath);
    }

    [Fact]
    public void AddSection_OnSlidePath_IsInvalidArgs()
    {
        CreateWith(1);
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "section", props: TestEnv.Props(("name", "X"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Sections_SurviveSlideReorder_ByTrackingSlideIds()
    {
        CreateWith(3);
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(
            ("name", "First"), ("afterSlide", JsonValue.Create(0)))));
        // The section tracks slide ids; moving slide 1 to the end keeps the id in the section.
        Edit(TestEnv.Op("move", "/slide[1]", position: "3"));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/section[1]"))));
        // The original slide 1 is now slide 3; the section still owns it (its id is unchanged).
        Assert.Contains("/slide[3]", detail["slides"]!.AsArray().Select(n => n!.GetValue<string>()));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
