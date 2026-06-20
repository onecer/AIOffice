using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.13 slide footer / slide-number / date placeholders: slide-level and
/// deck-wide set, the ph shapes + master p:hf, render sanity, get round-trips,
/// and a validator-clean package throughout.
/// </summary>
public sealed class FooterTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() =>
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Cover"))));

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private JsonObject Get(string path) =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));

    // ---- direct-OpenXml helpers (PptxDoc is internal) ----------------------

    private static List<SlidePart> SlidesInOrder(PresentationPart presentation) =>
        presentation.Presentation!.SlideIdList!.Elements<P.SlideId>()
            .Select(id => (SlidePart)presentation.GetPartById(id.RelationshipId!.Value!))
            .ToList();

    private static P.ShapeTree TreeOf(SlidePart slidePart) =>
        slidePart.Slide!.CommonSlideData!.ShapeTree!;

    private static P.Shape? Placeholder(P.ShapeTree tree, P.PlaceholderValues type) =>
        tree.Elements<P.Shape>().FirstOrDefault(s =>
            s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == type);

    private static string ShapeText(P.Shape shape) =>
        string.Concat(shape.Descendants<A.Text>().Select(t => t.Text));

    // ---- slide-level -------------------------------------------------------

    [Fact]
    public void SetFooterNumberDateOnSlide_AddsPhShapes_AndRoundTrips()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(
            ("footer", "Acme Confidential"),
            ("slideNumber", true),
            ("date", true))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tree = TreeOf(SlidesInOrder(doc.PresentationPart!).Single());

            var footer = Placeholder(tree, P.PlaceholderValues.Footer);
            Assert.NotNull(footer);
            Assert.Equal("Acme Confidential", ShapeText(footer!));

            var number = Placeholder(tree, P.PlaceholderValues.SlideNumber);
            Assert.NotNull(number);
            var numberField = number!.Descendants<A.Field>().Single();
            Assert.Equal("slidenum", numberField.Type!.Value);
            Assert.False(string.IsNullOrWhiteSpace(numberField.Text!.Text)); // cached so it renders

            var date = Placeholder(tree, P.PlaceholderValues.DateAndTime);
            Assert.NotNull(date);
            Assert.Equal("datetime1", date!.Descendants<A.Field>().Single().Type!.Value);
        }

        var detail = Get("/slide[1]");
        var footerState = detail["footer"]!.AsObject();
        Assert.Equal("Acme Confidential", footerState["footer"]!.GetValue<string>());
        Assert.True(footerState["slideNumber"]!.GetValue<bool>());
        Assert.True(footerState["date"]!.GetValue<bool>());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void FixedDateString_StoresPlainRun_NotAField()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("date", "October 2026"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tree = TreeOf(SlidesInOrder(doc.PresentationPart!).Single());
            var date = Placeholder(tree, P.PlaceholderValues.DateAndTime)!;
            Assert.Empty(date.Descendants<A.Field>()); // a fixed date is a plain run
            Assert.Equal("October 2026", ShapeText(date));
        }

        var footerState = Get("/slide[1]")["footer"]!.AsObject();
        Assert.True(footerState["date"]!.GetValue<bool>());
        Assert.Equal("October 2026", footerState["dateText"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void FooterFalse_RemovesTheShape()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("footer", "Draft"), ("slideNumber", true))));
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("footer", false), ("slideNumber", false))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tree = TreeOf(SlidesInOrder(doc.PresentationPart!).Single());
            Assert.Null(Placeholder(tree, P.PlaceholderValues.Footer));
            Assert.Null(Placeholder(tree, P.PlaceholderValues.SlideNumber));
        }

        var footerState = Get("/slide[1]")["footer"]!.AsObject();
        Assert.Null(footerState["footer"]);
        Assert.False(footerState["slideNumber"]!.GetValue<bool>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SlideNumber_RendersCachedTextInSvg()
    {
        Create();
        Edit(TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("slideNumber", true), ("footer", "Footnote"))));

        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();

        // The footer caption and the cached slide-number both reach the SVG.
        Assert.Contains("Footnote", svg, StringComparison.Ordinal);
        Assert.Contains(">1<", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void UnknownProp_StillRejected_WithFooterCandidates()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]", props: TestEnv.Props(("bogus", "x")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("footer", error.Candidates!);
        Assert.Contains("slideNumber", error.Candidates!);
        Assert.Contains("date", error.Candidates!);
    }

    // ---- deck-wide ---------------------------------------------------------

    [Fact]
    public void DeckWide_NumbersEverySlide_AndWiresMasterHf()
    {
        Create();
        Edit(
            TestEnv.Op("add", "/slide[1]", type: "slide", position: "after"),
            TestEnv.Op("add", "/slide[2]", type: "slide", position: "after"));

        Edit(TestEnv.Op("set", "/", props: TestEnv.Props(
            ("footer", "Q3 Review"),
            ("slideNumber", true),
            ("date", true))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var pres = doc.PresentationPart!;

            // Every slide carries the concrete placeholder shapes (so they render).
            foreach (var slidePart in SlidesInOrder(pres))
            {
                var tree = TreeOf(slidePart);
                Assert.NotNull(Placeholder(tree, P.PlaceholderValues.Footer));
                Assert.NotNull(Placeholder(tree, P.PlaceholderValues.SlideNumber));
                Assert.NotNull(Placeholder(tree, P.PlaceholderValues.DateAndTime));
            }

            // The master's p:hf records the visibility flags.
            var masterHf = pres.SlideMasterParts.Single().SlideMaster!.HeaderFooter!;
            Assert.True(masterHf.Footer!.Value);
            Assert.True(masterHf.SlideNumber!.Value);
            Assert.True(masterHf.DateTime!.Value);
        }

        // get / reports the deck footer state off the master.
        var deckFooter = Get("/")["footer"]!.AsObject();
        Assert.Equal("Q3 Review", deckFooter["footer"]!.GetValue<string>());
        Assert.True(deckFooter["slideNumber"]!.GetValue<bool>());
        Assert.True(deckFooter["date"]!.GetValue<bool>());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void DeckWide_CoexistsWithSlideSize_InOneOp()
    {
        Create();
        Edit(TestEnv.Op("set", "/", props: TestEnv.Props(
            ("slideSize", "4:3"),
            ("slideNumber", true))));

        var deck = Get("/");
        Assert.Equal("4:3", deck["slideSize"]!.GetValue<string>());
        Assert.True(deck["footer"]!.AsObject()["slideNumber"]!.GetValue<bool>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void DeckWide_SkipTitle_HidesFooterOnTitleLayoutSlide()
    {
        Create();
        // Slide 2 binds to a second (Blank) layout so the two slides have distinct
        // layout parts — only slide 1's layout is retyped to Title below.
        Edit(
            TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(("name", "Body"), ("basedOn", 1))),
            TestEnv.Op("add", "/slide[1]", type: "slide", position: "after", props: TestEnv.Props(("layout", 2))));

        // Make slide 1's layout a Title layout so skipTitle (default) hides its footer.
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true))
        {
            var firstSlide = SlidesInOrder(doc.PresentationPart!)[0];
            firstSlide.SlideLayoutPart!.SlideLayout!.Type = P.SlideLayoutValues.Title;
            doc.PresentationPart!.Presentation!.Save();
        }

        Edit(TestEnv.Op("set", "/", props: TestEnv.Props(("footer", "All Hands"), ("slideNumber", true))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slides = SlidesInOrder(doc.PresentationPart!);
            var titleTree = TreeOf(slides[0]);
            var bodyTree = TreeOf(slides[1]);

            // The title slide's footer is suppressed; its number still shows.
            Assert.Null(Placeholder(titleTree, P.PlaceholderValues.Footer));
            Assert.NotNull(Placeholder(titleTree, P.PlaceholderValues.SlideNumber));

            // A normal slide keeps the footer.
            Assert.NotNull(Placeholder(bodyTree, P.PlaceholderValues.Footer));
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void DeckWide_SkipTitleFalse_KeepsFooterOnTitleSlide()
    {
        Create();
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true))
        {
            var firstSlide = SlidesInOrder(doc.PresentationPart!)[0];
            firstSlide.SlideLayoutPart!.SlideLayout!.Type = P.SlideLayoutValues.Title;
            doc.PresentationPart!.Presentation!.Save();
        }

        Edit(TestEnv.Op("set", "/", props: TestEnv.Props(("footer", "Keep me"), ("skipTitle", false))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tree = TreeOf(SlidesInOrder(doc.PresentationPart!)[0]);
            var footer = Placeholder(tree, P.PlaceholderValues.Footer);
            Assert.NotNull(footer);
            Assert.Equal("Keep me", ShapeText(footer!));
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void DeckWide_ViaMasterPath_StampsThatMaster()
    {
        Create();
        Edit(TestEnv.Op("set", "/master[1]", props: TestEnv.Props(
            ("slideNumber", true),
            ("background", "0F172A"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var master = doc.PresentationPart!.SlideMasterParts.Single().SlideMaster!;
            Assert.True(master.HeaderFooter!.SlideNumber!.Value);
            Assert.NotNull(Placeholder(master.CommonSlideData!.ShapeTree!, P.PlaceholderValues.SlideNumber));

            // The background prop on the same op still applied (the rest flows on).
            Assert.NotNull(master.CommonSlideData!.Background);

            var slideTree = TreeOf(SlidesInOrder(doc.PresentationPart!).Single());
            Assert.NotNull(Placeholder(slideTree, P.PlaceholderValues.SlideNumber));
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void DeckWide_UnknownFooterKey_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/", props: TestEnv.Props(("footer", "x"), ("bogus", true)))]);

        // The rest (bogus) flows to slide-size, which rejects it as invalid_args.
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }
}
