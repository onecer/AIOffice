using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.4.0 zoom links: slide/section/summary navigation objects built as a p:graphicFrame
/// hosting the 2018 p159 zoom payload, referencing target slides via real slide
/// relationships. /slide[i]/zoom[k] addressing, validator-clean, SVG placeholder.
/// </summary>
public sealed class ZoomTests : IDisposable
{
    private const string ZoomNs = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";

    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    /// <summary>Creates deck.pptx and grows it to the given slide count.</summary>
    private void CreateWith(int slides)
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Slide 1"))));
        for (var i = 2; i <= slides; i++)
        {
            Edit(TestEnv.Op("add", $"/slide[{i}]", type: "slide", props: TestEnv.Props(("title", $"Slide {i}"))));
        }
    }

    private void AddSection(string name, int afterSlide) =>
        Edit(TestEnv.Op("add", "/", type: "section", props: TestEnv.Props(
            ("name", name), ("afterSlide", JsonValue.Create(afterSlide)))));

    private P.GraphicFrame ZoomFrame(PresentationDocument doc, int slideIndex = 1) =>
        doc.PresentationPart!.SlideParts.First(sp =>
                sp.Slide!.CommonSlideData!.ShapeTree!.Elements<P.GraphicFrame>()
                    .Any(f => f.Graphic?.GraphicData?.Uri?.Value == ZoomNs))
            .Slide!.CommonSlideData!.ShapeTree!.Elements<P.GraphicFrame>()
            .First(f => f.Graphic?.GraphicData?.Uri?.Value == ZoomNs);

    // ----- slide zoom -------------------------------------------------------------

    [Fact]
    public void AddSlideZoom_BuildsGraphicFrame_WithSlideRelationship_AndValidates()
    {
        CreateWith(3);
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(
            ("kind", "slide"), ("target", "slide 3"), ("x", "2cm"), ("y", "2cm"), ("w", "8cm"), ("h", "5cm"))));
        Assert.Equal("/slide[1]/zoom[1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var frame = ZoomFrame(doc);
            var graphicData = frame.Graphic!.GraphicData!;
            Assert.Equal(ZoomNs, graphicData.Uri!.Value);

            // The slideZoom payload carries an r:embed pointing at a real slide relationship.
            var xml = graphicData.OuterXml;
            Assert.Contains("slideZoom", xml, StringComparison.Ordinal);
            Assert.Contains("r:embed=", xml, StringComparison.Ordinal);
            var hostSlide = doc.PresentationPart!.SlideParts.First(sp =>
                sp.Slide!.CommonSlideData!.ShapeTree!.Elements<P.GraphicFrame>()
                    .Any(f => f.Graphic?.GraphicData?.Uri?.Value == ZoomNs));
            // The host slide carries a slide->slide relationship for the zoom target.
            Assert.Contains(hostSlide.Parts, p => p.OpenXmlPart is SlidePart);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_SlideZoom_ReportsKindAndResolvedTarget()
    {
        CreateWith(3);
        Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(
            ("kind", "slide"), ("target", "slide 3"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/zoom[1]"))));
        Assert.Equal("/slide[1]/zoom[1]", detail["path"]!.GetValue<string>());
        Assert.Equal("slide", detail["kind"]!.GetValue<string>());
        Assert.Equal("slide 3", detail["target"]!.GetValue<string>());
    }

    [Fact]
    public void SlideZoom_TargetAcceptsSlidePathForm()
    {
        CreateWith(4);
        Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(
            ("kind", "slide"), ("target", "/slide[4]"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/zoom[1]"))));
        Assert.Equal("slide 4", detail["target"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SlideZoom_OutOfRangeTarget_IsInvalidPath()
    {
        CreateWith(2);
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "zoom",
                props: TestEnv.Props(("kind", "slide"), ("target", "slide 9")))]),
            ErrorCodes.InvalidPath);
    }

    // ----- section zoom -----------------------------------------------------------

    [Fact]
    public void AddSectionZoom_ReferencesSectionFirstSlide_WithSectionIdx_AndValidates()
    {
        CreateWith(4);
        AddSection("Intro", 0);   // slides 1..2
        AddSection("Body", 2);    // slides 3..4

        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(
            ("kind", "section"), ("target", "Body"))));
        Assert.Equal("/slide[1]/zoom[1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var xml = ZoomFrame(doc).Graphic!.GraphicData!.OuterXml;
            Assert.Contains("sectionZoom", xml, StringComparison.Ordinal);
            Assert.Contains("sectionIdx=\"2\"", xml, StringComparison.Ordinal); // Body is the 2nd section
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/zoom[1]"))));
        Assert.Equal("section", detail["kind"]!.GetValue<string>());
        Assert.Equal("section 2", detail["target"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SectionZoom_UnknownSection_IsInvalidPath_WithSectionNameCandidates()
    {
        CreateWith(2);
        AddSection("Intro", 0);

        var error = TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "zoom",
                props: TestEnv.Props(("kind", "section"), ("target", "Nope")))]),
            ErrorCodes.InvalidPath);
        Assert.Contains("Intro", error.Candidates!);
    }

    // ----- summary zoom -----------------------------------------------------------

    [Fact]
    public void AddSummaryZoom_ReferencesEverySectionFirstSlide_AndValidates()
    {
        CreateWith(4);
        AddSection("Intro", 0);   // slides 1..2
        AddSection("Body", 2);    // slides 3..4

        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(("kind", "summary"))));
        Assert.Equal("/slide[1]/zoom[1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var xml = ZoomFrame(doc).Graphic!.GraphicData!.OuterXml;
            Assert.Contains("summaryZoom", xml, StringComparison.Ordinal);
            // One p159:section child per deck section.
            Assert.Equal(2, xml.Split("<p159:section ", StringSplitOptions.None).Length - 1);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/zoom[1]"))));
        Assert.Equal("summary", detail["kind"]!.GetValue<string>());
        Assert.Equal("2 section(s)", detail["target"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SummaryZoom_WithNoSections_IsInvalidArgs()
    {
        CreateWith(2);
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "zoom",
                props: TestEnv.Props(("kind", "summary")))]),
            ErrorCodes.InvalidArgs);
    }

    // ----- structure + multiple ---------------------------------------------------

    [Fact]
    public void Structure_ListsZoomsPerSlide()
    {
        CreateWith(3);
        Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(("kind", "slide"), ("target", "slide 2"))));
        Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(("kind", "slide"), ("target", "slide 3"))));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var zooms = data["slides"]![0]!["zooms"]!.AsArray();
        Assert.Equal(2, zooms.Count);
        Assert.Equal("/slide[1]/zoom[1]", zooms[0]!["path"]!.GetValue<string>());
        Assert.Equal("slide 2", zooms[0]!["target"]!.GetValue<string>());
        Assert.Equal("/slide[1]/zoom[2]", zooms[1]!["path"]!.GetValue<string>());
        Assert.Equal("slide 3", zooms[1]!["target"]!.GetValue<string>());
    }

    // ----- remove -----------------------------------------------------------------

    [Fact]
    public void Remove_Zoom_DropsTheFrame_AndStaysValid()
    {
        CreateWith(3);
        Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(("kind", "slide"), ("target", "slide 2"))));
        Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(("kind", "slide"), ("target", "slide 3"))));

        var data = Edit(TestEnv.Op("remove", "/slide[1]/zoom[1]"));
        Assert.Equal("/slide[1]/zoom[1]", data["results"]![0]!["target"]!.GetValue<string>());

        // The surviving zoom (formerly zoom[2]) re-indexes to zoom[1].
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/zoom[1]"))));
        Assert.Equal("slide 3", detail["target"]!.GetValue<string>());
        TestEnv.AssertFail(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/zoom[2]"))),
            ErrorCodes.InvalidPath);

        // Removing the zoom that pointed at slide 2 must NOT delete slide 2 from the deck.
        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        Assert.Equal(3, outline["slides"]!.AsArray().Count);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_ZoomOutOfRange_IsInvalidPath()
    {
        CreateWith(2);
        TestEnv.AssertFail(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/zoom[1]"))),
            ErrorCodes.InvalidPath);
    }

    // ----- guards -----------------------------------------------------------------

    [Fact]
    public void AddZoom_OnShapePath_IsInvalidArgs()
    {
        CreateWith(2);
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]/shape[1]", type: "zoom",
                props: TestEnv.Props(("kind", "slide"), ("target", "slide 2")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Zoom_UnknownKind_IsTypedUnsupported()
    {
        CreateWith(2);
        var error = TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "zoom",
                props: TestEnv.Props(("kind", "warp"), ("target", "slide 2")))]),
            ErrorCodes.UnsupportedFeature);
        Assert.Equal(["slide", "section", "summary"], error.Candidates!);
    }

    [Fact]
    public void SetZoom_IsTypedUnsupported_RemoveAndReadd()
    {
        CreateWith(3);
        Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(("kind", "slide"), ("target", "slide 2"))));
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", "/slide[1]/zoom[1]",
                props: TestEnv.Props(("target", "slide 3")))]),
            ErrorCodes.UnsupportedFeature);
    }

    // ----- render -----------------------------------------------------------------

    [Fact]
    public void Svg_DrawsZoomPlaceholder_WithThumbnailLabelAndPath()
    {
        CreateWith(3);
        Edit(TestEnv.Op("add", "/slide[1]", type: "zoom", props: TestEnv.Props(
            ("kind", "slide"), ("target", "slide 3"))));

        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();

        Assert.Contains("class=\"aio-zoom\"", svg, StringComparison.Ordinal);
        Assert.Contains("class=\"aio-zoom-thumb\"", svg, StringComparison.Ordinal);
        Assert.Contains("[slide zoom] slide 3", svg, StringComparison.Ordinal);
        Assert.Contains("<g data-aio-path=\"/slide[1]/shape[@id=", svg, StringComparison.Ordinal);
    }
}
