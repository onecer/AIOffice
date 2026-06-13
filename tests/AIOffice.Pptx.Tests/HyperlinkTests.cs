using System.Globalization;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>M8 shape hyperlinks / actions: external urls, slide jumps and show actions.</summary>
public sealed class HyperlinkTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private string Create()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        return _ws.PathOf("deck.pptx");
    }

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private void AddSlides(int extra)
    {
        for (var i = 0; i < extra; i++)
        {
            Edit(TestEnv.Op("add", $"/slide[{i + 1}]", type: "slide", position: "after"));
        }
    }

    /// <summary>Adds a textbox to slide 1 and returns its canonical @id shape path.</summary>
    private string AddShape(string text = "Click me")
    {
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", JsonValue.Create(text)))));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private JsonObject Get(string path) =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));

    private static uint IdOf(string shapePath) =>
        uint.Parse(shapePath.Split("@id=")[1].TrimEnd(']'), CultureInfo.InvariantCulture);

    /// <summary>Slide parts in show order (sldIdLst order, not part order).</summary>
    private static List<SlidePart> SlidesInOrder(PresentationPart presentation)
    {
        var result = new List<SlidePart>();
        foreach (var id in presentation.Presentation!.SlideIdList!.Elements<P.SlideId>())
        {
            result.Add((SlidePart)presentation.GetPartById(id.RelationshipId!.Value!));
        }

        return result;
    }

    private static P.Shape ShapeById(SlidePart slide, uint id) =>
        slide.Slide!.Descendants<P.Shape>().First(s =>
            s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == id);

    private A.HyperlinkOnClick? Hlink(string file, int slideIndex, uint shapeId)
    {
        using var doc = PresentationDocument.Open(file, false);
        var slide = SlidesInOrder(doc.PresentationPart!)[slideIndex - 1];
        var shape = ShapeById(slide, shapeId);
        return shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.GetFirstChild<A.HyperlinkOnClick>();
    }

    // ---- external url -------------------------------------------------------

    [Fact]
    public void SetExternalUrl_AddsRelationshipAndReportsBackViaGet()
    {
        var file = Create();
        var shapePath = AddShape();
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create("https://example.com/q3")))));

        var hlink = Hlink(file, 1, IdOf(shapePath));
        Assert.NotNull(hlink);
        Assert.Null(hlink!.Action?.Value); // an external url carries no ppaction
        Assert.False(string.IsNullOrEmpty(hlink.Id?.Value));

        Assert.Equal("https://example.com/q3", Get(shapePath)["hyperlink"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("ftp://example.com")]
    [InlineData("not a url")]
    public void SetExternalUrl_RejectsUnsupportedScheme(string url)
    {
        Create();
        var shapePath = AddShape();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create(url))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- slide jump ---------------------------------------------------------

    [Fact]
    public void SetSlideJump_PointsRelationshipAtTargetSlide()
    {
        var file = Create();
        AddSlides(3); // a 4-slide deck
        var shapePath = AddShape();
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create("#slide:4")))));

        var hlink = Hlink(file, 1, IdOf(shapePath));
        Assert.NotNull(hlink);
        Assert.Equal("ppaction://hlinksldjump", hlink!.Action?.Value);
        Assert.False(string.IsNullOrEmpty(hlink.Id?.Value));

        // The relationship (an internal part relationship) resolves to slide 4's part.
        using (var doc = PresentationDocument.Open(file, false))
        {
            var slides = SlidesInOrder(doc.PresentationPart!);
            var slide1 = slides[0];
            var slide4 = slides[3];
            var targetPart = slide1.Parts.First(p => p.RelationshipId == hlink.Id!.Value!).OpenXmlPart;
            Assert.Equal(slide4.Uri, targetPart.Uri);
        }

        Assert.Equal("#slide:4", Get(shapePath)["hyperlink"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetSlideJump_OutOfRange_IsRejected()
    {
        Create();
        AddSlides(1); // 2 slides
        var shapePath = AddShape();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create("#slide:9"))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- show actions -------------------------------------------------------

    [Theory]
    [InlineData("#first", "ppaction://hlinkshowjump?jump=firstslide")]
    [InlineData("#last", "ppaction://hlinkshowjump?jump=lastslide")]
    [InlineData("#next", "ppaction://hlinkshowjump?jump=nextslide")]
    [InlineData("#prev", "ppaction://hlinkshowjump?jump=previousslide")]
    [InlineData("#end", "ppaction://hlinkshowjump?jump=endshow")]
    public void SetShowAction_WritesPpactionAndReportsShorthand(string shorthand, string expectedAction)
    {
        var file = Create();
        var shapePath = AddShape();
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create(shorthand)))));

        var hlink = Hlink(file, 1, IdOf(shapePath));
        Assert.NotNull(hlink);
        Assert.Equal(expectedAction, hlink!.Action?.Value);

        Assert.Equal(shorthand, Get(shapePath)["hyperlink"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetUnknownAction_IsRejected()
    {
        Create();
        var shapePath = AddShape();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create("#middle"))))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- clearing -----------------------------------------------------------

    [Fact]
    public void ClearHyperlink_RemovesHlinkAndDropsRelationship()
    {
        var file = Create();
        var shapePath = AddShape();
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create("https://example.com")))));

        var relId = Hlink(file, 1, IdOf(shapePath))!.Id!.Value!;
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create("")))));

        Assert.Null(Hlink(file, 1, IdOf(shapePath)));
        using (var doc = PresentationDocument.Open(file, false))
        {
            var slide1 = SlidesInOrder(doc.PresentationPart!)[0];
            Assert.DoesNotContain(slide1.HyperlinkRelationships, r => r.Id == relId);
        }

        // get reports no hyperlink (the field is omitted as null on the wire).
        Assert.False(Get(shapePath).ContainsKey("hyperlink"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ResetHyperlink_DropsOldSlideJumpRelationship()
    {
        var file = Create();
        AddSlides(3);
        var shapePath = AddShape();
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create("#slide:3")))));
        var firstRelId = Hlink(file, 1, IdOf(shapePath))!.Id!.Value!;

        // Re-set to an external url: the old slide-jump relationship must be gone.
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create("https://example.com")))));
        using (var doc = PresentationDocument.Open(file, false))
        {
            var slide1 = SlidesInOrder(doc.PresentationPart!)[0];
            Assert.DoesNotContain(slide1.Parts, p => p.RelationshipId == firstRelId);
        }

        Assert.Equal("https://example.com", Get(shapePath)["hyperlink"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- query --------------------------------------------------------------

    [Fact]
    public void Query_FindsLinkedShapesByPresenceAndValue()
    {
        Create();
        var linked = AddShape("Linked");
        var plain = AddShape("Plain");
        Edit(TestEnv.Op("set", linked, props: TestEnv.Props(("hyperlink", JsonValue.Create("https://example.com/a")))));

        // hyperlink=* matches the linked shape only.
        var present = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape[hyperlink=*]"))));
        Assert.Equal(1, present["count"]!.GetValue<int>());
        Assert.Contains("@id=" + IdOf(linked), present["matches"]![0]!["path"]!.GetValue<string>());

        // A value containment match works too; the plain shape is excluded.
        var byValue = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape[hyperlink*=example.com]"))));
        Assert.Equal(1, byValue["count"]!.GetValue<int>());

        // hyperlink!=* finds the unlinked shape.
        var absent = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape[hyperlink!=*]"))));
        Assert.Contains(absent["matches"]!.AsArray(), m => m!["path"]!.GetValue<string>().Contains("@id=" + IdOf(plain)));
    }

    // ---- render -------------------------------------------------------------

    [Fact]
    public void Render_MarksLinkedShapes()
    {
        Create();
        var shapePath = AddShape("Go");
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(("hyperlink", JsonValue.Create("https://example.com/x")))));

        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();
        Assert.Contains("data-aio-hyperlink=\"https://example.com/x\"", svg);
        Assert.Contains("<title>link: https://example.com/x</title>", svg);
    }

    // ---- run-level link sugar ----------------------------------------------

    [Fact]
    public void LinkText_WrapsRunsInHyperlink()
    {
        var file = Create();
        var shapePath = AddShape("Visit");
        Edit(TestEnv.Op("set", shapePath, props: TestEnv.Props(
            ("hyperlink", JsonValue.Create("https://example.com")),
            ("linkText", JsonValue.Create("https://example.com")))));

        using (var doc = PresentationDocument.Open(file, false))
        {
            var slide = SlidesInOrder(doc.PresentationPart!)[0];
            var run = ShapeById(slide, IdOf(shapePath)).TextBody!.Descendants<A.Run>().First();
            Assert.NotNull(run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
