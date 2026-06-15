using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.5.0 action buttons: navigation shapes built on M8 shape hyperlinks. Each is a
/// p:sp with a preset action-button geometry plus the matching ppaction click
/// action. Covers add (every kind: show-jumps, slide-jump, url), reopen-verify of
/// the prstGeom + ppaction, get reporting the action, remove, the svg glyph, a
/// bad-action unsupported error, the round-trip law and validator-clean — all
/// platform-independent.
/// </summary>
public sealed class ActionButtonTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() =>
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Nav"))));

    private void AddSlides(int count)
    {
        for (var i = 0; i < count; i++)
        {
            Edit(TestEnv.Op("add", $"/slide[{i + 2}]", type: "slide"));
        }
    }

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private static JsonNode S(string value) => JsonValue.Create(value)!;

    /// <summary>The added action button's canonical shape path.</summary>
    private string AddButton(params (string Key, JsonNode? Value)[] props)
    {
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "actionButton", props: TestEnv.Props(props)));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    // ---- add: each show-jump kind ------------------------------------------

    [Theory]
    [InlineData("first", "actionButtonBeginning", "firstslide")]
    [InlineData("last", "actionButtonEnd", "lastslide")]
    [InlineData("next", "actionButtonForwardNext", "nextslide")]
    [InlineData("prev", "actionButtonBackPrevious", "previousslide")]
    [InlineData("home", "actionButtonHome", "firstslide")]
    [InlineData("end", "actionButtonReturn", "lastslide")]
    public void AddActionButton_ShowJump_HasPrstGeomAndPpaction_ReopenVerified(string action, string expectedGeometry, string expectedJump)
    {
        Create();
        var canonical = AddButton(("action", S(action)));
        Assert.Matches(@"^/slide\[1\]/shape\[@id=[0-9]+\]$", canonical);

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var shape = doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>()
            .Single(s => s.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText?.StartsWith("actionButton", StringComparison.Ordinal) == true);

        Assert.Equal(expectedGeometry, shape.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.Preset!.InnerText);
        var hlink = shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.GetFirstChild<A.HyperlinkOnClick>()!;
        Assert.Equal("ppaction://hlinkshowjump?jump=" + expectedJump, hlink.Action!.Value);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddActionButton_SlideJump_RelatesToTargetSlide_ReopenVerified()
    {
        Create();
        AddSlides(2); // deck now has 3 slides

        var canonical = AddButton(("action", S("slide")), ("target", S("slide 3")));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var presentation = doc.PresentationPart!;
        var slides = presentation.Presentation!.SlideIdList!.Elements<P.SlideId>()
            .Select(id => (SlidePart)presentation.GetPartById(id.RelationshipId!.Value!)).ToList();
        var slide1 = slides[0];

        var shape = slide1.Slide!.Descendants<P.Shape>()
            .Single(s => s.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText == "actionButtonBlank");
        var hlink = shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.GetFirstChild<A.HyperlinkOnClick>()!;
        Assert.Equal("ppaction://hlinksldjump", hlink.Action!.Value);

        // The relationship target is slide 3.
        var target = slide1.GetPartById(hlink.Id!.Value!);
        Assert.Equal(slides[2].Uri, ((SlidePart)target).Uri);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddActionButton_Url_AddsExternalRelationship_ReopenVerified()
    {
        Create();
        var canonical = AddButton(("action", S("url")), ("target", S("https://example.com/docs")));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var slidePart = doc.PresentationPart!.SlideParts.Single();
        var shape = slidePart.Slide!.Descendants<P.Shape>()
            .Single(s => s.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText == "actionButtonInformation");
        var hlink = shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.GetFirstChild<A.HyperlinkOnClick>()!;
        Assert.Equal("ppaction://hlinkfile", hlink.Action!.Value);

        var rel = slidePart.HyperlinkRelationships.Single(r => r.Id == hlink.Id!.Value);
        Assert.Equal("https://example.com/docs", rel.Uri.OriginalString);
        Assert.True(rel.IsExternal);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddActionButton_WithLabelAndGeometry_HonorsBoth()
    {
        Create();
        var canonical = AddButton(
            ("action", S("home")),
            ("label", S("Home")),
            ("x", S("3cm")), ("y", S("17cm")), ("w", S("2cm")), ("h", S("2cm")),
            ("fill", S("2563EB")));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal("Home", detail["text"]!.GetValue<string>());
        Assert.Equal("2563EB", detail["fill"]!.GetValue<string>());
        Assert.Equal(3.0, detail["x"]!.GetValue<double>(), 1);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- get reports the action ---------------------------------------------

    [Fact]
    public void GetActionButton_ReportsActionAndGeometry()
    {
        Create();
        AddSlides(1);
        var canonical = AddButton(("action", S("slide")), ("target", S("slide 2")));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        var button = detail["actionButton"]!;
        Assert.Equal("actionButtonBlank", button["geometry"]!.GetValue<string>());
        Assert.Equal("slide", button["action"]!.GetValue<string>());
        Assert.Equal("slide 2", button["target"]!.GetValue<string>());
    }

    [Fact]
    public void GetActionButton_ShowJump_ReportsActionWithNoTarget()
    {
        Create();
        var canonical = AddButton(("action", S("next")));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        var button = detail["actionButton"]!;
        Assert.Equal("next", button["action"]!.GetValue<string>());
        Assert.Null(button["target"]);
    }

    [Fact]
    public void GetPlainShape_HasNoActionButton()
    {
        Create();
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", S("hi")))));
        var canonical = data["results"]![0]!["target"]!.GetValue<string>();

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Null(detail["actionButton"]);
    }

    // ---- remove -------------------------------------------------------------

    [Fact]
    public void RemoveActionButton_DropsTheShape_ReopenVerified()
    {
        Create();
        var canonical = AddButton(("action", S("first")));

        Edit(TestEnv.Op("remove", canonical));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        Assert.DoesNotContain(
            doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>(),
            s => s.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText?.StartsWith("actionButton", StringComparison.Ordinal) == true);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- bad input ----------------------------------------------------------

    [Fact]
    public void AddActionButton_UnknownAction_IsUnsupportedWithCandidates()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "actionButton", props: TestEnv.Props(("action", S("teleport")))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("first", error.Candidates!);
        Assert.Contains("url", error.Candidates!);
    }

    [Fact]
    public void AddActionButton_MissingAction_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "actionButton", props: TestEnv.Props(("label", S("x")))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void AddActionButton_SlideTargetOutOfRange_IsInvalidArgs_AndWritesNothing()
    {
        Create();
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "actionButton", props: TestEnv.Props(("action", S("slide")), ("target", S("slide 9")))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    [Fact]
    public void AddActionButton_UrlWithoutAbsoluteTarget_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "actionButton", props: TestEnv.Props(("action", S("url")), ("target", S("/relative")))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void AddActionButton_ShowJumpWithTarget_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "actionButton", props: TestEnv.Props(("action", S("next")), ("target", S("slide 2")))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- svg render ---------------------------------------------------------

    [Fact]
    public void Render_ActionButton_DrawsGlyph()
    {
        Create();
        AddButton(("action", S("next")));

        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();
        Assert.Contains("aio-action-button", svg, StringComparison.Ordinal);
        Assert.Contains("►", svg, StringComparison.Ordinal); // the "next" glyph
    }

    // ---- round-trip law ------------------------------------------------------

    [Fact]
    public void GetAndRender_LeaveEveryByteUntouched()
    {
        Create();
        var canonical = AddButton(("action", S("home")));
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));

        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }
}
