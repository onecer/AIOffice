using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

public sealed class EditTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private string Create(string? title = null)
    {
        TestEnv.AssertOk(_handler.Create(
            title is null ? _ws.Ctx("deck.pptx") : _ws.Ctx("deck.pptx", ("title", title))));
        return _ws.PathOf("deck.pptx");
    }

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    [Fact]
    public void AddSlide_AppendsAndStaysValid()
    {
        Create();
        var data = Edit(TestEnv.Op("add", "/slide[2]", type: "slide"));

        Assert.Equal(2, data["slides"]!.GetValue<int>());
        Assert.Equal("/slide[2]", data["results"]![0]!["target"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddSlide_WithTitle_InsertsAtRequestedPosition()
    {
        Create("First");
        Edit(TestEnv.Op("add", "/slide[1]", type: "slide", props: TestEnv.Props(("title", "Zeroth"))));

        var text = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))))["text"]!.GetValue<string>();
        Assert.True(
            text.IndexOf("Zeroth", StringComparison.Ordinal) < text.IndexOf("First", StringComparison.Ordinal),
            "inserted slide should come first: " + text);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddShape_RoundTripsGeometryTextAndFont()
    {
        var file = Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("x", JsonValue.Create(5)),
            ("y", JsonValue.Create("3cm")),
            ("w", JsonValue.Create(10)),
            ("h", JsonValue.Create("1440000emu")),
            ("text", JsonValue.Create("Hello Q3")),
            ("fontSize", JsonValue.Create(24)),
            ("bold", JsonValue.Create(true)),
            ("color", JsonValue.Create("FF0000")),
            ("fill", JsonValue.Create("#FFEE00")),
            ("align", JsonValue.Create("center")),
            ("name", JsonValue.Create("Box1")))));

        using (var doc = PresentationDocument.Open(file, false))
        {
            var shape = doc.PresentationPart!.SlideParts.Single().Slide!
                .Descendants<P.Shape>()
                .Single(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Box1");

            var transform = shape.ShapeProperties!.Transform2D!;
            Assert.Equal(1_800_000L, transform.Offset!.X!.Value);
            Assert.Equal(1_080_000L, transform.Offset!.Y!.Value);
            Assert.Equal(3_600_000L, transform.Extents!.Cx!.Value);
            Assert.Equal(1_440_000L, transform.Extents!.Cy!.Value);

            Assert.Equal("FFEE00", shape.ShapeProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);

            var paragraph = shape.TextBody!.Elements<A.Paragraph>().Single();
            Assert.Equal(A.TextAlignmentTypeValues.Center, paragraph.ParagraphProperties!.Alignment!.Value);

            var run = paragraph.Elements<A.Run>().Single();
            Assert.Equal("Hello Q3", run.Text!.Text);
            Assert.Equal(2400, run.RunProperties!.FontSize!.Value);
            Assert.True(run.RunProperties.Bold!.Value);
            Assert.Equal("FF0000", run.RunProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetShape_UpdatesTextGeometryAndFill()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "old"))));
        var canonical = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", canonical, props: TestEnv.Props(
            ("text", JsonValue.Create("new text")),
            ("x", JsonValue.Create(7.5)),
            ("fill", JsonValue.Create("00FF00")))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal("new text", detail["text"]!.GetValue<string>());
        Assert.Equal(7.5, detail["x"]!.GetValue<double>());
        Assert.Equal("00FF00", detail["fill"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetParagraph_RewritesOneParagraphOnly()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "line one\nline two"))));

        Edit(TestEnv.Op("set", "/slide[1]/shape[1]/p[2]", props: TestEnv.Props(("text", "LINE TWO"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/shape[1]"))));
        Assert.Equal("line one\nLINE TWO", detail["text"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveShape_ByStableId()
    {
        Create();
        var added = Edit(
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "keep me"))),
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "drop me"))));
        var dropPath = added["results"]![1]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("remove", dropPath));

        var query = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape"))));
        Assert.Equal(1, query["count"]!.GetValue<int>());
        Assert.Equal("keep me", query["matches"]![0]!["text"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveSlide_DeletesThePart()
    {
        Create("Doomed");
        Edit(TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("title", "Survivor"))));

        var data = Edit(TestEnv.Op("remove", "/slide[1]"));
        Assert.Equal(1, data["slides"]!.GetValue<int>());

        var text = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))))["text"]!.GetValue<string>();
        Assert.DoesNotContain("Doomed", text, StringComparison.Ordinal);
        Assert.Contains("Survivor", text, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MoveSlide_ReordersTheDeck()
    {
        Create("A");
        Edit(
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("title", "B"))),
            TestEnv.Op("add", "/slide[3]", type: "slide", props: TestEnv.Props(("title", "C"))));

        var data = Edit(TestEnv.Op("move", "/slide[3]", position: "1"));
        Assert.Equal("/slide[1]", data["results"]![0]!["target"]!.GetValue<string>());

        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        var slides = outline["slides"]!.AsArray();
        Assert.Equal("C", slides[0]!["shapes"]![0]!["text"]!.GetValue<string>());
        Assert.Equal("A", slides[1]!["shapes"]![0]!["text"]!.GetValue<string>());
        Assert.Equal("B", slides[2]!["shapes"]![0]!["text"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void MoveSlide_SupportsBeforeAnchors()
    {
        Create("A");
        Edit(
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("title", "B"))),
            TestEnv.Op("add", "/slide[3]", type: "slide", props: TestEnv.Props(("title", "C"))));

        Edit(TestEnv.Op("move", "/slide[1]", position: "before:/slide[3]"));

        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "outline"))));
        var slides = outline["slides"]!.AsArray();
        Assert.Equal("B", slides[0]!["shapes"]![0]!["text"]!.GetValue<string>());
        Assert.Equal("A", slides[1]!["shapes"]![0]!["text"]!.GetValue<string>());
        Assert.Equal("C", slides[2]!["shapes"]![0]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void Edit_IsAtomic_NoWriteWhenAnyOpFails()
    {
        var file = Create("Untouched");
        var before = File.ReadAllBytes(file);

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/slide[1]/shape[1]", props: TestEnv.Props(("text", "changed"))),
            TestEnv.Op("set", "/slide[1]/shape[99]", props: TestEnv.Props(("text", "boom"))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.NotNull(error.Candidates);
        Assert.NotEmpty(error.Candidates!);
        Assert.Equal(before, File.ReadAllBytes(file));
    }

    [Fact]
    public void Edit_ExpectRevMismatch_IsStaleAddressAndNoWrite()
    {
        var file = Create("Untouched");
        var before = File.ReadAllBytes(file);

        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx", ("expectRev", "000000000000")),
            [TestEnv.Op("set", "/slide[1]/shape[1]", props: TestEnv.Props(("text", "changed")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.StaleAddress);
        Assert.Equal(before, File.ReadAllBytes(file));
    }

    [Fact]
    public void Edit_ExpectRevMatch_Succeeds()
    {
        var file = Create("Old Title");
        var rev = Rev.OfFile(file);

        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx", ("expectRev", rev)),
            [TestEnv.Op("set", "/slide[1]/shape[1]", props: TestEnv.Props(("text", "New Title")))]);

        TestEnv.AssertOk(envelope);
        Assert.NotEqual(rev, envelope.Meta.Rev);
    }

    [Fact]
    public void Edit_SnapshotsThePreImage()
    {
        var file = Create("Snapshot me");
        var before = File.ReadAllBytes(file);
        var store = new SnapshotStore(Path.Combine(_ws.Dir, ".snapshots"));
        var handler = new PptxHandler(store);

        TestEnv.AssertOk(handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/slide[1]/shape[1]", props: TestEnv.Props(("text", "edited")))]));

        var entry = Assert.Single(store.List(file));
        Assert.Equal(before, File.ReadAllBytes(entry.Path));
        Assert.NotEqual(before, File.ReadAllBytes(file));
    }

    [Fact]
    public void AddUnsupportedType_IsTypedUnsupportedFeature()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "video")]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Equal(ExitCodes.UnsupportedFeature, envelope.ExitCode);
        Assert.Contains("slide", error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }
}
