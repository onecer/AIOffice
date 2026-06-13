using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

public sealed class AltTextTests : IDisposable
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

    private string AddShape(string text = "Hi")
    {
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("text", JsonValue.Create(text)))));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    [Fact]
    public void SetAltText_RoundTripsThroughGet()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("altText", JsonValue.Create("Sales chart Q3")),
            ("altTitle", JsonValue.Create("Q3 sales")))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("Sales chart Q3", detail["altText"]!.GetValue<string>());
        Assert.Equal("Q3 sales", detail["altTitle"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetAltText_WritesCNvPrDescr()
    {
        var file = Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("altText", JsonValue.Create("A picture")))));

        using var doc = PresentationDocument.Open(file, false);
        var shape = doc.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<P.Shape>().First(s => s.TextBody is not null);
        Assert.Equal("A picture", shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Description!.Value);
    }

    [Fact]
    public void EmptyAltText_ClearsTheDescription()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("altText", JsonValue.Create("temp")))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("altText", JsonValue.Create("")))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Null(detail["altText"]);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RenderSvg_EmitsTitleFromAltText()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("altText", JsonValue.Create("Logo image")))));

        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();
        Assert.Contains("<title>Logo image</title>", svg);
    }

    [Fact]
    public void ReadingOrderMove_ReordersDocumentOrder()
    {
        Create();
        var first = AddShape("First");
        var second = AddShape("Second");

        // Move the second shape to be narrated first.
        Edit(TestEnv.Op("move", second, position: "readingOrder 1"));

        // Document order is the narration order; the moved shape is now first.
        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var shapes = structure["slides"]![0]!["shapes"]!.AsArray();
        Assert.Equal(second, shapes[0]!["path"]!.GetValue<string>());
        Assert.Equal(first, shapes[1]!["path"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ReadingOrderMove_OutOfRange_Fails()
    {
        Create();
        var path = AddShape();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("move", path, position: "readingOrder 9")]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void ReadingOrderMove_NonNumeric_Fails()
    {
        Create();
        var path = AddShape();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("move", path, position: "readingOrder x")]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }
}
