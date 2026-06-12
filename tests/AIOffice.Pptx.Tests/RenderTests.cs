using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Pptx.Tests;

public sealed class RenderTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void CreateDeck()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Cover"))));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide"),
            TestEnv.Op("add", "/slide[2]", type: "shape", props: TestEnv.Props(
                ("x", JsonValue.Create(2)),
                ("y", JsonValue.Create(3)),
                ("w", JsonValue.Create(10)),
                ("h", JsonValue.Create(2)),
                ("text", JsonValue.Create("Revenue & growth")),
                ("bold", JsonValue.Create(true)),
                ("fill", JsonValue.Create("FFEE00")))),
        ]));
    }

    [Fact]
    public void Svg_PlacesShapeRectAndTextAtScaledCoordinates()
    {
        CreateDeck();
        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[2]"))));

        Assert.Equal(1, data["slideCount"]!.GetValue<int>());
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();

        // 2cm = 720000 EMU = 75.6px at 96dpi; 3cm = 113.4px.
        Assert.Contains("x=\"75.6\"", svg, StringComparison.Ordinal);
        Assert.Contains("y=\"113.4\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#FFEE00\"", svg, StringComparison.Ordinal);
        Assert.Contains("Revenue &amp; growth", svg, StringComparison.Ordinal);
        Assert.Contains("font-weight=\"bold\"", svg, StringComparison.Ordinal);
        Assert.Contains("<g data-aio-path=\"/slide[2]/shape[@id=", svg, StringComparison.Ordinal);
    }

    /// <summary>The data-aio-path render contract: every shape is wrapped in a g
    /// whose attribute is the canonical stable-id path, so a browser click maps
    /// back to an addressable node.</summary>
    [Fact]
    public void Svg_WrapsEveryShapeInADataAioPathGroup()
    {
        CreateDeck();
        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"))));

        var slides = data["slides"]!.AsArray();
        Assert.Equal(2, slides.Count);
        foreach (var slide in slides)
        {
            var path = slide!["path"]!.GetValue<string>();
            var svg = slide["svg"]!.GetValue<string>();
            var groups = svg.Split("<g data-aio-path=\"").Length - 1;
            var shapeCount = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))))["shapeCount"]!.GetValue<int>();

            Assert.Equal(shapeCount, groups);
            Assert.Equal(groups, svg.Split("</g>").Length - 1);
            for (var ordinal = 1; ordinal <= shapeCount; ordinal++)
            {
                var canonical = TestEnv.AssertOk(_handler.Get(
                    _ws.Ctx("deck.pptx", ("path", $"{path}/shape[{ordinal}]"))))["path"]!.GetValue<string>();
                Assert.Contains($"<g data-aio-path=\"{canonical}\"", svg, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void Html_CarriesDataAioPathAttributes()
    {
        CreateDeck();
        var html = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "html"))))["html"]!.GetValue<string>();

        Assert.Contains("<g data-aio-path=\"/slide[1]/shape[@id=", html, StringComparison.Ordinal);
        Assert.Contains("<g data-aio-path=\"/slide[2]/shape[@id=", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_WrapsEverySlideSvg()
    {
        CreateDeck();
        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "html"))));

        var html = data["html"]!.GetValue<string>();
        Assert.StartsWith("<!DOCTYPE html>", html, StringComparison.Ordinal);
        Assert.Equal(2, html.Split("<svg ").Length - 1);
        Assert.Contains("Cover", html, StringComparison.Ordinal);
    }

    [Fact]
    public void TextRender_MatchesReadTextView()
    {
        CreateDeck();
        var rendered = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "text"))))["text"]!.GetValue<string>();
        var read = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))))["text"]!.GetValue<string>();

        Assert.Equal(read, rendered);
    }

    [Fact]
    public void Png_IsTypedUnsupportedWithWorkaround()
    {
        CreateDeck();
        var envelope = _handler.Render(_ws.Ctx("deck.pptx", ("to", "png")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Equal(ExitCodes.UnsupportedFeature, envelope.ExitCode);
        Assert.Contains("svg", error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Output_WritesScopedSvgInsideTheSandbox()
    {
        CreateDeck();
        var data = TestEnv.AssertOk(_handler.Render(
            _ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"), ("output", "out/slide1.svg"))));

        var output = data["output"]!.GetValue<string>();
        Assert.True(File.Exists(output));
        Assert.Contains("Cover", File.ReadAllText(output), StringComparison.Ordinal);
    }

    [Fact]
    public void Output_MultiSlideSvg_IsRejectedWithGuidance()
    {
        CreateDeck();
        var envelope = _handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("output", "all.svg")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("--scope", error.Suggestion, StringComparison.Ordinal);
    }
}
