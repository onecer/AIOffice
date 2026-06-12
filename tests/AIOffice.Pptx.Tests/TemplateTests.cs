using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

public sealed class TemplateTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void CreateDeckWithPlaceholders()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("text", JsonValue.Create("Hello {{name}}, welcome to {{name}}'s review for {{quarter}}.")))),
        ]));
    }

    [Fact]
    public void Template_ReplacesAllOccurrencesInPlace()
    {
        CreateDeckWithPlaceholders();
        var data = TestEnv.AssertOk(_handler.Template(_ws.Ctx(
            "deck.pptx",
            ("data", new JsonObject { ["name"] = "Dana", ["quarter"] = "Q3" }))));

        Assert.Equal(3, data["replacements"]!.GetValue<int>());

        var text = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))))["text"]!.GetValue<string>();
        Assert.Equal("Hello Dana, welcome to Dana's review for Q3.", text);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Template_MergesPlaceholdersThatSpanRuns()
    {
        CreateDeckWithPlaceholders();

        // Simulate PowerPoint splitting "{{quarter}}" across two runs.
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true))
        {
            var paragraph = doc.PresentationPart!.SlideParts.Single().Slide!
                .Descendants<P.Shape>()
                .Single(s => PptxText(s).Contains("{{quarter}}", StringComparison.Ordinal))
                .TextBody!.Elements<A.Paragraph>().Single();

            var run = paragraph.Elements<A.Run>().Single();
            var full = run.Text!.Text;
            var splitAt = full.IndexOf("{{qua", StringComparison.Ordinal) + 5;
            run.Text.Text = full[..splitAt];

            var second = (A.Run)run.CloneNode(true);
            second.Text!.Text = full[splitAt..];
            paragraph.InsertAfter(second, run);
        }

        var data = TestEnv.AssertOk(_handler.Template(_ws.Ctx(
            "deck.pptx",
            ("data", new JsonObject { ["name"] = "Dana", ["quarter"] = "Q3" }))));

        Assert.Equal(3, data["replacements"]!.GetValue<int>());
        var text = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))))["text"]!.GetValue<string>();
        Assert.Equal("Hello Dana, welcome to Dana's review for Q3.", text);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Template_ToOutputFile_LeavesTheSourceUntouched()
    {
        CreateDeckWithPlaceholders();
        var sourceBytes = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        var data = TestEnv.AssertOk(_handler.Template(_ws.Ctx(
            "deck.pptx",
            ("data", new JsonObject { ["name"] = "Dana", ["quarter"] = "Q3" }),
            ("output", "merged.pptx"))));

        Assert.Equal(sourceBytes, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
        Assert.True(File.Exists(data["output"]!.GetValue<string>()));

        var text = TestEnv.AssertOk(_handler.Read(_ws.Ctx("merged.pptx", ("view", "text"))))["text"]!.GetValue<string>();
        Assert.Contains("Dana", text, StringComparison.Ordinal);
        TestEnv.AssertValid(_ws, "merged.pptx");
    }

    [Fact]
    public void Template_WithoutData_IsInvalidArgs()
    {
        CreateDeckWithPlaceholders();
        var envelope = _handler.Template(_ws.Ctx("deck.pptx"));

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    private static string PptxText(P.Shape shape) =>
        string.Concat(shape.Descendants<A.Text>().Select(t => t.Text));
}
