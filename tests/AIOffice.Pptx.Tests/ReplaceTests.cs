using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>The shared M4 find/replace contract: slide, shape and notes scopes.</summary>
public sealed class ReplaceTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    /// <summary>deck.pptx with one slide and one textbox; returns the shape's canonical path.</summary>
    private string CreateWithText(string text)
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", text))),
        ]));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private Envelope Replace(string path, params (string Key, JsonNode? Value)[] props) =>
        _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("replace", path, props: TestEnv.Props(props))]);

    private string ShapeText(string shapePath)
    {
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", shapePath))));
        return detail["text"]!.GetValue<string>();
    }

    [Fact]
    public void Replace_SlideScope_CountsAndLocates()
    {
        var shapePath = CreateWithText("Q3 plan\nQ3 risks and Q3 wins");

        var data = TestEnv.AssertOk(Replace("/slide[1]", ("find", "Q3"), ("replace", "Q4")));
        var result = data["results"]![0]!.AsObject();
        Assert.Equal("/slide[1]", result["target"]!.GetValue<string>());
        Assert.Equal(3, result["replacements"]!.GetValue<int>());
        Assert.Equal(
            [$"{shapePath}/p[1]", $"{shapePath}/p[2]"],
            result["locations"]!.AsArray().Select(l => l!.GetValue<string>()));

        Assert.Equal("Q4 plan\nQ4 risks and Q4 wins", ShapeText(shapePath));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Replace_ShapeScope_TouchesOnlyThatShape()
    {
        var first = CreateWithText("alpha beta");
        var added = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "alpha gamma"))),
        ]));
        var second = added["results"]![0]!["target"]!.GetValue<string>();

        var data = TestEnv.AssertOk(Replace(second, ("find", "alpha"), ("replace", "omega")));
        Assert.Equal(1, data["results"]![0]!["replacements"]!.GetValue<int>());

        Assert.Equal("alpha beta", ShapeText(first));
        Assert.Equal("omega gamma", ShapeText(second));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Replace_FindSplitAcrossRuns_KeepsFirstRunFormatting()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", "placeholder"))),
        ]));

        // Hand-build a paragraph whose match spans three differently formatted runs:
        // "We ship " + [bold red]"Pro" + [italic]"ject" + " Phoenix today".
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true))
        {
            var shape = doc.PresentationPart!.SlideParts.Single()
                .Slide!.Descendants<P.Shape>().Single(s => s.InnerText.Contains("placeholder", StringComparison.Ordinal));
            var paragraph = shape.TextBody!.Elements<A.Paragraph>().First();
            paragraph.RemoveAllChildren();
            paragraph.Append(new A.Run(new A.Text("We ship ")));
            paragraph.Append(new A.Run(
                new A.RunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })) { Bold = true },
                new A.Text("Pro")));
            paragraph.Append(new A.Run(new A.RunProperties { Italic = true }, new A.Text("ject")));
            paragraph.Append(new A.Run(new A.Text(" Phoenix today")));
        }

        var data = TestEnv.AssertOk(Replace("/slide[1]", ("find", "Project Phoenix"), ("replace", "Operation Kestrel")));
        Assert.Equal(1, data["results"]![0]!["replacements"]!.GetValue<int>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var shape = doc.PresentationPart!.SlideParts.Single()
                .Slide!.Descendants<P.Shape>().Single(s => s.InnerText.Contains("Kestrel", StringComparison.Ordinal));
            var paragraph = shape.TextBody!.Elements<A.Paragraph>().First();
            Assert.Equal("We ship Operation Kestrel today", string.Concat(
                paragraph.Elements<A.Run>().Select(r => r.Text!.Text)));

            // The replacement inherits the formatting of the first affected run (bold red).
            var replaced = paragraph.Elements<A.Run>().Single(r => r.Text!.Text == "Operation Kestrel");
            Assert.True(replaced.RunProperties!.Bold!.Value);
            Assert.Equal("FF0000", replaced.RunProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);

            // The unmatched head run before the match is untouched ("We ship " has no run props).
            Assert.Equal("We ship ", paragraph.Elements<A.Run>().First().Text!.Text);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Replace_SlideScope_IncludesNotesByDefault()
    {
        CreateWithText("Q3 body");
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("text", "mention Q3 twice: Q3"))),
        ]));

        var data = TestEnv.AssertOk(Replace("/slide[1]", ("find", "Q3"), ("replace", "Q4")));
        var result = data["results"]![0]!.AsObject();
        Assert.Equal(3, result["replacements"]!.GetValue<int>());
        Assert.Contains(
            "/slide[1]/notes",
            result["locations"]!.AsArray().Select(l => l!.GetValue<string>()));

        var notes = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/notes"))));
        Assert.Equal("mention Q4 twice: Q4", notes["text"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Replace_IncludeNotesFalse_LeavesNotesAlone()
    {
        CreateWithText("Q3 body");
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("text", "notes Q3"))),
        ]));

        var data = TestEnv.AssertOk(Replace(
            "/slide[1]", ("find", "Q3"), ("replace", "Q4"), ("includeNotes", JsonValue.Create(false))));
        Assert.Equal(1, data["results"]![0]!["replacements"]!.GetValue<int>());

        var notes = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/notes"))));
        Assert.Equal("notes Q3", notes["text"]!.GetValue<string>());
    }

    [Fact]
    public void Replace_NotesScope_TouchesOnlyNotes()
    {
        var shapePath = CreateWithText("Q3 body");
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("text", "notes Q3"))),
        ]));

        var data = TestEnv.AssertOk(Replace("/slide[1]/notes", ("find", "Q3"), ("replace", "Q4")));
        var result = data["results"]![0]!.AsObject();
        Assert.Equal("/slide[1]/notes", result["target"]!.GetValue<string>());
        Assert.Equal(1, result["replacements"]!.GetValue<int>());
        Assert.Equal("Q3 body", ShapeText(shapePath));
    }

    [Fact]
    public void Replace_NoMatch_IsOkWithFindNoMatchWarning()
    {
        CreateWithText("nothing to see");

        var envelope = Replace("/slide[1]", ("find", "unicorn"), ("replace", "horse"));
        var data = TestEnv.AssertOk(envelope);
        Assert.Equal(0, data["results"]![0]!["replacements"]!.GetValue<int>());
        Assert.Empty(data["results"]![0]!["locations"]!.AsArray());

        var warning = Assert.Single(envelope.Meta.Warnings!);
        Assert.Equal("find_no_match", warning.Code);
    }

    [Fact]
    public void Replace_MatchCase_And_WholeWord_AreHonored()
    {
        var shapePath = CreateWithText("Cat cat catalog");

        TestEnv.AssertOk(Replace(
            "/slide[1]",
            ("find", "cat"), ("replace", "dog"),
            ("matchCase", JsonValue.Create(true)), ("wholeWord", JsonValue.Create(true))));

        Assert.Equal("Cat dog catalog", ShapeText(shapePath));
    }

    [Fact]
    public void Replace_Regex_WithGroupSubstitution()
    {
        var shapePath = CreateWithText("Revenue: 120 in FY24 and 150 in FY25");

        var data = TestEnv.AssertOk(Replace(
            "/slide[1]",
            ("find", @"FY(\d\d)"), ("replace", "fiscal 20$1"),
            ("regex", JsonValue.Create(true))));
        Assert.Equal(2, data["results"]![0]!["replacements"]!.GetValue<int>());
        Assert.Equal("Revenue: 120 in fiscal 2024 and 150 in fiscal 2025", ShapeText(shapePath));
    }

    [Fact]
    public void Replace_InvalidRegex_IsInvalidArgs()
    {
        CreateWithText("text");
        var envelope = Replace("/slide[1]", ("find", "(unclosed"), ("regex", JsonValue.Create(true)));
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Replace_MissingFind_IsInvalidArgs()
    {
        CreateWithText("text");
        var envelope = Replace("/slide[1]", ("replace", "x"));
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Replace_OnChartPath_IsInvalidArgs()
    {
        CreateWithText("text");
        var envelope = Replace("/slide[1]/chart[1]", ("find", "a"), ("replace", "b"));
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    /// <summary>The wire path: --ops JSON with the M4 op kind and paths parses through Core.</summary>
    [Fact]
    public void ParseBatch_AcceptsReplaceOp_AndM4Paths()
    {
        var ops = EditOp.ParseBatch(
            "[{\"op\":\"replace\",\"path\":\"/slide[1]\",\"props\":{\"find\":\"a\",\"replace\":\"b\"}}," +
            "{\"op\":\"remove\",\"path\":\"/slide[1]/animation[2]\"}," +
            "{\"op\":\"remove\",\"path\":\"/slide[1]/comment[@id=3]\"}]");

        Assert.Equal(3, ops.Count);
        Assert.Equal("replace", ops[0].Op);
        Assert.Equal("/slide[1]/animation[2]", ops[1].Path);
        Assert.Equal("/slide[1]/comment[@id=3]", ops[2].Path);
    }

    [Fact]
    public void Replace_FailedOpInBatch_LeavesFileUntouched()
    {
        CreateWithText("Q3 once");
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("replace", "/slide[1]", props: TestEnv.Props(("find", "Q3"), ("replace", "Q4"))),
            TestEnv.Op("replace", "/slide[9]", props: TestEnv.Props(("find", "Q3"), ("replace", "Q4"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }
}
