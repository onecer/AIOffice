using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// M6 custom slide size — set / {slideSize | width+height} rewrites p:sldSz,
/// leaving every existing shape's EMU coordinates untouched. get / reports the
/// preset name (when matched), the dimensions in cm and the counts.
/// </summary>
public sealed class SlideSizeTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() => TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Cover"))));

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    [Fact]
    public void GetRoot_ReportsDefault16x9()
    {
        Create();
        var data = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/"))));

        Assert.Equal("/", data["path"]!.GetValue<string>());
        Assert.Equal("16:9", data["slideSize"]!.GetValue<string>());
        Assert.Equal(33.87, data["widthCm"]!.GetValue<double>(), 2);
        Assert.Equal(19.05, data["heightCm"]!.GetValue<double>(), 2);
        Assert.Equal(1, data["slideCount"]!.GetValue<int>());
    }

    [Theory]
    [InlineData("4:3", 9_144_000, 6_858_000)]
    [InlineData("16:10", 10_972_800, 6_858_000)]
    [InlineData("A4", 10_692_000, 7_560_000)]
    [InlineData("letter", 9_144_000, 6_858_000)]
    public void SetNamedPreset_RewritesSldSz_ReopenVerified(string preset, int expectedCx, int expectedCy)
    {
        Create();
        var data = Edit(TestEnv.Op("set", "/", props: TestEnv.Props(("slideSize", preset))));
        Assert.Equal("/", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var size = doc.PresentationPart!.Presentation!.SlideSize!;
            Assert.Equal(expectedCx, size.Cx!.Value);
            Assert.Equal(expectedCy, size.Cy!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetExplicitDimensions_RewritesSldSz_AsCustom()
    {
        Create();
        Edit(TestEnv.Op("set", "/", props: TestEnv.Props(("width", "33.87cm"), ("height", "19.05cm"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var size = doc.PresentationPart!.Presentation!.SlideSize!;
            Assert.Equal(P.SlideSizeValues.Custom, size.Type!.Value);
            Assert.Equal(12_193_200, size.Cx!.Value); // 33.87cm in EMU
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetSlideSize_LeavesExistingShapeCoordinatesUntouched()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("text", "fixed"), ("x", "5cm"), ("y", "4cm"), ("w", "8cm"), ("h", "3cm"))));
        var before = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/shape[1]"))));

        Edit(TestEnv.Op("set", "/", props: TestEnv.Props(("slideSize", "4:3"))));

        var after = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/shape[1]"))));
        Assert.Equal(before["x"]!.GetValue<double>(), after["x"]!.GetValue<double>());
        Assert.Equal(before["y"]!.GetValue<double>(), after["y"]!.GetValue<double>());
        Assert.Equal(before["w"]!.GetValue<double>(), after["w"]!.GetValue<double>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void GetRoot_AfterPreset_ReportsTheName()
    {
        Create();
        Edit(TestEnv.Op("set", "/", props: TestEnv.Props(("slideSize", "4:3"))));

        var data = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/"))));
        Assert.Equal("4:3", data["slideSize"]!.GetValue<string>());
    }

    [Fact]
    public void GetRoot_AfterCustom_ReportsNullPresetButDimensions()
    {
        Create();
        Edit(TestEnv.Op("set", "/", props: TestEnv.Props(("width", "30cm"), ("height", "20cm"))));

        var data = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/"))));
        Assert.Null(data["slideSize"]);
        Assert.Equal(30.0, data["widthCm"]!.GetValue<double>(), 2);
        Assert.Equal(20.0, data["heightCm"]!.GetValue<double>(), 2);
    }

    [Fact]
    public void Structure_ReportsSlideSizePreset()
    {
        Create();
        Edit(TestEnv.Op("set", "/", props: TestEnv.Props(("slideSize", "16:10"))));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        Assert.Equal("16:10", data["slideSize"]!.GetValue<string>());
    }

    [Fact]
    public void UnknownPreset_IsInvalidArgsWithCandidates()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/", props: TestEnv.Props(("slideSize", "widescreen"))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("16:9", error.Candidates!);
    }

    [Fact]
    public void SlideSizeAndExplicitDimensions_Combined_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/", props: TestEnv.Props(("slideSize", "4:3"), ("width", "30cm"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void WidthWithoutHeight_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/", props: TestEnv.Props(("width", "30cm"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void UnknownRootProp_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/", props: TestEnv.Props(("background", "FF0000"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void DimensionsOutOfRange_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/", props: TestEnv.Props(("width", "200cm"), ("height", "20cm"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }
}
