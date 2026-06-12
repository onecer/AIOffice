using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

public sealed class CreateTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    [Fact]
    public void Create_ProducesValidatorCleanDeckWithOneSlide()
    {
        var envelope = _handler.Create(_ws.Ctx("deck.pptx"));
        var data = TestEnv.AssertOk(envelope);

        Assert.True(File.Exists(_ws.PathOf("deck.pptx")));
        Assert.Equal(1, data["slides"]!.GetValue<int>());
        Assert.Equal("pptx", data["kind"]!.GetValue<string>());
        Assert.NotNull(envelope.Meta.Rev);
        Assert.Equal(Rev.Length, envelope.Meta.Rev!.Length);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Create_BuildsTheRequiredPartGraph()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var presentation = doc.PresentationPart;
        Assert.NotNull(presentation);
        Assert.Single(presentation!.SlideParts);

        var master = Assert.Single(presentation.SlideMasterParts);
        var layout = Assert.Single(master.SlideLayoutParts);
        Assert.NotNull(master.ThemePart);
        Assert.NotNull(layout.SlideMasterPart);

        var slideIds = presentation.Presentation!.SlideIdList!.Elements<P.SlideId>().ToList();
        var slideId = Assert.Single(slideIds);
        Assert.True(slideId.Id!.Value >= 256U);
    }

    [Fact]
    public void Create_WithTitle_PutsTheTitleOnSlideOne()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Quarterly Review"))));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))));
        Assert.Contains("Quarterly Review", data["text"]!.GetValue<string>(), StringComparison.Ordinal);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Create_RefusesToOverwriteAnExistingFile()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        var envelope = _handler.Create(_ws.Ctx("deck.pptx"));
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Equal(ExitCodes.UserError, envelope.ExitCode);
        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    [Fact]
    public void Read_OnMissingFile_IsTypedFileNotFound()
    {
        var envelope = _handler.Read(_ws.Ctx("nope.pptx"));
        TestEnv.AssertFail(envelope, ErrorCodes.FileNotFound);
    }
}
