using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// 1.8 master theme fonts: <c>set /master[m] {majorFont, minorFont}</c> renames
/// the master theme part's font scheme faces (a:majorFont/a:minorFont →
/// a:latin@typeface). This fixes the previously-broken AGENT-GUIDE recipe. The
/// master <c>get</c> reports the two faces back.
/// </summary>
public sealed class MasterFontTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() =>
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Cover"))));

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    [Fact]
    public void SetMasterFonts_WritesThemeFontScheme_RoundTrips()
    {
        Create();
        var data = Edit(TestEnv.Op("set", "/master[1]", props: TestEnv.Props(
            ("majorFont", "Montserrat"), ("minorFont", "Inter"))));
        Assert.Equal("/master[1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var fontScheme = doc.PresentationPart!.SlideMasterParts.Single().ThemePart!.Theme!.ThemeElements!.FontScheme!;
            Assert.Equal("Montserrat", fontScheme.MajorFont!.LatinFont!.Typeface!.Value);
            Assert.Equal("Inter", fontScheme.MinorFont!.LatinFont!.Typeface!.Value);
        }

        // get /master[1] reports the two faces.
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]"))));
        Assert.Equal("Montserrat", detail["majorFont"]!.GetValue<string>());
        Assert.Equal("Inter", detail["minorFont"]!.GetValue<string>());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetMajorFontAlone_LeavesMinorUntouched()
    {
        Create();
        Edit(TestEnv.Op("set", "/master[1]", props: TestEnv.Props(("majorFont", "Playfair Display"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var fontScheme = doc.PresentationPart!.SlideMasterParts.Single().ThemePart!.Theme!.ThemeElements!.FontScheme!;
        Assert.Equal("Playfair Display", fontScheme.MajorFont!.LatinFont!.Typeface!.Value);
        // The default minor face survives (factory seeds Calibri).
        Assert.Equal("Calibri", fontScheme.MinorFont!.LatinFont!.Typeface!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetMasterFonts_DoesNotLoseTheEastAsianAndCsFaces()
    {
        Create();
        Edit(TestEnv.Op("set", "/master[1]", props: TestEnv.Props(("majorFont", "Georgia"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var major = doc.PresentationPart!.SlideMasterParts.Single().ThemePart!.Theme!.ThemeElements!.FontScheme!.MajorFont!;
        Assert.NotNull(major.GetFirstChild<A.EastAsianFont>());
        Assert.NotNull(major.GetFirstChild<A.ComplexScriptFont>());
        // latin is still the first child (correct schema order).
        Assert.IsType<A.LatinFont>(major.FirstChild);
    }

    [Fact]
    public void FontsCombineWithAccentColorsInOneSet()
    {
        Create();
        Edit(TestEnv.Op("set", "/master[1]", props: TestEnv.Props(
            ("majorFont", "Oswald"), ("accent1", "0EA5E9"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var theme = doc.PresentationPart!.SlideMasterParts.Single().ThemePart!.Theme!.ThemeElements!;
        Assert.Equal("Oswald", theme.FontScheme!.MajorFont!.LatinFont!.Typeface!.Value);
        Assert.Equal("0EA5E9", theme.ColorScheme!.Accent1Color!.RgbColorModelHex!.Val!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void EmptyFontName_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(
            _ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", "/master[1]", props: TestEnv.Props(("majorFont", "")))]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }
}
