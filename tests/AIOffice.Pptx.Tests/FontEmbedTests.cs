using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.5.0 embedded fonts: a font file embedded into the deck as a FontPart and
/// registered in p:embeddedFontLst, so the deck renders without the font being
/// installed. Covers add (ttf + otf), the embeddedFontLst registration, /fonts
/// get + per-font get, remove, sandbox denial on an escaping src, reopen-verify,
/// the round-trip law and validator-clean — all platform-independent.
/// </summary>
public sealed class FontEmbedTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() =>
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Branded"))));

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private string WriteFont(string name, byte[] bytes)
    {
        File.WriteAllBytes(_ws.PathOf(name), bytes);
        return name;
    }

    private static JsonNode S(string value) => JsonValue.Create(value)!;

    // ---- add ----------------------------------------------------------------

    [Fact]
    public void AddFont_EmbedsFontPart_AndRegistersInEmbeddedFontLst_ReopenVerified()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());

        var data = Edit(TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(
            ("src", S("Brand.ttf")),
            ("name", S("Brand Sans")))));
        Assert.Equal("/fonts/font[@name=Brand Sans]", data["results"]![0]!["target"]!.GetValue<string>());

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var presentation = doc.PresentationPart!;

        // The font part exists and carries the verbatim bytes.
        var fontPart = Assert.Single(presentation.FontParts);
        using (var stream = fontPart.GetStream())
        {
            Assert.Equal(TestFonts.Ttf().Length, stream.Length);
        }

        // It is registered in p:embeddedFontLst with the regular slot pointing at it.
        var list = presentation.Presentation!.EmbeddedFontList!;
        var embedded = Assert.Single(list.Elements<P.EmbeddedFont>());
        Assert.Equal("Brand Sans", embedded.Font!.Typeface!.Value);
        Assert.Equal(presentation.GetIdOfPart(fontPart), embedded.RegularFont!.Id!.Value);
        Assert.True(presentation.Presentation.EmbedTrueTypeFonts!.Value);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddFont_OtfSource_IsEmbedded_AndValid()
    {
        Create();
        WriteFont("Display.otf", TestFonts.Otf());

        Edit(TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Display.otf")))));

        // The name defaults to the file name without extension.
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/fonts/font[@name=Display]"))));
        Assert.Equal("Display", detail["name"]!.GetValue<string>());

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        Assert.Single(doc.PresentationPart!.FontParts);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddFont_EmbedAll_WithAllFourFiles_EmbedsAllFourSlots_ReopenVerified()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());
        WriteFont("Brand-Bold.ttf", TestFonts.Ttf(80));
        WriteFont("Brand-Italic.ttf", TestFonts.Ttf(96));
        WriteFont("Brand-BoldItalic.ttf", TestFonts.Ttf(112));

        Edit(TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(
            ("src", S("Brand.ttf")),
            ("name", S("Brand")),
            ("embedAll", JsonValue.Create(true)),
            ("bold", S("Brand-Bold.ttf")),
            ("italic", S("Brand-Italic.ttf")),
            ("boldItalic", S("Brand-BoldItalic.ttf")))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var presentation = doc.PresentationPart!;
        var embedded = Assert.Single(presentation.Presentation!.EmbeddedFontList!.Elements<P.EmbeddedFont>());

        // embedAll fills regular/bold/italic/boldItalic; each references a font part.
        Assert.NotNull(embedded.RegularFont);
        Assert.NotNull(embedded.BoldFont);
        Assert.NotNull(embedded.ItalicFont);
        Assert.NotNull(embedded.BoldItalicFont);
        Assert.Equal(4, presentation.FontParts.Count());

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/fonts/font[@name=Brand]"))));
        Assert.Equal(new[] { "regular", "bold", "italic", "boldItalic" },
            detail["styles"]!.AsArray().Select(n => n!.GetValue<string>()).ToArray());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddFont_PerStyleFilesWithoutEmbedAll_EmbedsJustThoseSlots()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());
        WriteFont("Brand-Bold.ttf", TestFonts.Ttf(80));

        // No embedAll: regular + the one explicit bold file, nothing fabricated.
        Edit(TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(
            ("src", S("Brand.ttf")),
            ("name", S("Brand")),
            ("bold", S("Brand-Bold.ttf")))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var presentation = doc.PresentationPart!;
        var embedded = Assert.Single(presentation.Presentation!.EmbeddedFontList!.Elements<P.EmbeddedFont>());
        Assert.NotNull(embedded.RegularFont);
        Assert.NotNull(embedded.BoldFont);
        Assert.Null(embedded.ItalicFont);
        Assert.Equal(2, presentation.FontParts.Count());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddFont_EmbedAllWithoutStyleFiles_IsInvalidArgs()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());

        // embedAll with no bold/italic/boldItalic files: honest refusal, no fabrication.
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(
                ("src", S("Brand.ttf")),
                ("name", S("Brand")),
                ("embedAll", JsonValue.Create(true)))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("embedAll", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AddFont_MissingSrc_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("name", S("Nope")))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("src", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AddFont_DuplicateName_IsInvalidArgs()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());
        Edit(TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Brand.ttf")), ("name", S("Brand")))));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Brand.ttf")), ("name", S("Brand")))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void AddFont_NonFontFile_IsInvalidArgs()
    {
        Create();
        WriteFont("notafont.ttf", "this is plain text, not a font"u8.ToArray());

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("notafont.ttf")))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("font", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ---- sandbox ------------------------------------------------------------

    [Fact]
    public void AddFont_SrcOutsideWorkspace_IsSandboxDenied_AndWritesNothing()
    {
        Create();
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("../escapee.ttf")))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.SandboxDenied);
        Assert.Equal(ExitCodes.SandboxDenied, envelope.ExitCode);
        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    // ---- get / list ---------------------------------------------------------

    [Fact]
    public void GetFonts_ListsEveryEmbeddedFont_WithStylesAndSize()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());
        WriteFont("Display.otf", TestFonts.Otf());
        Edit(
            TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Brand.ttf")), ("name", S("Brand")))),
            TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Display.otf")), ("name", S("Display")))));

        var list = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/fonts"))));
        Assert.Equal("/fonts", list["path"]!.GetValue<string>());
        Assert.Equal(2, list["count"]!.GetValue<int>());
        Assert.True(list["embedTrueTypeFonts"]!.GetValue<bool>());

        var first = list["fonts"]!.AsArray()[0]!;
        Assert.Equal("/fonts/font[@name=Brand]", first["path"]!.GetValue<string>());
        Assert.Equal("Brand", first["name"]!.GetValue<string>());
        Assert.Equal(new[] { "regular" }, first["styles"]!.AsArray().Select(n => n!.GetValue<string>()).ToArray());
        Assert.True(first["size"]!.GetValue<long>() > 0);
    }

    [Fact]
    public void GetFont_ByName_ReportsTheFont()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());
        Edit(TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Brand.ttf")), ("name", S("Brand")))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/fonts/font[@name=Brand]"))));
        Assert.Equal("Brand", detail["name"]!.GetValue<string>());
        Assert.Equal("/fonts/font[@name=Brand]", detail["path"]!.GetValue<string>());
    }

    [Fact]
    public void GetFont_UnknownName_IsInvalidPathWithCandidates()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());
        Edit(TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Brand.ttf")), ("name", S("Brand")))));

        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/fonts/font[@name=Missing]")));
        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.Contains("/fonts/font[@name=Brand]", error.Candidates!);
    }

    // ---- remove -------------------------------------------------------------

    [Fact]
    public void RemoveFont_DropsRegistrationAndPart_ReopenVerified()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());
        Edit(TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Brand.ttf")), ("name", S("Brand")))));

        var data = Edit(TestEnv.Op("remove", "/fonts/font[@name=Brand]"));
        Assert.Equal("/fonts/font[@name=Brand]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var presentation = doc.PresentationPart!;
            Assert.Empty(presentation.FontParts); // the orphaned part is dropped
            // The empty list is removed, and the flag cleared with the last font gone.
            Assert.Null(presentation.Presentation!.EmbeddedFontList);
            Assert.Null(presentation.Presentation.EmbedTrueTypeFonts);
        }

        var list = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/fonts"))));
        Assert.Equal(0, list["count"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveFont_KeepsOtherFonts()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());
        WriteFont("Display.otf", TestFonts.Otf());
        Edit(
            TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Brand.ttf")), ("name", S("Brand")))),
            TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Display.otf")), ("name", S("Display")))));

        Edit(TestEnv.Op("remove", "/fonts/font[@name=Brand]"));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var presentation = doc.PresentationPart!;
        var remaining = Assert.Single(presentation.Presentation!.EmbeddedFontList!.Elements<P.EmbeddedFont>());
        Assert.Equal("Display", remaining.Font!.Typeface!.Value);
        Assert.Single(presentation.FontParts);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveFont_WholeList_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/fonts")]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- set is unsupported (additive in-place edits are not offered) --------

    [Fact]
    public void SetFonts_IsUnsupported()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/fonts", props: TestEnv.Props(("name", S("x")))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
    }

    // ---- round-trip law ------------------------------------------------------

    [Fact]
    public void GetFonts_LeavesEveryByteUntouched()
    {
        Create();
        WriteFont("Brand.ttf", TestFonts.Ttf());
        Edit(TestEnv.Op("add", "/fonts", type: "font", props: TestEnv.Props(("src", S("Brand.ttf")), ("name", S("Brand")))));
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/fonts"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/fonts/font[@name=Brand]"))));

        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }
}
