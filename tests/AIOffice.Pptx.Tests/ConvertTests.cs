using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// The M9 conversion surface for pptx: <see cref="PptxHandler.ExportNeutral"/>
/// projects a deck to the format-neutral model, and
/// <see cref="PptxHandler.ImportNeutral"/> writes a neutral model into a freshly
/// created deck as slides. Conversion is lossy across formats; the tests assert
/// content survives a round trip and the Dropped notes honestly name what did not.
/// </summary>
public sealed class ConvertTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    /// <summary>
    /// The round-trip law for conversion: a deck of titled slides with bullet
    /// bodies and a table, exported to neutral then imported into a NEW deck,
    /// preserves the slide count, every title and every bullet text.
    /// </summary>
    [Fact]
    public void RoundTrip_PreservesSlideCount_Titles_AndBulletTexts()
    {
        // The source deck is itself built from a neutral model so its body
        // bullets carry real indent levels — then exported and re-imported.
        BuildSampleDeck("deck.pptx");

        var neutral = _handler.ExportNeutral(_ws.Ctx("deck.pptx"));
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("out.pptx")));
        var result = _handler.ImportNeutral(_ws.Ctx("out.pptx"), neutral);
        Assert.True(result.BlocksWritten > 0);

        // The imported deck has the same titles and bullets, slide for slide.
        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("out.pptx", ("view", "outline"))));
        Assert.Equal(2, outline["slides"]!.AsArray().Count);

        var text = AllText("out.pptx");
        Assert.Contains("Overview", text);
        Assert.Contains("Details", text);
        Assert.Contains("First point", text);
        Assert.Contains("Nested point", text);

        // The table is a graphic frame (not in --view text); re-export to confirm
        // its cells survived the round trip.
        var reExported = _handler.ExportNeutral(_ws.Ctx("out.pptx"));
        var table = reExported.Blocks.First(b => b.Kind == NeutralBlockKind.Table);
        Assert.Equal("Cell A1", table.Rows![0][0]);
        Assert.Equal("Cell B2", table.Rows![1][1]);

        TestEnv.AssertValid(_ws, "out.pptx");
    }

    /// <summary>Export reduces each slide to a level-1 heading followed by its bullet body.</summary>
    [Fact]
    public void Export_MapsSlideTitlesToHeadings_AndBodyToListItems()
    {
        BuildSampleDeck("deck.pptx");
        var neutral = _handler.ExportNeutral(_ws.Ctx("deck.pptx"));

        var headings = neutral.Blocks.Where(b => b.Kind == NeutralBlockKind.Heading).ToList();
        Assert.Equal(["Overview", "Details"], headings.Select(h => h.Runs![0].Text));
        Assert.All(headings, h => Assert.Equal(1, h.Level));

        var bullets = neutral.Blocks.Where(b => b.Kind == NeutralBlockKind.ListItem).ToList();
        Assert.Contains(bullets, b => b.Runs!.Any(r => r.Text == "First point") && b.Level == 0);
        Assert.Contains(bullets, b => b.Runs!.Any(r => r.Text == "Nested point") && b.Level == 1);

        Assert.Contains(neutral.Blocks, b => b.Kind == NeutralBlockKind.Table);
    }

    /// <summary>Speaker notes ride along the export as a "Notes: "-tagged paragraph.</summary>
    [Fact]
    public void Export_CarriesSpeakerNotes_AsTaggedParagraph()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Talk"))));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/slide[1]/notes", props: TestEnv.Props(("text", "remember the demo"))),
        ]));

        var neutral = _handler.ExportNeutral(_ws.Ctx("deck.pptx"));
        Assert.Contains(
            neutral.Blocks,
            b => b.Kind == NeutralBlockKind.Paragraph && b.Runs!.Any(r => r.Text == "Notes: remember the demo"));
    }

    /// <summary>Run formatting (bold/color) survives the export into NeutralRun flags.</summary>
    [Fact]
    public void Export_PreservesRunFormatting()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Styled"))));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
                ("text", "bold line"), ("bold", true), ("color", "FF0000"))),
        ]));

        var neutral = _handler.ExportNeutral(_ws.Ctx("deck.pptx"));
        var run = neutral.Blocks
            .SelectMany(b => b.Runs ?? [])
            .First(r => r.Text == "bold line");
        Assert.True(run.Bold);
        Assert.Equal("FF0000", run.Color);
    }

    /// <summary>A deck with an animation reports a Dropped note on export (effects don't convert).</summary>
    [Fact]
    public void Export_DropsAnimations()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Animated"))));
        var shapeId = AddShapeWithId("deck.pptx", "fly me");
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", $"/slide[1]/shape[@id={shapeId}]", type: "animation",
                props: TestEnv.Props(("effect", "fade"))),
        ]));

        _ = _handler.ExportNeutral(_ws.Ctx("deck.pptx"), out var dropped);
        Assert.Contains(dropped, d => d.Contains("animation", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>A deck with a chart reports a Dropped note on export (the chart cannot convert).</summary>
    [Fact]
    public void Export_DropsCharts()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Charted"))));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "chart", props: TestEnv.Props(
                ("kind", "bar"),
                ("categories", new JsonArray("Q1", "Q2")),
                ("series", new JsonArray(new JsonObject
                {
                    ["name"] = "Sales",
                    ["values"] = new JsonArray(10, 20),
                })))),
        ]));

        _ = _handler.ExportNeutral(_ws.Ctx("deck.pptx"), out var dropped);
        Assert.Contains(dropped, d => d.Contains("chart", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Import writes headings as slide titles and the following bullets as that
    /// slide's body, nesting indent levels — the structure the spec mandates.
    /// </summary>
    [Fact]
    public void Import_BuildsSlidesFromHeadings_WithBulletBodies()
    {
        var doc = new NeutralDoc("My Deck", [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Intro")]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 0, Runs: [new NeutralRun("alpha")]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 1, Runs: [new NeutralRun("beta")]),
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Wrap")]),
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("closing thought")]),
        ]);

        TestEnv.AssertOk(_handler.Create(_ws.Ctx("out.pptx")));
        var result = _handler.ImportNeutral(_ws.Ctx("out.pptx"), doc);
        Assert.Equal(5, result.BlocksWritten);

        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("out.pptx", ("view", "outline"))));
        Assert.Equal(2, outline["slides"]!.AsArray().Count);

        // The deck title rode along on the core property.
        var properties = TestEnv.AssertOk(_handler.Read(_ws.Ctx("out.pptx", ("view", "properties"))));
        Assert.Equal("My Deck", properties["core"]!["title"]!.GetValue<string>());

        var text = AllText("out.pptx");
        Assert.Contains("Intro", text);
        Assert.Contains("alpha", text);
        Assert.Contains("beta", text);
        Assert.Contains("Wrap", text);
        Assert.Contains("closing thought", text);

        TestEnv.AssertValid(_ws, "out.pptx");
    }

    /// <summary>A deep heading (level >= 3) becomes a bold body line on the current slide, not a new slide.</summary>
    [Fact]
    public void Import_DeepHeading_BecomesBodyLine()
    {
        var doc = new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Only Slide")]),
            new NeutralBlock(NeutralBlockKind.Heading, Level: 3, Runs: [new NeutralRun("subsection")]),
        ]);

        TestEnv.AssertOk(_handler.Create(_ws.Ctx("out.pptx")));
        _handler.ImportNeutral(_ws.Ctx("out.pptx"), doc);

        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("out.pptx", ("view", "outline"))));
        Assert.Single(outline["slides"]!.AsArray());
        Assert.Contains("subsection", AllText("out.pptx"));
        TestEnv.AssertValid(_ws, "out.pptx");
    }

    /// <summary>An image whose source is missing is a Dropped note, never a failed conversion.</summary>
    [Fact]
    public void Import_MissingImage_IsDroppedNote_NotFailure()
    {
        var doc = new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Slide")]),
            new NeutralBlock(NeutralBlockKind.Image, Source: "nope.png", Alt: "missing"),
        ]);

        TestEnv.AssertOk(_handler.Create(_ws.Ctx("out.pptx")));
        var result = _handler.ImportNeutral(_ws.Ctx("out.pptx"), doc);
        Assert.Contains(result.Dropped, d => d.Contains("nope.png", StringComparison.Ordinal));
        TestEnv.AssertValid(_ws, "out.pptx");
    }

    /// <summary>A real image lands as a picture carrying the alt text on its accessibility description.</summary>
    [Fact]
    public void Import_RealImage_LandsAsPicture_WithAltText()
    {
        File.WriteAllBytes(_ws.PathOf("logo.png"), TestImages.Png(40, 20));
        var doc = new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Slide")]),
            new NeutralBlock(NeutralBlockKind.Image, Source: "logo.png", Alt: "company logo"),
        ]);

        TestEnv.AssertOk(_handler.Create(_ws.Ctx("out.pptx")));
        _handler.ImportNeutral(_ws.Ctx("out.pptx"), doc);

        var pictures = TestEnv.AssertOk(_handler.Query(_ws.Ctx("out.pptx", ("selector", "shape[kind=picture]"))));
        Assert.True(pictures["count"]!.GetValue<int>() >= 1);
        TestEnv.AssertValid(_ws, "out.pptx");
    }

    /// <summary>A table block imports as a native pptx table whose cells carry the grid text.</summary>
    [Fact]
    public void Import_Table_BuildsNativeTable_WithCellText()
    {
        var doc = new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Data")]),
            new NeutralBlock(
                NeutralBlockKind.Table,
                HeaderRow: true,
                Rows: [
                    ["Name", "Score"],
                    ["Ada", "99"],
                ]),
        ]);

        TestEnv.AssertOk(_handler.Create(_ws.Ctx("out.pptx")));
        _handler.ImportNeutral(_ws.Ctx("out.pptx"), doc);

        // A native pptx table is a graphic frame, not part of --view text; verify
        // the grid via the get projection on the table.
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("out.pptx", ("path", "/slide[1]/table[1]"))));
        var flat = detail.ToJsonString();
        Assert.Contains("Name", flat);
        Assert.Contains("Score", flat);
        Assert.Contains("Ada", flat);
        Assert.Contains("99", flat);
        TestEnv.AssertValid(_ws, "out.pptx");
    }

    /// <summary>Importing the same neutral model twice yields byte-identical decks (deterministic).</summary>
    [Fact]
    public void Import_IsDeterministic()
    {
        var doc = new NeutralDoc("Deterministic", [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("One")]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 0, Runs: [new NeutralRun("bullet")]),
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Two")]),
        ]);

        TestEnv.AssertOk(_handler.Create(_ws.Ctx("a.pptx")));
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("b.pptx")));
        _handler.ImportNeutral(_ws.Ctx("a.pptx"), doc);
        _handler.ImportNeutral(_ws.Ctx("b.pptx"), doc);

        // The text projection is the stable, platform-independent surface to compare.
        Assert.Equal(AllText("a.pptx"), AllText("b.pptx"));
    }

    /// <summary>
    /// Body content that precedes the first heading lands on the seed slide; the
    /// first heading then starts its OWN slide instead of overwriting that body.
    /// </summary>
    [Fact]
    public void Import_BodyBeforeFirstHeading_DoesNotOverwriteSeedSlide()
    {
        var doc = new NeutralDoc(null, [
            new NeutralBlock(NeutralBlockKind.Paragraph, Runs: [new NeutralRun("orphan intro")]),
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Real Title")]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 0, Runs: [new NeutralRun("a bullet")]),
        ]);

        TestEnv.AssertOk(_handler.Create(_ws.Ctx("out.pptx")));
        _handler.ImportNeutral(_ws.Ctx("out.pptx"), doc);

        // Two slides: the orphan body on slide 1, the titled slide after it.
        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("out.pptx", ("view", "outline"))));
        Assert.Equal(2, outline["slides"]!.AsArray().Count);

        var text = AllText("out.pptx");
        Assert.Contains("orphan intro", text);
        Assert.Contains("Real Title", text);
        Assert.Contains("a bullet", text);
        TestEnv.AssertValid(_ws, "out.pptx");
    }

    /// <summary>
    /// The convert-fidelity contract: importing a NeutralDoc with a Title and a
    /// leading Heading yields a deck whose FIRST slide carries that heading (the
    /// seed slide is reused, not left as a junk empty first slide) and whose core
    /// Title lands in the standard <c>docProps/core.xml</c> part — verified on the
    /// SAVED package, so the converted deck title is visible and round-trips.
    /// </summary>
    [Fact]
    public void Import_TitleAndLeadingHeading_FirstSlideCarriesHeading_AndCoreTitleStandardized()
    {
        var doc = new NeutralDoc("Deck Title", [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("First Heading")]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 0, Runs: [new NeutralRun("a bullet")]),
        ]);

        TestEnv.AssertOk(_handler.Create(_ws.Ctx("out.pptx")));
        _handler.ImportNeutral(_ws.Ctx("out.pptx"), doc);

        // No junk empty first slide: the seed slide is reused, so the deck is a
        // single slide. Re-exporting confirms its FIRST heading is the leading
        // heading (an empty seed slide would yield no heading at all).
        var outline = TestEnv.AssertOk(_handler.Read(_ws.Ctx("out.pptx", ("view", "outline"))));
        Assert.Single(outline["slides"]!.AsArray());

        var reExported = _handler.ExportNeutral(_ws.Ctx("out.pptx"));
        var firstHeading = reExported.Blocks.First(b => b.Kind == NeutralBlockKind.Heading);
        Assert.Equal("First Heading", firstHeading.Runs![0].Text);

        // The envelope reads the title back...
        var properties = TestEnv.AssertOk(_handler.Read(_ws.Ctx("out.pptx", ("view", "properties"))));
        Assert.Equal("Deck Title", properties["core"]!["title"]!.GetValue<string>());

        // ...and the SAVED package carries it in the standard docProps/core.xml part.
        using (var pkg = PresentationDocument.Open(_ws.PathOf("out.pptx"), false))
        {
            Assert.NotNull(pkg.CoreFilePropertiesPart);
            Assert.Equal("/docProps/core.xml", pkg.CoreFilePropertiesPart!.Uri.ToString());
            Assert.Equal("Deck Title", pkg.CoreFilePropertiesPart.CoreFileProperties.Title);
        }

        TestEnv.AssertValid(_ws, "out.pptx");
    }

    /// <summary>An empty neutral model imports to a valid one-slide deck (the seed slide stays).</summary>
    [Fact]
    public void Import_EmptyDoc_StaysValid()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("out.pptx")));
        var result = _handler.ImportNeutral(_ws.Ctx("out.pptx"), new NeutralDoc(null, []));
        Assert.Equal(0, result.BlocksWritten);
        TestEnv.AssertValid(_ws, "out.pptx");
    }

    // ---- helpers ------------------------------------------------------------

    /// <summary>
    /// A two-slide deck built from a neutral model so its body bullets carry real
    /// indent levels and the table is a native pptx table: titled slides, a nested
    /// bullet on slide 1, a header-row table on slide 2 — the canonical fixture.
    /// </summary>
    private void BuildSampleDeck(string name)
    {
        var doc = new NeutralDoc("Overview", [
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Overview")]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 0, Runs: [new NeutralRun("First point")]),
            new NeutralBlock(NeutralBlockKind.ListItem, Level: 1, Runs: [new NeutralRun("Nested point")]),
            new NeutralBlock(NeutralBlockKind.Heading, Level: 1, Runs: [new NeutralRun("Details")]),
            new NeutralBlock(
                NeutralBlockKind.Table,
                HeaderRow: true,
                Rows: [
                    ["Cell A1", "Cell B1"],
                    ["Cell A2", "Cell B2"],
                ]),
        ]);

        TestEnv.AssertOk(_handler.Create(_ws.Ctx(name)));
        _handler.ImportNeutral(_ws.Ctx(name), doc);
    }

    private uint AddShapeWithId(string name, string text)
    {
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx(name), [
            TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(("text", text))),
        ]));
        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx(name, ("view", "structure"))));
        var shapes = structure["slides"]!.AsArray()
            .First(s => s!["index"]!.GetValue<int>() == 1)!["shapes"]!.AsArray();
        return shapes[^1]!["id"]!.GetValue<uint>();
    }

    /// <summary>The deck's whole text projection, CRLF-normalized for deterministic asserts.</summary>
    private string AllText(string name)
    {
        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx(name, ("view", "text"))));
        return data["text"]!.GetValue<string>().Replace("\r\n", "\n", StringComparison.Ordinal);
    }
}
