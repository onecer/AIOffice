using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// M6 master/layout editing (the flagship that pays the M1 read-only debt):
/// master/layout background, theme accent colors, placeholder-shape edits and
/// add/clone/remove layout. Every edit reopens validator-clean and leaves the
/// existing slides valid.
/// </summary>
public sealed class MasterEditTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private string Create(string title = "Cover")
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", title))));
        return _ws.PathOf("deck.pptx");
    }

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    /// <summary>Puts a title placeholder shape onto the master's own shape tree.</summary>
    private void AddMasterTitleShape()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true);
        var master = doc.PresentationPart!.SlideMasterParts.Single();
        master.SlideMaster!.CommonSlideData!.ShapeTree!.Append(new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2U, Name = "Master Title" },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 831_850L, Y = 365_125L },
                    new A.Extents { Cx = 10_515_600L, Cy = 1_325_563L }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text("Master title text"))))));
    }

    // ---- master background ---------------------------------------------------

    [Fact]
    public void SetMasterBackground_WritesBgSolidFill_ReopenVerified()
    {
        Create();
        var data = Edit(TestEnv.Op("set", "/master[1]", props: TestEnv.Props(("background", "0F172A"))));
        Assert.Equal("/master[1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slideData = doc.PresentationPart!.SlideMasterParts.Single().SlideMaster!.CommonSlideData!;
            Assert.Equal(
                "0F172A",
                slideData.Background!.BackgroundProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetLayoutBackground_WritesBgSolidFill_ReopenVerified()
    {
        Create();
        var data = Edit(TestEnv.Op("set", "/master[1]/layout[1]", props: TestEnv.Props(("background", "1E293B"))));
        Assert.Equal("/master[1]/layout[1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slideData = doc.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts.Single()
                .SlideLayout!.CommonSlideData!;
            Assert.Equal(
                "1E293B",
                slideData.Background!.BackgroundProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- theme accent colors -------------------------------------------------

    [Fact]
    public void SetMasterAccents_RecolorsTheThemeColorScheme_ReopenVerified()
    {
        Create();
        Edit(TestEnv.Op("set", "/master[1]", props: TestEnv.Props(
            ("accent1", "38BDF8"), ("accent2", "F472B6"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var scheme = doc.PresentationPart!.SlideMasterParts.Single().ThemePart!.Theme!.ThemeElements!.ColorScheme!;
            Assert.Equal("38BDF8", scheme.Accent1Color!.RgbColorModelHex!.Val!.Value);
            Assert.Equal("F472B6", scheme.Accent2Color!.RgbColorModelHex!.Val!.Value);
            // Untouched accents keep their original values.
            Assert.Equal("A5A5A5", scheme.Accent3Color!.RgbColorModelHex!.Val!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]"))));
        Assert.Equal("/master[1]", detail["path"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetMasterBackgroundAndAccent_InOneOp_BothApply()
    {
        Create();
        Edit(TestEnv.Op("set", "/master[1]", props: TestEnv.Props(
            ("background", "0B1120"), ("accent1", "22D3EE"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var master = doc.PresentationPart!.SlideMasterParts.Single();
        Assert.Equal(
            "0B1120",
            master.SlideMaster!.CommonSlideData!.Background!.BackgroundProperties!.GetFirstChild<A.SolidFill>()!
                .RgbColorModelHex!.Val!.Value);
        Assert.Equal("22D3EE", master.ThemePart!.Theme!.ThemeElements!.ColorScheme!.Accent1Color!.RgbColorModelHex!.Val!.Value);
    }

    [Fact]
    public void UnknownMasterProp_IsInvalidArgsWithCandidates()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/master[1]", props: TestEnv.Props(("transition", "fade"))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("accent1", error.Candidates!);
        Assert.Contains("background", error.Candidates!);
    }

    [Fact]
    public void UnknownLayoutProp_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("set", "/master[1]/layout[1]", props: TestEnv.Props(("accent1", "FF0000"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- placeholder shape edits --------------------------------------------

    [Fact]
    public void SetMasterShape_TextFillAndGeometry_ReopenVerified()
    {
        Create();
        AddMasterTitleShape();

        var data = Edit(
            TestEnv.Op("set", "/master[1]/shape[1]", props: TestEnv.Props(
                ("text", "New master heading"), ("fill", "334155"), ("x", "1cm"), ("y", "1cm"))));
        Assert.Equal("/master[1]/shape[@id=2]", data["results"]![0]!["target"]!.GetValue<string>());

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/shape[1]"))));
        Assert.Equal("New master heading", detail["text"]!.GetValue<string>());
        Assert.Equal("334155", detail["fill"]!.GetValue<string>());
        Assert.Equal(1.0, detail["x"]!.GetValue<double>(), 3);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddShapeToLayout_ThenRemoveIt_StaysValid()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/master[1]/layout[1]", type: "shape", props: TestEnv.Props(
            ("text", "Footer hint"), ("shape", "rect"), ("x", "2cm"), ("y", "17cm"), ("w", "10cm"), ("h", "1cm"))));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();
        Assert.Matches(@"^/master\[1\]/layout\[1\]/shape\[@id=[0-9]+\]$", shapePath);
        TestEnv.AssertValid(_ws, "deck.pptx");

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", shapePath))));
        Assert.Equal("Footer hint", detail["text"]!.GetValue<string>());

        Edit(TestEnv.Op("remove", shapePath));
        TestEnv.AssertFail(_handler.Get(_ws.Ctx("deck.pptx", ("path", shapePath))), ErrorCodes.InvalidPath);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- add / clone / remove layout ----------------------------------------

    [Fact]
    public void AddLayout_ClonesAnExistingOne_AndIsUsableByAddSlide()
    {
        Create();
        var data = Edit(TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(
            ("name", "My Layout"), ("basedOn", JsonValue.Create(1)))));
        Assert.Equal("/master[1]/layout[2]", data["results"]![0]!["target"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");

        var layout = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/layout[2]"))));
        Assert.Equal("My Layout", layout["name"]!.GetValue<string>());

        // A new slide can bind to the cloned layout by index.
        Edit(TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("layout", JsonValue.Create(2)))));
        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        Assert.Equal("/master[1]/layout[2]", structure["slides"]!.AsArray()[1]!["layout"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddLayout_DefaultBasedOn_ClonesTheFirstLayout()
    {
        Create();
        Edit(TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(("name", "Cloned"))));

        var master = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]"))));
        Assert.Equal(2, master["layoutCount"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddLayout_BasedOnOutOfRange_IsInvalidArgsAndNoWrite()
    {
        var file = Create();
        var before = File.ReadAllBytes(file);

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(("basedOn", JsonValue.Create(9)))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("/master[1]/layout[1]", error.Candidates!);
        Assert.Equal(before, File.ReadAllBytes(file));
    }

    [Fact]
    public void RemoveUnreferencedLayout_Succeeds_StaysValid()
    {
        Create();
        // Two layouts; the second is unreferenced (slide 1 uses the first).
        Edit(TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(("name", "Spare"))));
        TestEnv.AssertValid(_ws, "deck.pptx");

        Edit(TestEnv.Op("remove", "/master[1]/layout[2]"));
        var master = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]"))));
        Assert.Equal(1, master["layoutCount"]!.GetValue<int>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveReferencedLayout_IsInvalidArgs_NamingSlides_AndNoWrite()
    {
        var file = Create(); // slide 1 uses layout 1
        var before = File.ReadAllBytes(file);

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("remove", "/master[1]/layout[1]"),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("slide", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("/slide[1]", error.Candidates!);
        Assert.Equal(before, File.ReadAllBytes(file));
    }

    [Fact]
    public void RemoveOnlyLayout_WhenUnreferenced_IsStillInvalidArgs()
    {
        // Remove slide 1 so no slide references layout 1, then try to drop the master's only layout.
        Create();
        Edit(TestEnv.Op("add", "/slide[2]", type: "slide"));
        Edit(TestEnv.Op("remove", "/slide[1]"));
        Edit(TestEnv.Op("remove", "/slide[1]")); // now zero slides

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("remove", "/master[1]/layout[1]"),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("at least one layout", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RemoveMaster_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/master[1]")]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void AddImageToMaster_IsUnsupportedFeature()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/master[1]", type: "image", props: TestEnv.Props(("src", "x.png"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
    }
}
