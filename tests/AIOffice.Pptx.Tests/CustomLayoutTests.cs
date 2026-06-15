using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.5.0 custom slide layouts (deepening M6 master editing): add a new slideLayout
/// part on the master with specified placeholder shapes (correct ph type/idx), then
/// bind a slide to it by name. Covers add-with-placeholders, reopen-verify of the
/// placeholders, get /master[m]/layout[@name=...] reporting them, a slide using the
/// layout by name, remove (when unreferenced), validation errors and validator-clean
/// — all platform-independent.
/// </summary>
public sealed class CustomLayoutTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() =>
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Decked"))));

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private static JsonNode S(string value) => JsonValue.Create(value)!;

    /// <summary>Builds a placeholders array: {type,x,y,w,h} entries.</summary>
    private static JsonArray Placeholders(params (string Type, string X, string Y, string W, string H)[] entries)
    {
        var array = new JsonArray();
        foreach (var (type, x, y, w, h) in entries)
        {
            array.Add(new JsonObject
            {
                ["type"] = type,
                ["x"] = x,
                ["y"] = y,
                ["w"] = w,
                ["h"] = h,
            });
        }

        return array;
    }

    /// <summary>Adds a custom layout named "My Layout" with title/body/pic/chart/table placeholders.</summary>
    private string AddMyLayout()
    {
        var data = Edit(TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(
            ("name", S("My Layout")),
            ("placeholders", Placeholders(
                ("title", "2cm", "1cm", "28cm", "3cm"),
                ("body", "2cm", "5cm", "13cm", "10cm"),
                ("pic", "16cm", "5cm", "13cm", "5cm"),
                ("chart", "16cm", "11cm", "13cm", "4cm"),
                ("table", "2cm", "16cm", "28cm", "3cm"))))));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    // ---- add ----------------------------------------------------------------

    [Fact]
    public void AddLayout_WithPlaceholders_CreatesLayoutPart_ReopenVerified()
    {
        Create();
        var canonical = AddMyLayout();
        Assert.Equal("/master[1]/layout[2]", canonical);

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var master = doc.PresentationPart!.SlideMasterParts.Single();
        var layout = master.SlideLayoutParts.Single(l => l.SlideLayout!.CommonSlideData!.Name!.Value == "My Layout");

        var placeholders = layout.SlideLayout!.CommonSlideData!.ShapeTree!.Elements<P.Shape>()
            .Where(s => s.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape is not null)
            .ToList();
        Assert.Equal(5, placeholders.Count);

        // The title placeholder has no idx; the others get distinct idx values.
        var title = placeholders[0].NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape!;
        Assert.Equal(P.PlaceholderValues.Title, title.Type!.Value);
        Assert.Null(title.Index);

        var idxs = placeholders.Skip(1)
            .Select(p => p.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape!.Index!.Value)
            .ToList();
        Assert.Equal(idxs.Count, idxs.Distinct().Count());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddLayout_PlaceholderTypesMapToCorrectOoxmlTypes()
    {
        Create();
        AddMyLayout();

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var layout = doc.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts
            .Single(l => l.SlideLayout!.CommonSlideData!.Name!.Value == "My Layout");
        var types = layout.SlideLayout!.CommonSlideData!.ShapeTree!.Elements<P.Shape>()
            .Select(s => s.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape)
            .Where(ph => ph is not null)
            .Select(ph => ph!.Type?.InnerText ?? "body")
            .ToList();

        Assert.Equal(new[] { "title", "body", "pic", "chart", "tbl" }, types);
    }

    // ---- get reports the placeholders ---------------------------------------

    [Fact]
    public void GetLayoutByName_ReportsPlaceholders()
    {
        Create();
        AddMyLayout();

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/layout[@name=My Layout]"))));
        Assert.Equal("My Layout", detail["name"]!.GetValue<string>());

        var placeholders = detail["placeholders"]!.AsArray();
        Assert.Equal(5, placeholders.Count);
        Assert.Equal("title", placeholders[0]!["type"]!.GetValue<string>());
        Assert.Equal(2.0, placeholders[0]!["x"]!.GetValue<double>(), 1);
    }

    [Fact]
    public void GetLayoutByName_UnknownName_IsInvalidPathWithCandidates()
    {
        Create();
        AddMyLayout();

        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/layout[@name=Nope]")));
        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.Contains("/master[1]/layout[@name=My Layout]", error.Candidates!);
    }

    // ---- a slide using the layout by name -----------------------------------

    [Fact]
    public void AddSlide_WithLayoutName_BindsThatLayout_ReopenVerified()
    {
        Create();
        AddMyLayout();

        Edit(TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("layoutName", S("My Layout")))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var presentation = doc.PresentationPart!;
            var slideIds = presentation.Presentation!.SlideIdList!.Elements<P.SlideId>().ToList();
            var newSlide = (SlidePart)presentation.GetPartById(slideIds[1].RelationshipId!.Value!);
            Assert.Equal("My Layout", newSlide.SlideLayoutPart!.SlideLayout!.CommonSlideData!.Name!.Value);
        }

        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        Assert.Equal("/master[1]/layout[2]", structure["slides"]!.AsArray()[1]!["layout"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddSlide_WithUnknownLayoutName_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("layoutName", S("Ghost")))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- remove -------------------------------------------------------------

    [Fact]
    public void RemoveCustomLayout_WhenUnreferenced_Succeeds_ReopenVerified()
    {
        Create();
        AddMyLayout();

        Edit(TestEnv.Op("remove", "/master[1]/layout[2]"));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var master = doc.PresentationPart!.SlideMasterParts.Single();
        Assert.DoesNotContain(master.SlideLayoutParts, l => l.SlideLayout!.CommonSlideData!.Name!.Value == "My Layout");
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveCustomLayout_WhenSlideUsesIt_IsInvalidArgs()
    {
        Create();
        AddMyLayout();
        Edit(TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("layoutName", S("My Layout")))));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/master[1]/layout[2]")]);
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- validation errors --------------------------------------------------

    [Fact]
    public void AddLayout_UnknownPlaceholderType_IsUnsupportedWithCandidates()
    {
        Create();
        var bad = new JsonArray { new JsonObject { ["type"] = "sidebar", ["x"] = "2cm", ["y"] = "2cm", ["w"] = "10cm", ["h"] = "5cm" } };
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(("name", S("Bad")), ("placeholders", bad))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("body", error.Candidates!);
    }

    [Fact]
    public void AddLayout_DuplicateName_IsInvalidArgs()
    {
        Create();
        AddMyLayout();

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(
                ("name", S("My Layout")),
                ("placeholders", Placeholders(("title", "2cm", "1cm", "28cm", "3cm"))))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void AddLayout_BasedOnCombinedWithPlaceholders_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(
                ("name", S("Mixed")),
                ("basedOn", JsonValue.Create(1)),
                ("placeholders", Placeholders(("title", "2cm", "1cm", "28cm", "3cm"))))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void AddLayout_TwoTitles_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(
                ("name", S("TwoTitles")),
                ("placeholders", Placeholders(
                    ("title", "2cm", "1cm", "28cm", "3cm"),
                    ("title", "2cm", "5cm", "28cm", "3cm"))))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- clone-from still works (basedOn, no placeholders) ------------------

    [Fact]
    public void AddLayout_BasedOn_StillClonesAsBefore()
    {
        Create();
        Edit(TestEnv.Op("add", "/master[1]", type: "layout", props: TestEnv.Props(
            ("name", S("Cloned")),
            ("basedOn", JsonValue.Create(1)))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        Assert.Contains(
            doc.PresentationPart!.SlideMasterParts.Single().SlideLayoutParts,
            l => l.SlideLayout!.CommonSlideData!.Name!.Value == "Cloned");
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
