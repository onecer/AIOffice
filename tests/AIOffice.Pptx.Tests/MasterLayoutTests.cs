using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// M1 master/layout read addressing: /master[1], /master[1]/layout[2] and
/// shapes beneath are gettable/queryable/listable, while every edit on them
/// stays a typed unsupported_feature until M2.
/// </summary>
public sealed class MasterLayoutTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private string Create(string title = "Cover")
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", title))));
        return _ws.PathOf("deck.pptx");
    }

    /// <summary>Adds a second, "Title Only" layout (with a title placeholder shape) to the deck's master.</summary>
    private void AddTitleOnlyLayout()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true);
        var master = doc.PresentationPart!.SlideMasterParts.Single();

        var layout = master.AddNewPart<SlideLayoutPart>();
        layout.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(TreeWith(BuildTitlePlaceholder(2U, "Title Placeholder 1", "Layout title")))
            {
                Name = "Title Only",
            },
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.TitleOnly,
        };
        layout.AddPart(master);

        var idList = master.SlideMaster!.SlideLayoutIdList!;
        var nextId = idList.Elements<P.SlideLayoutId>().Max(i => i.Id!.Value) + 1;
        idList.Append(new P.SlideLayoutId { Id = nextId, RelationshipId = master.GetIdOfPart(layout) });
    }

    /// <summary>Puts a title placeholder shape onto the master's own shape tree.</summary>
    private void AddMasterTitleShape()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true);
        var master = doc.PresentationPart!.SlideMasterParts.Single();
        master.SlideMaster!.CommonSlideData!.ShapeTree!
            .Append(BuildTitlePlaceholder(2U, "Master Title", "Master title text"));
    }

    private static P.ShapeTree TreeWith(params P.Shape[] shapes)
    {
        var tree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = (UInt32Value)1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup()));
        foreach (var shape in shapes)
        {
            tree.Append(shape);
        }

        return tree;
    }

    private static P.Shape BuildTitlePlaceholder(uint id, string name, string text) => new(
        new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = (UInt32Value)id, Name = name },
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
            new A.Paragraph(new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(text)))));

    // ---- get ----------------------------------------------------------------

    [Fact]
    public void Get_Master_ListsItsLayoutsAndShapes()
    {
        Create();
        var data = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]"))));

        Assert.Equal("/master[1]", data["path"]!.GetValue<string>());
        Assert.Equal(1, data["index"]!.GetValue<int>());
        Assert.Equal("master", data["kind"]!.GetValue<string>());
        Assert.Equal(1, data["layoutCount"]!.GetValue<int>());
        Assert.Equal(0, data["shapeCount"]!.GetValue<int>());

        var layout = data["layouts"]!.AsArray()[0]!;
        Assert.Equal("/master[1]/layout[1]", layout["path"]!.GetValue<string>());
        Assert.Equal("Blank", layout["name"]!.GetValue<string>());
        Assert.Equal("blank", layout["type"]!.GetValue<string>());
    }

    [Fact]
    public void Get_Layout_ReportsNameTypeAndWhichSlidesUseIt()
    {
        Create();
        var data = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/layout[1]"))));

        Assert.Equal("/master[1]/layout[1]", data["path"]!.GetValue<string>());
        Assert.Equal(1, data["master"]!.GetValue<int>());
        Assert.Equal("layout", data["kind"]!.GetValue<string>());
        Assert.Equal("Blank", data["name"]!.GetValue<string>());
        Assert.Equal("blank", data["type"]!.GetValue<string>());
        Assert.Equal(new[] { 1 }, data["usedBySlides"]!.AsArray().Select(n => n!.GetValue<int>()).ToArray());
    }

    [Fact]
    public void Get_LayoutShape_ReportsNameKindTextAndPlaceholderType()
    {
        Create();
        AddTitleOnlyLayout();
        TestEnv.AssertValid(_ws, "deck.pptx");

        var data = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/layout[2]/shape[1]"))));

        Assert.Equal("/master[1]/layout[2]/shape[@id=2]", data["path"]!.GetValue<string>());
        Assert.Equal("/master[1]/layout[2]/shape[1]", data["ordinalPath"]!.GetValue<string>());
        Assert.Equal(1, data["master"]!.GetValue<int>());
        Assert.Equal(2, data["layout"]!.GetValue<int>());
        Assert.Equal("shape", data["kind"]!.GetValue<string>());
        Assert.Equal("Title Placeholder 1", data["name"]!.GetValue<string>());
        Assert.Equal("title", data["placeholder"]!.GetValue<string>());
        Assert.Equal("Layout title", data["text"]!.GetValue<string>());
    }

    [Fact]
    public void Get_MasterShape_ResolvesByOrdinalAndStableId()
    {
        Create();
        AddMasterTitleShape();
        TestEnv.AssertValid(_ws, "deck.pptx");

        var byOrdinal = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/shape[1]"))));
        Assert.Equal("/master[1]/shape[@id=2]", byOrdinal["path"]!.GetValue<string>());
        Assert.Equal("title", byOrdinal["placeholder"]!.GetValue<string>());
        Assert.Equal("Master title text", byOrdinal["text"]!.GetValue<string>());

        var byId = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/shape[@id=2]"))));
        Assert.Equal(byOrdinal["path"]!.GetValue<string>(), byId["path"]!.GetValue<string>());
    }

    [Fact]
    public void Get_MasterOutOfRange_IsInvalidPathWithCandidates()
    {
        Create();
        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[9]")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.Equal(new[] { "/master[1]" }, error.Candidates!);
    }

    [Fact]
    public void Get_LayoutOutOfRange_IsInvalidPathWithCandidates()
    {
        Create();
        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/layout[9]")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
        Assert.Equal(new[] { "/master[1]/layout[1]" }, error.Candidates!);
    }

    [Fact]
    public void Get_MasterShapeOutOfRange_IsInvalidPath()
    {
        Create();
        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/shape[5]")));

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
    }

    [Fact]
    public void Get_ParagraphUnderMaster_IsUnsupportedPlannedM2()
    {
        Create();
        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/shape[1]/p[1]")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("M2", error.Message, StringComparison.Ordinal);
        Assert.Contains("/slide[", error.Suggestion, StringComparison.Ordinal);
    }

    // ---- read --view structure ----------------------------------------------

    [Fact]
    public void Structure_ListsMastersLayoutsAndPerSlideLayoutUse()
    {
        Create();
        AddTitleOnlyLayout();
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("layout", JsonValue.Create(2)))),
        ]));

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));

        var master = data["masters"]!.AsArray().Single()!;
        Assert.Equal("/master[1]", master["path"]!.GetValue<string>());
        Assert.Equal(2, master["layoutCount"]!.GetValue<int>());

        var layouts = master["layouts"]!.AsArray();
        Assert.Equal("/master[1]/layout[1]", layouts[0]!["path"]!.GetValue<string>());
        Assert.Equal("blank", layouts[0]!["type"]!.GetValue<string>());
        Assert.Equal(new[] { 1 }, layouts[0]!["usedBySlides"]!.AsArray().Select(n => n!.GetValue<int>()).ToArray());
        Assert.Equal("titleOnly", layouts[1]!["type"]!.GetValue<string>());
        Assert.Equal(new[] { 2 }, layouts[1]!["usedBySlides"]!.AsArray().Select(n => n!.GetValue<int>()).ToArray());

        var slides = data["slides"]!.AsArray();
        Assert.Equal("/master[1]/layout[1]", slides[0]!["layout"]!.GetValue<string>());
        Assert.Equal("/master[1]/layout[2]", slides[1]!["layout"]!.GetValue<string>());
    }

    // ---- query ----------------------------------------------------------------

    [Fact]
    public void Query_MasterElement_ListsMasters()
    {
        Create();
        var data = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "master"))));

        Assert.Equal(1, data["count"]!.GetValue<int>());
        var match = data["matches"]![0]!;
        Assert.Equal("/master[1]", match["path"]!.GetValue<string>());
        Assert.Equal("master", match["kind"]!.GetValue<string>());
        Assert.Equal(1, match["layoutCount"]!.GetValue<int>());
    }

    [Fact]
    public void Query_LayoutElement_FiltersByType()
    {
        Create();
        AddTitleOnlyLayout();
        var data = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "layout[type=titleOnly]"))));

        Assert.Equal(1, data["count"]!.GetValue<int>());
        var match = data["matches"]![0]!;
        Assert.Equal("/master[1]/layout[2]", match["path"]!.GetValue<string>());
        Assert.Equal("Title Only", match["name"]!.GetValue<string>());
    }

    [Fact]
    public void Query_ShapeScopedIntoLayout_ReturnsMasterPaths()
    {
        Create();
        AddTitleOnlyLayout();
        var data = TestEnv.AssertOk(_handler.Query(
            _ws.Ctx("deck.pptx", ("selector", "shape"), ("scope", "/master[1]/layout[2]"))));

        Assert.Equal("/master[1]/layout[2]", data["scope"]!.GetValue<string>());
        Assert.Equal(1, data["count"]!.GetValue<int>());
        var match = data["matches"]![0]!;
        Assert.Matches(new Regex(@"^/master\[1\]/layout\[2\]/shape\[@id=[0-9]+\]$"), match["path"]!.GetValue<string>());
        Assert.Equal("title", match["placeholder"]!.GetValue<string>());
        Assert.Equal("Layout title", match["text"]!.GetValue<string>());
    }

    [Fact]
    public void Query_SlideElementInMasterScope_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Query(_ws.Ctx("deck.pptx", ("selector", "slide"), ("scope", "/master[1]")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("shape", error.Candidates!);
    }

    // ---- edits stay closed until M2 -------------------------------------------

    [Fact]
    public void Edit_AnyOpOnMasterOrLayoutPaths_IsUnsupportedPlannedM2AndNoWrite()
    {
        var file = Create();
        AddTitleOnlyLayout();
        var before = File.ReadAllBytes(file);

        EditOp[] ops =
        [
            TestEnv.Op("set", "/master[1]/shape[1]", props: TestEnv.Props(("text", JsonValue.Create("nope")))),
            TestEnv.Op("set", "/master[1]/layout[2]/shape[1]", props: TestEnv.Props(("fill", JsonValue.Create("FF0000")))),
            TestEnv.Op("add", "/master[1]", type: "shape"),
            TestEnv.Op("remove", "/master[1]/layout[2]"),
            TestEnv.Op("move", "/master[1]", position: "1"),
        ];

        foreach (var op in ops)
        {
            var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [op]);
            var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
            Assert.Equal(ExitCodes.UnsupportedFeature, envelope.ExitCode);
            Assert.Contains("M2", error.Message, StringComparison.Ordinal);
            Assert.Contains("add slide", error.Suggestion, StringComparison.Ordinal);
        }

        Assert.Equal(before, File.ReadAllBytes(file));
    }

    [Fact]
    public void Move_WithMasterAnchor_IsInvalidArgs()
    {
        Create();
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[2]", type: "slide")]));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("move", "/slide[2]", position: "before:/master[1]"),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Render_MasterScope_IsUnsupportedPlannedM2()
    {
        Create();
        var envelope = _handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/master[1]")));

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("/slide[", error.Suggestion, StringComparison.Ordinal);
    }

    // ---- round-trip law over the new read surface ------------------------------

    [Fact]
    public void MasterReadVerbs_LeaveEveryByteUntouched()
    {
        var file = Create();
        AddTitleOnlyLayout();
        var before = File.ReadAllBytes(file);

        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/layout[2]"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/master[1]/layout[2]/shape[1]"))));
        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "layout"))));
        TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "shape"), ("scope", "/master[1]"))));

        Assert.Equal(before, File.ReadAllBytes(file));
    }

    // ---- add slide with a layout pick ------------------------------------------

    [Fact]
    public void AddSlide_WithLayoutIndex_BindsThatLayout_ReopenVerified()
    {
        var file = Create();
        AddTitleOnlyLayout();

        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(
                ("layout", JsonValue.Create(2)),
                ("title", JsonValue.Create("Uses Title Only")))),
        ]));
        Assert.Equal(2, data["slides"]!.GetValue<int>());

        using (var doc = PresentationDocument.Open(file, false))
        {
            var presentation = doc.PresentationPart!;
            var slideIds = presentation.Presentation!.SlideIdList!.Elements<P.SlideId>().ToList();
            var newSlide = (SlidePart)presentation.GetPartById(slideIds[1].RelationshipId!.Value!);

            var layout = newSlide.SlideLayoutPart!;
            Assert.Equal(P.SlideLayoutValues.TitleOnly, layout.SlideLayout!.Type!.Value);
            Assert.Equal("Title Only", layout.SlideLayout.CommonSlideData!.Name!.Value);
        }

        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        Assert.Equal("/master[1]/layout[2]", structure["slides"]!.AsArray()[1]!["layout"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddSlide_WithStringLayoutIndex_BindsThatLayout()
    {
        // The CLI sugar and the MCP schema deliver props string-valued
        // ({"layout":"2"}); the index must parse exactly like the number form.
        Create();
        AddTitleOnlyLayout();

        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("layout", JsonValue.Create("2")))),
        ]));

        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        Assert.Equal("/master[1]/layout[2]", structure["slides"]!.AsArray()[1]!["layout"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddSlide_DefaultLayout_StaysTheMastersFirst()
    {
        Create();
        AddTitleOnlyLayout();
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[2]", type: "slide")]));

        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        Assert.Equal("/master[1]/layout[1]", structure["slides"]!.AsArray()[1]!["layout"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddSlide_LayoutOutOfRange_IsInvalidArgsWithCandidatesAndNoWrite()
    {
        var file = Create();
        var before = File.ReadAllBytes(file);

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("layout", JsonValue.Create(7)))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("/master[1]/layout[1]", error.Candidates!);
        Assert.Contains("structure", error.Suggestion, StringComparison.Ordinal);
        Assert.Equal(before, File.ReadAllBytes(file));
    }

    [Fact]
    public void AddSlide_LayoutNotAnInteger_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide", props: TestEnv.Props(("layout", JsonValue.Create("Blank")))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }
}
