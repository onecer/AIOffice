using System.Text;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Dgm = DocumentFormat.OpenXml.Drawing.Diagrams;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// M5 SmartArt (read-only): the fixture deck is built in-test from raw diagram
/// part XML (a minimal but valid data+layout+colors+quickStyle quartet), so no
/// external files are needed.
/// </summary>
public sealed class SmartArtTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private const string DgmNs = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    private const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    private const string DataModelXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dgm:ptLst>
            <dgm:pt modelId="{10000000-0000-0000-0000-000000000001}" type="doc"/>
            <dgm:pt modelId="{10000000-0000-0000-0000-000000000002}">
              <dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Plan</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="{10000000-0000-0000-0000-000000000003}">
              <dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Build</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="{10000000-0000-0000-0000-000000000004}">
              <dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Ship</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="{10000000-0000-0000-0000-000000000005}">
              <dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Iterate fast</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
          </dgm:ptLst>
          <dgm:cxnLst>
            <dgm:cxn modelId="{20000000-0000-0000-0000-000000000001}" srcId="{10000000-0000-0000-0000-000000000001}" destId="{10000000-0000-0000-0000-000000000002}" srcOrd="0" destOrd="0"/>
            <dgm:cxn modelId="{20000000-0000-0000-0000-000000000002}" srcId="{10000000-0000-0000-0000-000000000001}" destId="{10000000-0000-0000-0000-000000000003}" srcOrd="1" destOrd="0"/>
            <dgm:cxn modelId="{20000000-0000-0000-0000-000000000003}" srcId="{10000000-0000-0000-0000-000000000001}" destId="{10000000-0000-0000-0000-000000000004}" srcOrd="2" destOrd="0"/>
            <dgm:cxn modelId="{20000000-0000-0000-0000-000000000004}" srcId="{10000000-0000-0000-0000-000000000003}" destId="{10000000-0000-0000-0000-000000000005}" srcOrd="0" destOrd="0"/>
          </dgm:cxnLst>
          <dgm:bg/>
          <dgm:whole/>
        </dgm:dataModel>
        """;

    private const string LayoutXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:aioffice/test/process1">
          <dgm:title val="Basic Process"/>
          <dgm:desc val=""/>
          <dgm:catLst><dgm:cat type="process" pri="1000"/></dgm:catLst>
          <dgm:layoutNode name="root"/>
        </dgm:layoutDef>
        """;

    private const string ColorsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:aioffice/test/colors1"/>
        """;

    private const string StyleXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:aioffice/test/style1">
          <dgm:styleLbl name="node0"/>
        </dgm:styleDef>
        """;

    /// <summary>Creates deck.pptx and injects a minimal SmartArt frame on slide 1.</summary>
    private void CreateWithSmartArt()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "SmartArt fixture"))));

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), true);
        var slidePart = doc.PresentationPart!.SlideParts.Single();

        var dataPart = slidePart.AddNewPart<DiagramDataPart>();
        FeedXml(dataPart, DataModelXml);
        var layoutPart = slidePart.AddNewPart<DiagramLayoutDefinitionPart>();
        FeedXml(layoutPart, LayoutXml);
        var colorsPart = slidePart.AddNewPart<DiagramColorsPart>();
        FeedXml(colorsPart, ColorsXml);
        var stylePart = slidePart.AddNewPart<DiagramStylePart>();
        FeedXml(stylePart, StyleXml);

        var tree = slidePart.Slide!.CommonSlideData!.ShapeTree!;
        tree.Append(new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = 90U, Name = "Diagram 1" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = 914_400L, Y = 1_828_800L },
                new A.Extents { Cx = 6_096_000L, Cy = 3_429_000L }),
            new A.Graphic(new A.GraphicData(new Dgm.RelationshipIds
            {
                DataPart = slidePart.GetIdOfPart(dataPart),
                LayoutPart = slidePart.GetIdOfPart(layoutPart),
                ColorPart = slidePart.GetIdOfPart(colorsPart),
                StylePart = slidePart.GetIdOfPart(stylePart),
            })
            {
                Uri = DgmNs,
            })));
    }

    private static void FeedXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(xml);
    }

    [Fact]
    public void Fixture_IsValidatorClean()
    {
        CreateWithSmartArt();
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_SmartArtPath_ReportsLayoutNodeCountAndTexts()
    {
        CreateWithSmartArt();
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/smartart[1]"))));

        Assert.Equal("/slide[1]/smartart[1]", detail["path"]!.GetValue<string>());
        Assert.Equal("smartart", detail["kind"]!.GetValue<string>());
        Assert.True(detail["readOnly"]!.GetValue<bool>());
        Assert.Equal("Basic Process", detail["layout"]!.GetValue<string>());
        Assert.Equal(4, detail["nodeCount"]!.GetValue<int>());
        Assert.Equal("/slide[1]/shape[@id=90]", detail["shapePath"]!.GetValue<string>());

        var texts = detail["texts"]!.AsArray();
        Assert.Equal(3, texts.Count);
        Assert.Equal("Plan", texts[0]!["text"]!.GetValue<string>());
        Assert.Equal(0, texts[0]!["level"]!.GetValue<int>());
        Assert.Equal("Build", texts[1]!["text"]!.GetValue<string>());
        Assert.Equal("Iterate fast", texts[1]!["children"]![0]!["text"]!.GetValue<string>());
        Assert.Equal(1, texts[1]!["children"]![0]!["level"]!.GetValue<int>());
        Assert.Equal("Ship", texts[2]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void Get_SmartArtOutOfRange_IsInvalidPath()
    {
        CreateWithSmartArt();
        TestEnv.AssertFail(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/smartart[2]"))),
            ErrorCodes.InvalidPath);
    }

    [Fact]
    public void ReadTextView_FlattensNodeTextsWithIndentation()
    {
        CreateWithSmartArt();
        var text = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))))["text"]!.GetValue<string>();

        Assert.Contains("Plan\nBuild\n  Iterate fast\nShip", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Structure_ListsSmartArtPerSlide()
    {
        CreateWithSmartArt();
        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));

        var smartArt = data["slides"]![0]!["smartArt"]!.AsArray();
        var row = Assert.Single(smartArt)!;
        Assert.Equal("/slide[1]/smartart[1]", row["path"]!.GetValue<string>());
        Assert.Equal("Basic Process", row["layout"]!.GetValue<string>());
        Assert.Equal(4, row["nodeCount"]!.GetValue<int>());
        Assert.Equal(["Plan", "Build", "  Iterate fast", "Ship"],
            row["texts"]!.AsArray().Select(t => t!.GetValue<string>()).ToList());
    }

    [Fact]
    public void Query_ContainsMatchesSmartArtText()
    {
        CreateWithSmartArt();

        var smartart = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "smartart:contains('Iterate')"))));
        Assert.Equal(1, smartart["count"]!.GetValue<int>());
        var match = smartart["matches"]![0]!;
        Assert.Equal("/slide[1]/smartart[1]", match["path"]!.GetValue<string>());
        Assert.Equal("Basic Process", match["layout"]!.GetValue<string>());

        var slides = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "slide:contains('Iterate')"))));
        Assert.Equal(1, slides["count"]!.GetValue<int>());

        var none = TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "smartart:contains('Quarterly')"))));
        Assert.Equal(0, none["count"]!.GetValue<int>());
    }

    [Fact]
    public void Get_ShapePathOfDiagramFrame_LinksSmartArtPath()
    {
        CreateWithSmartArt();
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/shape[@id=90]"))));
        Assert.Equal("/slide[1]/smartart[1]", detail["smartArtPath"]!.GetValue<string>());

        var slide = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]"))));
        Assert.Equal("/slide[1]/smartart[1]", slide["smartArt"]![0]!.GetValue<string>());
    }

    [Theory]
    [InlineData("set")]
    [InlineData("remove")]
    [InlineData("move")]
    [InlineData("replace")]
    public void AnyEditOp_OnSmartArtPath_IsTypedUnsupported(string op)
    {
        CreateWithSmartArt();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            op, "/slide[1]/smartart[1]",
            position: op == "move" ? "front" : null,
            props: TestEnv.Props(("text", "nope"), ("find", "Plan")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
        Assert.Contains("PowerPoint", error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void AddOnSmartArtPath_IsTypedUnsupported()
    {
        CreateWithSmartArt();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "add", "/slide[1]/smartart[1]", type: "shape", props: TestEnv.Props(("text", "x")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.UnsupportedFeature);
    }

    [Fact]
    public void Svg_DrawsALabeledPlaceholder_NeverAFakeRedraw()
    {
        CreateWithSmartArt();
        var data = TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "svg"), ("scope", "/slide[1]"))));
        var svg = data["slides"]![0]!["svg"]!.GetValue<string>();

        Assert.Contains("<g data-aio-path=\"/slide[1]/shape[@id=90]\"", svg, StringComparison.Ordinal);
        Assert.Contains("[smartart] Basic Process", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-dasharray", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadSideVerbs_LeaveTheFixtureUntouched()
    {
        CreateWithSmartArt();
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "text"))));
        TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/smartart[1]"))));
        TestEnv.AssertOk(_handler.Query(_ws.Ctx("deck.pptx", ("selector", "smartart:contains('Plan')"))));
        TestEnv.AssertOk(_handler.Render(_ws.Ctx("deck.pptx", ("to", "html"))));
        TestEnv.AssertOk(_handler.Validate(_ws.Ctx("deck.pptx")));

        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }
}
