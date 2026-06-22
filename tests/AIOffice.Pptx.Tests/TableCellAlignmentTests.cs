using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.7.0 pptx table-cell polish: vertical alignment (a:tcPr/@anchor), cell margins
/// (marL/marR/marT/marB) and text direction (a:tcPr/@vert). Each reopens
/// validator-clean and is surfaced by get.
/// </summary>
public sealed class TableCellAlignmentTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private JsonObject Edit2(string file, params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx(file), ops));

    private string CreateWithTable()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "table",
            props: TestEnv.Props(("rows", 3), ("cols", 3), ("x", "2cm"), ("y", "5cm"), ("w", "27cm"))));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private A.TableCell Cell(PresentationDocument doc, int row, int col) =>
        doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().Single().Graphic!.GraphicData!.GetFirstChild<A.Table>()!
            .Elements<A.TableRow>().ElementAt(row - 1).Elements<A.TableCell>().ElementAt(col - 1);

    [Fact]
    public void SetValign_WritesAnchor_ReopenVerified()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[1]", props: TestEnv.Props(("valign", "middle"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Equal(A.TextAnchoringTypeValues.Center, Cell(doc, 2, 1).TableCellProperties!.Anchor!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]/tr[2]/tc[1]"))));
        Assert.Equal("middle", detail["vAlign"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("top")]
    [InlineData("bottom")]
    public void SetValign_AllAnchors(string token)
    {
        var expected = token == "top" ? A.TextAnchoringTypeValues.Top : A.TextAnchoringTypeValues.Bottom;
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("valign", token))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Equal(expected, Cell(doc, 1, 1).TableCellProperties!.Anchor!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]/tr[1]/tc[1]"))));
        Assert.Equal(token, detail["vAlign"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetMargins_WriteEmuAttributes_ReopenVerified()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[2]", props: TestEnv.Props(
            ("marginLeft", "0.2cm"), ("marginRight", "0.2cm"), ("marginTop", "0.1cm"), ("marginBottom", "0.1cm"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tcPr = Cell(doc, 1, 2).TableCellProperties!;
            Assert.Equal(72_000, tcPr.LeftMargin!.Value);   // 0.2cm = 72000 EMU
            Assert.Equal(72_000, tcPr.RightMargin!.Value);
            Assert.Equal(36_000, tcPr.TopMargin!.Value);    // 0.1cm = 36000 EMU
            Assert.Equal(36_000, tcPr.BottomMargin!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]/tr[1]/tc[2]"))));
        Assert.Equal(0.2, detail["marginLeft"]!.GetValue<double>());
        Assert.Equal(0.1, detail["marginTop"]!.GetValue<double>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetTextDirection_WritesVert_ReopenVerified()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("textDirection", "vertical"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Equal(A.TextVerticalValues.Vertical, Cell(doc, 1, 1).TableCellProperties!.Vertical!.Value);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]/tr[1]/tc[1]"))));
        Assert.Equal("vertical", detail["textDirection"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ValignAndFillCoexistOnTheSameCell()
    {
        CreateWithTable();
        // Fill appends a child to a:tcPr; valign sets an attribute — both must validate together.
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[2]", props: TestEnv.Props(
            ("fill", "FFEEAA"), ("valign", "bottom"), ("marginLeft", "0.3cm"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tcPr = Cell(doc, 2, 2).TableCellProperties!;
            Assert.Equal(A.TextAnchoringTypeValues.Bottom, tcPr.Anchor!.Value);
            Assert.Equal("FFEEAA", tcPr.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Equal(108_000, tcPr.LeftMargin!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void BadValign_Fails()
    {
        CreateWithTable();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("valign", "centre-ish")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void BadTextDirection_Fails()
    {
        CreateWithTable();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("textDirection", "diagonal")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void NegativeMargin_Fails()
    {
        CreateWithTable();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("marginLeft", "-1cm")))]),
            ErrorCodes.InvalidArgs);
    }

    // ----- v1.17.0 per-edge cell borders -------------------------------------------

    private JsonObject Border(params (string Key, JsonNode? Value)[] pairs) => TestEnv.Props(pairs);

    [Fact]
    public void SetCell_PartialEdges_CoexistWithFill()
    {
        CreateWithTable();
        // Fill first (it appends LAST in a:tcPr), THEN borders: the a:ln* edges must be
        // re-homed to the FRONT so they precede the fill group and the part stays repair-free.
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[2]", props: TestEnv.Props(("fill", "FFEEAA"))));
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[2]", props: TestEnv.Props(("borders",
            new JsonObject { ["top"] = Border(("color", "C00000")), ["right"] = Border(("widthPt", 2)) }))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tcPr = Cell(doc, 2, 2).TableCellProperties!;
            var lnT = tcPr.GetFirstChild<A.TopBorderLineProperties>();
            var lnR = tcPr.GetFirstChild<A.RightBorderLineProperties>();
            var fill = tcPr.GetFirstChild<A.SolidFill>();
            Assert.NotNull(lnT);
            Assert.NotNull(lnR);
            Assert.NotNull(fill);

            // Both line edges must precede the fill in document order.
            var children = tcPr.ChildElements.ToList();
            Assert.True(children.IndexOf(lnT!) < children.IndexOf(fill!));
            Assert.True(children.IndexOf(lnR!) < children.IndexOf(fill!));
            // Schema order among the lines: lnR (right) follows lnT? No — lnL,lnR,lnT,lnB:
            // here only top+right are present, lnR must precede lnT.
            Assert.True(children.IndexOf(lnR!) < children.IndexOf(lnT!));
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetCell_Borders_WriteLineEdgesWithColorWidthAndStyles()
    {
        CreateWithTable();
        Edit(TestEnv.Op("set", "/slide[1]/table[1]/tr[2]/tc[1]", props: TestEnv.Props(("borders",
            new JsonObject
            {
                ["top"] = Border(("color", "C00000"), ("widthPt", 1.5), ("style", "single")),
                ["left"] = Border(("widthPt", 1), ("style", "dashed")),
            }))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tcPr = Cell(doc, 2, 1).TableCellProperties!;
            var lnT = tcPr.GetFirstChild<A.TopBorderLineProperties>()!;
            Assert.Equal(19_050, lnT.Width!.Value); // 1.5pt * 12700
            Assert.Equal("C00000", lnT.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Null(lnT.GetFirstChild<A.PresetDash>());

            var lnL = tcPr.GetFirstChild<A.LeftBorderLineProperties>()!;
            Assert.Equal(12_700, lnL.Width!.Value);
            Assert.Equal(A.PresetLineDashValues.Dash, lnL.GetFirstChild<A.PresetDash>()!.Val!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetCell_Borders_All_DoubleAndNone_AndReplaceOnRepeat()
    {
        CreateWithTable();
        var path = "/slide[1]/table[1]/tr[2]/tc[3]";

        // 'all' paints four double edges.
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("borders",
            new JsonObject { ["all"] = Border(("color", "112233"), ("style", "double")) }))));
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tcPr = Cell(doc, 2, 3).TableCellProperties!;
            foreach (var line in new A.LinePropertiesType?[]
            {
                tcPr.GetFirstChild<A.LeftBorderLineProperties>(),
                tcPr.GetFirstChild<A.RightBorderLineProperties>(),
                tcPr.GetFirstChild<A.TopBorderLineProperties>(),
                tcPr.GetFirstChild<A.BottomBorderLineProperties>(),
            })
            {
                Assert.NotNull(line);
                Assert.Equal(A.CompoundLineValues.Double, line!.CompoundLineType!.Value);
                Assert.Equal("112233", line.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            }
        }

        // Repeat: clear top with bare "none", replace left with a single solid edge (no stacking).
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("borders",
            new JsonObject { ["top"] = JsonValue.Create("none"), ["left"] = Border(("style", "single")) }))));
        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tcPr = Cell(doc, 2, 3).TableCellProperties!;
            Assert.Null(tcPr.GetFirstChild<A.TopBorderLineProperties>()); // cleared
            Assert.Single(tcPr.Elements<A.LeftBorderLineProperties>());    // replaced, not stacked
            Assert.Null(tcPr.GetFirstChild<A.LeftBorderLineProperties>()!.CompoundLineType); // now single
            // bottom + right untouched from the 'all' pass.
            Assert.NotNull(tcPr.GetFirstChild<A.BottomBorderLineProperties>());
            Assert.NotNull(tcPr.GetFirstChild<A.RightBorderLineProperties>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void GetCell_Borders_RoundTrip_IsByteStable()
    {
        CreateWithTable();
        var path = "/slide[1]/table[1]/tr[3]/tc[2]";
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("borders",
            new JsonObject
            {
                ["top"] = Border(("color", "C00000"), ("widthPt", 1.5), ("style", "single")),
                ["bottom"] = Border(("style", "dotted")),
            }))));

        // A neutral (preset-styled) cell with no a:ln* edges reports no borders.
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("neutral.pptx")));
        Edit2("neutral.pptx", TestEnv.Op("add", "/slide[1]", type: "table",
            props: TestEnv.Props(("rows", 2), ("cols", 2), ("style", "none"))));
        var plain = TestEnv.AssertOk(_handler.Get(_ws.Ctx("neutral.pptx", ("path", "/slide[1]/table[1]/tr[1]/tc[1]"))));
        Assert.Null(plain["borders"]);

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("single", detail["borders"]!["top"]!["style"]!.GetValue<string>());
        Assert.Equal("C00000", detail["borders"]!["top"]!["color"]!.GetValue<string>());
        Assert.Equal(1.5, detail["borders"]!["top"]!["widthPt"]!.GetValue<double>());
        Assert.Equal("dotted", detail["borders"]!["bottom"]!["style"]!.GetValue<string>());
        // The default light look painted lnL/lnR at 1pt; the untouched left edge round-trips.
        Assert.Equal("single", detail["borders"]!["left"]!["style"]!.GetValue<string>());
        Assert.Equal(1.0, detail["borders"]!["left"]!["widthPt"]!.GetValue<double>());

        // Capture the a:tcPr after the first write.
        string Tc()
        {
            using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
            return Cell(doc, 3, 2).TableCellProperties!.OuterXml;
        }

        var first = Tc();

        // Re-applying the same edges (set->get->set) leaves the a:tcPr byte-identical.
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("borders",
            new JsonObject
            {
                ["top"] = Border(("color", "C00000"), ("widthPt", 1.5), ("style", "single")),
                ["bottom"] = Border(("style", "dotted")),
            }))));
        Assert.Equal(first, Tc());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void BadBorderEdge_Fails()
    {
        CreateWithTable();
        var fail = TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("borders",
                    new JsonObject { ["topleft"] = Border(("color", "000000")) })))]),
            ErrorCodes.InvalidArgs);
        Assert.Contains("top", fail.Candidates!);
        Assert.Contains("all", fail.Candidates!);
    }

    [Fact]
    public void BadBorderStyle_Fails()
    {
        CreateWithTable();
        var fail = TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("borders",
                    new JsonObject { ["top"] = Border(("style", "groovy")) })))]),
            ErrorCodes.InvalidArgs);
        Assert.Contains("single", fail.Candidates!);
        Assert.Contains("none", fail.Candidates!);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData("wide")]
    public void BadBorderWidth_Fails(object width)
    {
        CreateWithTable();
        var widthNode = width is int i ? JsonValue.Create(i) : JsonValue.Create((string)width);
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("set", "/slide[1]/table[1]/tr[1]/tc[1]", props: TestEnv.Props(("borders",
                    new JsonObject { ["top"] = new JsonObject { ["widthPt"] = widthNode } })))]),
            ErrorCodes.InvalidArgs);
    }
}
