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
}
