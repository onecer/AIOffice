using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// v1.4.0 built-in table style presets (additive on the table op): style none/light1/.../dark2
/// maps to an a:tableStyleId GUID, plus firstRow/lastRow/bandRow/firstCol a:tblPr flags. set on an
/// existing table can change the style; get reports it. The legacy light/medium/dark direct-paint
/// looks still work unchanged.
/// </summary>
public sealed class TableStyleTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private void CreatePreset(string style, params (string Key, JsonNode? Value)[] extra)
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        var props = TestEnv.Props(("rows", 3), ("cols", 3), ("style", style));
        foreach (var (k, v) in extra)
        {
            props[k] = v;
        }

        Edit(TestEnv.Op("add", "/slide[1]", type: "table", props: props));
    }

    private A.Table OpenTable(PresentationDocument doc) =>
        doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!
            .Elements<P.GraphicFrame>().Single().Graphic!.GraphicData!.GetFirstChild<A.Table>()!;

    // ----- add with preset --------------------------------------------------------

    [Fact]
    public void AddTable_WithPreset_WritesTableStyleId_NoDirectFills_AndValidates()
    {
        CreatePreset("medium2", ("bandRow", JsonValue.Create(true)));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var table = OpenTable(doc);
            var tblPr = table.TableProperties!;
            Assert.Equal("{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}", tblPr.GetFirstChild<A.TableStyleId>()!.Text);
            Assert.True(tblPr.BandRow!.Value);

            // Preset cells are neutral: no direct solid fill, no explicit borders.
            var cell = table.Elements<A.TableRow>().First().Elements<A.TableCell>().First();
            Assert.Null(cell.TableCellProperties!.GetFirstChild<A.SolidFill>());
            Assert.Null(cell.TableCellProperties!.GetFirstChild<A.LeftBorderLineProperties>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("light1", "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}")]
    [InlineData("light2", "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}")]
    [InlineData("medium1", "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}")]
    [InlineData("medium3", "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}")]
    [InlineData("dark1", "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}")]
    [InlineData("dark2", "{5940675A-B579-460E-94D1-54222C63F5DA}")]
    public void AddTable_EachPreset_MapsToItsBuiltInGuid_AndValidates(string preset, string guid)
    {
        CreatePreset(preset);

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Equal(guid, OpenTable(doc).TableProperties!.GetFirstChild<A.TableStyleId>()!.Text);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddTable_NonePreset_WritesNoStyleId_ButKeepsFlags_AndValidates()
    {
        CreatePreset("none", ("firstRow", JsonValue.Create(true)), ("firstCol", JsonValue.Create(true)));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tblPr = OpenTable(doc).TableProperties!;
            Assert.Null(tblPr.GetFirstChild<A.TableStyleId>());
            Assert.True(tblPr.FirstRow!.Value);
            Assert.True(tblPr.FirstColumn!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddTable_AllStyleOptions_WriteTheTblPrFlags()
    {
        CreatePreset("medium2",
            ("firstRow", JsonValue.Create(true)),
            ("lastRow", JsonValue.Create(true)),
            ("bandRow", JsonValue.Create(true)),
            ("firstCol", JsonValue.Create(true)));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tblPr = OpenTable(doc).TableProperties!;
            Assert.True(tblPr.FirstRow!.Value);
            Assert.True(tblPr.LastRow!.Value);
            Assert.True(tblPr.BandRow!.Value);
            Assert.True(tblPr.FirstColumn!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void StyleOptions_OnDirectPaintLook_IsInvalidArgs()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("add", "/slide[1]", type: "table",
                props: TestEnv.Props(("rows", 2), ("cols", 2), ("style", "medium"), ("bandRow", JsonValue.Create(true))))]),
            ErrorCodes.InvalidArgs);
    }

    // ----- get --------------------------------------------------------------------

    [Fact]
    public void Get_PresetTable_ReportsStyleAndOptions()
    {
        CreatePreset("dark2", ("bandRow", JsonValue.Create(true)), ("firstRow", JsonValue.Create(true)));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]"))));
        Assert.Equal("dark2", detail["style"]!.GetValue<string>());
        Assert.True(detail["styleOptions"]!["bandRow"]!.GetValue<bool>());
        Assert.True(detail["styleOptions"]!["firstRow"]!.GetValue<bool>());
    }

    [Fact]
    public void Get_DirectPaintTable_HasNoPresetStyle()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        Edit(TestEnv.Op("add", "/slide[1]", type: "table",
            props: TestEnv.Props(("rows", 2), ("cols", 2), ("style", "medium"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]"))));
        Assert.Null(detail["style"]); // direct-paint look carries no a:tableStyleId
    }

    // ----- set --------------------------------------------------------------------

    [Fact]
    public void SetStyle_OnExistingTable_ChangesTheStyleId_AndValidates()
    {
        CreatePreset("light1");
        Edit(TestEnv.Op("set", "/slide[1]/table[1]", props: TestEnv.Props(("style", "dark1"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Equal("{E8034E78-7F5D-4C2E-B375-FC64B27BC917}",
                OpenTable(doc).TableProperties!.GetFirstChild<A.TableStyleId>()!.Text);
        }

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/table[1]"))));
        Assert.Equal("dark1", detail["style"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetStyleNone_OnExistingPresetTable_ClearsTheStyleId()
    {
        CreatePreset("medium2");
        Edit(TestEnv.Op("set", "/slide[1]/table[1]", props: TestEnv.Props(("style", "none"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.Null(OpenTable(doc).TableProperties!.GetFirstChild<A.TableStyleId>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetStyleAndFlags_InOneOp_BothApply()
    {
        CreatePreset("light1");
        Edit(TestEnv.Op("set", "/slide[1]/table[1]", props: TestEnv.Props(
            ("style", "medium2"), ("bandRow", JsonValue.Create(true)), ("lastRow", JsonValue.Create(true)))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var tblPr = OpenTable(doc).TableProperties!;
            Assert.Equal("{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}", tblPr.GetFirstChild<A.TableStyleId>()!.Text);
            Assert.True(tblPr.BandRow!.Value);
            Assert.True(tblPr.LastRow!.Value);
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetFlags_OnDirectPaintTable_IsInvalidArgs()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        Edit(TestEnv.Op("add", "/slide[1]", type: "table",
            props: TestEnv.Props(("rows", 2), ("cols", 2), ("style", "medium"))));

        // Flags alone need a built-in preset (the table has no a:tableStyleId).
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", "/slide[1]/table[1]",
                props: TestEnv.Props(("bandRow", JsonValue.Create(true))))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void SetStyle_ToADirectPaintLook_IsTypedUnsupported()
    {
        CreatePreset("medium2");
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("set", "/slide[1]/table[1]",
                props: TestEnv.Props(("style", "light")))]),
            ErrorCodes.UnsupportedFeature);
    }

    // ----- legacy direct-paint looks still work -----------------------------------

    [Fact]
    public void DirectPaintLooks_StillPaintCells_WithNoStyleId()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        Edit(TestEnv.Op("add", "/slide[1]", type: "table",
            props: TestEnv.Props(("rows", 3), ("cols", 2), ("headerRow", JsonValue.Create(true)), ("style", "dark"))));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var table = OpenTable(doc);
            Assert.Empty(table.TableProperties!.Elements<A.TableStyleId>());
            Assert.True(table.TableProperties!.FirstRow!.Value); // headerRow still recorded
            var header = table.Elements<A.TableRow>().First().Elements<A.TableCell>().First();
            Assert.NotNull(header.TableCellProperties!.GetFirstChild<A.SolidFill>()); // painted
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
