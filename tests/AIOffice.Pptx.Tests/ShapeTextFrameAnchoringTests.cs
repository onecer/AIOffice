using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// Covers text-frame anchoring on a text SHAPE's a:bodyPr, bringing it to parity with
/// table cells: vAlign -> @anchor, textDirection -> @vert, marginLeft/Right/Top/Bottom
/// -> lIns/rIns/tIns/bIns. These are bodyPr attributes (no element-ordering hazard) and
/// read back through 'get'.
/// </summary>
public sealed class ShapeTextFrameAnchoringTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private string Create()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        return _ws.PathOf("deck.pptx");
    }

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    private JsonObject Get(string path) =>
        TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));

    /// <summary>The shape's a:bodyPr, opened read-only off disk.</summary>
    private A.BodyProperties BodyPr(string file, string name)
    {
        using var doc = PresentationDocument.Open(file, false);
        var shape = doc.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<P.Shape>()
            .Single(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == name);
        return shape.TextBody!.GetFirstChild<A.BodyProperties>()!;
    }

    private string AddShape(string name = "Box")
    {
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("text", "Some text"),
            ("name", JsonValue.Create(name)))));
        return added["results"]![0]!["target"]!.GetValue<string>();
    }

    [Theory]
    [InlineData("top")]
    [InlineData("middle")]
    [InlineData("bottom")]
    public void SetVAlign_WritesBodyPrAnchor_AndGetProjectsItBack(string token)
    {
        var file = Create();
        var path = AddShape();
        var expected = token switch
        {
            "top" => A.TextAnchoringTypeValues.Top,
            "middle" => A.TextAnchoringTypeValues.Center,
            _ => A.TextAnchoringTypeValues.Bottom,
        };

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("vAlign", JsonValue.Create(token)))));

        // The saved bodyPr carries the matching @anchor value (t/ctr/b).
        Assert.Equal(expected, BodyPr(file, "Box").Anchor!.Value);

        // get projects vAlign back.
        Assert.Equal(token, Get(path)["vAlign"]!.GetValue<string>());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Theory]
    [InlineData("horizontal")]
    [InlineData("vertical")]
    [InlineData("vertical270")]
    public void SetTextDirection_WritesBodyPrVert_AndGetProjectsItBack(string token)
    {
        var file = Create();
        var path = AddShape();
        var expected = token switch
        {
            "horizontal" => A.TextVerticalValues.Horizontal,
            "vertical" => A.TextVerticalValues.Vertical,
            _ => A.TextVerticalValues.Vertical270,
        };

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("textDirection", JsonValue.Create(token)))));

        Assert.Equal(expected, BodyPr(file, "Box").Vertical!.Value);

        Assert.Equal(token, Get(path)["textDirection"]!.GetValue<string>());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetMargins_WritesBodyPrInsetEmu_GetReturnsThem_AndReSetIsByteIdentical()
    {
        var file = Create();
        var path = AddShape();

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("marginLeft", JsonValue.Create("0.2cm")),
            ("marginRight", JsonValue.Create("0.2cm")),
            ("marginTop", JsonValue.Create("0.2cm")),
            ("marginBottom", JsonValue.Create("0.2cm")))));

        // 0.2cm -> 72000 EMU on each inset attribute.
        var bodyPr = BodyPr(file, "Box");
        Assert.Equal(72000, bodyPr.LeftInset!.Value);
        Assert.Equal(72000, bodyPr.RightInset!.Value);
        Assert.Equal(72000, bodyPr.TopInset!.Value);
        Assert.Equal(72000, bodyPr.BottomInset!.Value);
        var firstXml = bodyPr.OuterXml;

        // get returns the margins (EMU -> cm).
        var detail = Get(path);
        Assert.Equal(0.2, detail["marginLeft"]!.GetValue<double>());
        Assert.Equal(0.2, detail["marginRight"]!.GetValue<double>());
        Assert.Equal(0.2, detail["marginTop"]!.GetValue<double>());
        Assert.Equal(0.2, detail["marginBottom"]!.GetValue<double>());

        // Re-setting with the returned values yields byte-identical bodyPr XML (no dropped attrs).
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("marginLeft", detail["marginLeft"]!.DeepClone()),
            ("marginRight", detail["marginRight"]!.DeepClone()),
            ("marginTop", detail["marginTop"]!.DeepClone()),
            ("marginBottom", detail["marginBottom"]!.DeepClone()))));

        Assert.Equal(firstXml, BodyPr(file, "Box").OuterXml);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void CombineAllFourMargins_AndVAlign_TextDirection_AndAutofit_WithoutClobbering()
    {
        var file = Create();
        var path = AddShape();

        // Seed an autofit child on the bodyPr first.
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("autofit", JsonValue.Create("shrink")))));

        // Set every anchoring prop in one op.
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("vAlign", JsonValue.Create("middle")),
            ("textDirection", JsonValue.Create("vertical")),
            ("marginLeft", JsonValue.Create("0.1cm")),
            ("marginRight", JsonValue.Create("0.2cm")),
            ("marginTop", JsonValue.Create("0.3cm")),
            ("marginBottom", JsonValue.Create("0.4cm")))));

        var bodyPr = BodyPr(file, "Box");
        Assert.Equal(A.TextAnchoringTypeValues.Center, bodyPr.Anchor!.Value);
        Assert.Equal(A.TextVerticalValues.Vertical, bodyPr.Vertical!.Value);
        Assert.Equal(36000, bodyPr.LeftInset!.Value);
        Assert.Equal(72000, bodyPr.RightInset!.Value);
        Assert.Equal(108000, bodyPr.TopInset!.Value);
        Assert.Equal(144000, bodyPr.BottomInset!.Value);

        // The pre-existing autofit child survives the attribute writes.
        Assert.NotNull(bodyPr.GetFirstChild<A.NormalAutoFit>());

        var detail = Get(path);
        Assert.Equal("middle", detail["vAlign"]!.GetValue<string>());
        Assert.Equal("vertical", detail["textDirection"]!.GetValue<string>());
        Assert.Equal("shrink", detail["autofit"]!["mode"]!.GetValue<string>());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ShapeWithoutAnchoringProps_ProjectsNullFields_AndStaysByteStable()
    {
        var file = Create();
        var path = AddShape();
        var before = BodyPr(file, "Box").OuterXml;

        // A shape that never set these props projects null (not "" / 0).
        var detail = Get(path);
        Assert.True(detail["vAlign"] is null);
        Assert.True(detail["textDirection"] is null);
        Assert.True(detail["marginLeft"] is null);
        Assert.True(detail["marginRight"] is null);
        Assert.True(detail["marginTop"] is null);
        Assert.True(detail["marginBottom"] is null);

        // Setting an unrelated prop leaves the bodyPr byte-stable for the anchoring attrs.
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("name", JsonValue.Create("Box")))));
        Assert.Equal(before, BodyPr(file, "Box").OuterXml);
    }

    [Theory]
    [InlineData("up")]
    [InlineData("center-ish")]
    public void InvalidVAlign_IsRejectedWithInvalidArgs(string bad)
    {
        Create();
        var path = AddShape();

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("vAlign", JsonValue.Create(bad))))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Equal(new[] { "top", "middle", "bottom" }, error.Candidates);
    }

    [Theory]
    [InlineData("vert360")]
    [InlineData("sideways")]
    public void InvalidTextDirection_IsRejectedWithInvalidArgs(string bad)
    {
        Create();
        var path = AddShape();

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("textDirection", JsonValue.Create(bad))))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Equal(new[] { "horizontal", "vertical" }, error.Candidates);
    }

    [Fact]
    public void NegativeMargin_IsRejectedWithInvalidArgs()
    {
        Create();
        var path = AddShape();

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("marginLeft", JsonValue.Create("-0.2cm"))))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void OverRangeMargin_IsRejectedWithInvalidArgs()
    {
        // The second half of ParseMarginEmu's guard: an EMU value past int.MaxValue
        // (100000cm = 3.6e10 EMU >> 2.1e9) is rejected, not silently truncated.
        Create();
        var path = AddShape();

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("marginTop", JsonValue.Create("100000cm"))))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void GroupChildShape_RoundTripsAnchoringProps_ThroughGet()
    {
        // The read side does NOT auto-mirror table cells, so the group-child projection
        // (GroupDetail child) must be verified, not assumed-symmetric: set anchoring on a
        // shape INSIDE a group via the group-child path, then get it back.
        var file = Create();
        var childId = AddShape("Inner").Split("@id=")[1].TrimEnd(']');
        var siblingId = AddShape("Sibling").Split("@id=")[1].TrimEnd(']'); // a group needs >= 2 shapes
        var grouped = Edit(TestEnv.Op("add", "/slide[1]", type: "group", props: TestEnv.Props(
            ("shapes", new JsonArray("@" + childId, "@" + siblingId)))));
        var groupPath = grouped["results"]![0]!["target"]!.GetValue<string>();
        var inGroupPath = $"{groupPath}/shape[@id={childId}]";

        Edit(TestEnv.Op("set", inGroupPath, props: TestEnv.Props(
            ("vAlign", JsonValue.Create("middle")),
            ("textDirection", JsonValue.Create("vertical270")),
            ("marginLeft", JsonValue.Create("0.2cm")))));

        var detail = Get(inGroupPath);
        Assert.Equal("middle", detail["vAlign"]!.GetValue<string>());
        Assert.Equal("vertical270", detail["textDirection"]!.GetValue<string>());
        Assert.Equal(0.2, detail["marginLeft"]!.GetValue<double>());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddShape_AcceptsAnchoringPropsAtCreateTime_AndValidates()
    {
        var file = Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("text", "Created with anchoring"),
            ("name", JsonValue.Create("AddBox")),
            ("vAlign", JsonValue.Create("bottom")),
            ("textDirection", JsonValue.Create("vertical270")),
            ("marginLeft", JsonValue.Create("0.2cm")),
            ("marginRight", JsonValue.Create("0.2cm")),
            ("marginTop", JsonValue.Create("0.2cm")),
            ("marginBottom", JsonValue.Create("0.2cm")))));

        var bodyPr = BodyPr(file, "AddBox");
        Assert.Equal(A.TextAnchoringTypeValues.Bottom, bodyPr.Anchor!.Value);
        Assert.Equal(A.TextVerticalValues.Vertical270, bodyPr.Vertical!.Value);
        Assert.Equal(72000, bodyPr.LeftInset!.Value);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void TableCellPath_StillAcceptsTheSameProps_NoBehaviorChange()
    {
        Create();
        // Add a 2x2 table.
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "table", props: TestEnv.Props(
            ("rows", JsonValue.Create(2)),
            ("cols", JsonValue.Create(2)),
            ("x", JsonValue.Create("2cm")),
            ("y", JsonValue.Create("5cm")),
            ("w", JsonValue.Create("20cm")))));
        var tablePath = added["results"]![0]!["target"]!.GetValue<string>();
        var cellPath = $"{tablePath}/tr[1]/tc[1]";

        // The promoted-to-internal parsers must keep the table-cell path working unchanged.
        Edit(TestEnv.Op("set", cellPath, props: TestEnv.Props(
            ("valign", JsonValue.Create("middle")),
            ("textDirection", JsonValue.Create("vertical")),
            ("marginLeft", JsonValue.Create("0.2cm")))));

        var cell = Get(cellPath);
        Assert.Equal("middle", cell["vAlign"]!.GetValue<string>());
        Assert.Equal("vertical", cell["textDirection"]!.GetValue<string>());
        Assert.Equal(0.2, cell["marginLeft"]!.GetValue<double>());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
