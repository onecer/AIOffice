using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// Covers the additive shape body prop "autofit", which drives the a:bodyPr autofit
/// child: "shrink" -> a:normAutofit, "resize" -> a:spAutoFit, "none" -> a:noAutofit.
/// </summary>
public sealed class AutofitTests : IDisposable
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
            ("text", "Some text that might overflow"),
            ("name", JsonValue.Create(name)))));
        return added["results"]![0]!["target"]!.GetValue<string>();
    }

    [Theory]
    [InlineData("shrink")]
    [InlineData("resize")]
    [InlineData("none")]
    public void SetAutofit_RoundTripsModeAndWritesTheRightElement(string mode)
    {
        var file = Create();
        var path = AddShape();

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("autofit", JsonValue.Create(mode)))));

        // get reports the mode back.
        var detail = Get(path);
        Assert.Equal(mode, detail["autofit"]!["mode"]!.GetValue<string>());

        // The saved bodyPr carries exactly the matching autofit element.
        var bodyPr = BodyPr(file, "Box");
        switch (mode)
        {
            case "shrink":
                Assert.NotNull(bodyPr.GetFirstChild<A.NormalAutoFit>());
                Assert.Null(bodyPr.GetFirstChild<A.ShapeAutoFit>());
                Assert.Null(bodyPr.GetFirstChild<A.NoAutoFit>());
                break;
            case "resize":
                Assert.NotNull(bodyPr.GetFirstChild<A.ShapeAutoFit>());
                Assert.Null(bodyPr.GetFirstChild<A.NormalAutoFit>());
                Assert.Null(bodyPr.GetFirstChild<A.NoAutoFit>());
                break;
            case "none":
                Assert.NotNull(bodyPr.GetFirstChild<A.NoAutoFit>());
                Assert.Null(bodyPr.GetFirstChild<A.NormalAutoFit>());
                Assert.Null(bodyPr.GetFirstChild<A.ShapeAutoFit>());
                break;
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void BareShrink_WritesPlainNormAutofit_NoScaleAttributes()
    {
        var file = Create();
        var path = AddShape();

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("autofit", JsonValue.Create("shrink")))));

        var norm = BodyPr(file, "Box").GetFirstChild<A.NormalAutoFit>()!;
        Assert.Null(norm.FontScale);
        Assert.Null(norm.LineSpaceReduction);

        // get omits the scale fields when the normAutofit carries none.
        var autofit = Get(path)["autofit"]!.AsObject();
        Assert.False(autofit.ContainsKey("fontScale"));
        Assert.False(autofit.ContainsKey("lineSpaceReduction"));
    }

    [Fact]
    public void ShrinkObject_WritesExplicitScaleInThousandths_AndRoundTripsAsPercent()
    {
        var file = Create();
        var path = AddShape();

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("autofit", new JsonObject
        {
            ["mode"] = "shrink",
            ["fontScale"] = 90,
            ["lineSpaceReduction"] = 10,
        }))));

        // OOXML stores thousandths of a percent: 90% -> 90000, 10% -> 10000.
        var norm = BodyPr(file, "Box").GetFirstChild<A.NormalAutoFit>()!;
        Assert.Equal(90000, norm.FontScale!.Value);
        Assert.Equal(10000, norm.LineSpaceReduction!.Value);

        // get reports the percentages back.
        var autofit = Get(path)["autofit"]!;
        Assert.Equal("shrink", autofit["mode"]!.GetValue<string>());
        Assert.Equal(90.0, autofit["fontScale"]!.GetValue<double>());
        Assert.Equal(10.0, autofit["lineSpaceReduction"]!.GetValue<double>());

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddShape_AcceptsAutofitAtCreateTime()
    {
        var file = Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: TestEnv.Props(
            ("text", "Created with autofit"),
            ("name", JsonValue.Create("AddBox")),
            ("autofit", JsonValue.Create("shrink")))));

        Assert.NotNull(BodyPr(file, "AddBox").GetFirstChild<A.NormalAutoFit>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SwitchingModes_ReplacesTheChild_NeverDuplicates()
    {
        var file = Create();
        var path = AddShape();

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("autofit", JsonValue.Create("shrink")))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("autofit", JsonValue.Create("resize")))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("autofit", JsonValue.Create("none")))));

        var bodyPr = BodyPr(file, "Box");
        // A bodyPr holds at most one autofit child; the last write wins and is unique.
        var autofitChildren = bodyPr.ChildElements
            .Count(c => c is A.NormalAutoFit or A.ShapeAutoFit or A.NoAutoFit);
        Assert.Equal(1, autofitChildren);
        Assert.NotNull(bodyPr.GetFirstChild<A.NoAutoFit>());

        Assert.Equal("none", Get(path)["autofit"]!["mode"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ShrinkObjectReplacesPlainShrink_UpgradingToExplicitScale()
    {
        var file = Create();
        var path = AddShape();

        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("autofit", JsonValue.Create("shrink")))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("autofit", new JsonObject
        {
            ["mode"] = "shrink",
            ["fontScale"] = 75,
        }))));

        var bodyPr = BodyPr(file, "Box");
        Assert.Equal(1, bodyPr.ChildElements.Count(c => c is A.NormalAutoFit));
        Assert.Equal(75000, bodyPr.GetFirstChild<A.NormalAutoFit>()!.FontScale!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void NoAutofit_OnAShapeWithoutOne_ReportsNullAutofit()
    {
        Create();
        var path = AddShape();

        // A freshly added shape (no autofit set) reports no autofit child.
        var detail = Get(path);
        Assert.True(detail["autofit"] is null);
    }

    [Fact]
    public void UnknownAutofitMode_IsRejectedWithInvalidArgs()
    {
        Create();
        var path = AddShape();

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("autofit", JsonValue.Create("stretch"))))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void OutOfRangeFontScale_IsRejectedWithInvalidArgs()
    {
        Create();
        var path = AddShape();

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("set", path, props: TestEnv.Props(("autofit", new JsonObject
            {
                ["mode"] = "shrink",
                ["fontScale"] = 150,
            })))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }
}
