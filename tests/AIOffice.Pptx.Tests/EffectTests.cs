using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>v1.1 shape visual effects: shadow/glow/reflection (a:effectLst) and outline (a:ln).</summary>
public sealed class EffectTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() => TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    /// <summary>Adds a textbox and returns its shape path.</summary>
    private string AddShape()
    {
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "textbox", props: TestEnv.Props(
            ("text", "Effects"), ("x", "2cm"), ("y", "2cm"), ("w", "8cm"), ("h", "3cm"))));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private P.Shape SingleShape()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        return doc.PresentationPart!.SlideParts.Single().Slide!.Descendants<P.Shape>()
            .First(s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("TextBox", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void SetShadow_WritesOuterShadowAndReflects()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "404040"))));

        var shadow = SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!.GetFirstChild<A.OuterShadow>()!;
        Assert.Equal("404040", shadow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("404040", detail["effects"]!["shadow"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetGlow_WritesGlowAndReflects()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("glow", "FFD700"))));

        var glow = SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!.GetFirstChild<A.Glow>()!;
        Assert.Equal("FFD700", glow.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("FFD700", detail["effects"]!["glow"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetReflection_WritesReflectionAndReflects()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("reflection", true))));

        Assert.NotNull(SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!.GetFirstChild<A.Reflection>());

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.True(detail["effects"]!["reflection"]!.GetValue<bool>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetOutline_WritesLineAndReflects()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("outline", "FF0000"))));

        var outline = SingleShape().ShapeProperties!.GetFirstChild<A.Outline>()!;
        Assert.Equal("FF0000", outline.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Equal("FF0000", detail["effects"]!["outline"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetAllEffectsTogether_WritesEffectListInSchemaOrder()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(
            ("glow", "00FF00"), ("shadow", "303030"), ("reflection", true), ("outline", "0000FF"))));

        var shape = SingleShape();
        var effectList = shape.ShapeProperties!.GetFirstChild<A.EffectList>()!;
        // a:effectLst schema order: glow precedes outerShdw precedes reflection.
        var order = effectList.ChildElements.Select(e => e.LocalName).ToList();
        Assert.True(order.IndexOf("glow") < order.IndexOf("outerShdw"));
        Assert.True(order.IndexOf("outerShdw") < order.IndexOf("reflection"));
        Assert.NotNull(shape.ShapeProperties.GetFirstChild<A.Outline>());

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        var effects = detail["effects"]!;
        Assert.Equal("00FF00", effects["glow"]!.GetValue<string>());
        Assert.Equal("303030", effects["shadow"]!.GetValue<string>());
        Assert.True(effects["reflection"]!.GetValue<bool>());
        Assert.Equal("0000FF", effects["outline"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetShadowFalse_ClearsTheShadow()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "404040"))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", false))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));
        Assert.Null(detail["effects"]);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void SetShadow_IsIdempotent()
    {
        Create();
        var path = AddShape();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "404040"))));
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("shadow", "808080"))));

        // A second shadow set replaces the first, not stacks a duplicate.
        var effectList = SingleShape().ShapeProperties!.GetFirstChild<A.EffectList>()!;
        Assert.Single(effectList.Elements<A.OuterShadow>());
        Assert.Equal("808080", effectList.GetFirstChild<A.OuterShadow>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Effects_OnPicture_AlsoWork()
    {
        Create();
        File.WriteAllBytes(Path.Combine(_ws.Dir, "logo.png"), TestImages.Png(40, 30));
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", "logo.png"))));
        var picPath = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("set", picPath, props: TestEnv.Props(("shadow", "202020"))));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", picPath))));
        Assert.Equal("202020", detail["effects"]!["shadow"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
