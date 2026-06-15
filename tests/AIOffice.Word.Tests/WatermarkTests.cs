using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using V = DocumentFormat.OpenXml.Vml;

namespace AIOffice.Word.Tests;

public sealed class WatermarkTests : WordTestBase
{
    [Fact]
    public void Add_watermark_creates_default_header_and_vml_shape()
    {
        var file = CreateDoc(title: "Confidential doc");

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"text":"DRAFT"}}]""");
        Assert.Equal("/watermark[1]", Data(envelope)["ops"]!.AsArray()[0]!["path"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var part = Assert.Single(doc.MainDocumentPart!.HeaderParts);
            var shape = Assert.Single(part.Header!.Descendants<V.Shape>());
            Assert.StartsWith("PowerPlusWaterMarkObject", shape.Id!.Value!, StringComparison.Ordinal);
            Assert.Equal("#c0c0c0", shape.FillColor!.Value); // default color
            Assert.Contains("rotation:315", shape.Style!.Value!, StringComparison.Ordinal); // diagonal by default
            var textPath = Assert.Single(shape.Descendants<V.TextPath>());
            Assert.Equal("DRAFT", textPath.String!.Value);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Watermark_reuses_the_existing_header()
    {
        var file = CreateDoc(title: "Already headed");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"Quarterly"}}]""");

        Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"text":"INTERNAL","diagonal":false,"color":"FF0000"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var part = Assert.Single(doc.MainDocumentPart!.HeaderParts); // no second header created
            Assert.Contains("Quarterly", part.Header!.InnerText, StringComparison.Ordinal); // existing text kept
            var shape = Assert.Single(part.Header.Descendants<V.Shape>());
            Assert.Equal("#ff0000", shape.FillColor!.Value);
            Assert.DoesNotContain("rotation:315", shape.Style!.Value!, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_watermark_reports_text_color_diagonal()
    {
        var file = CreateDoc(title: "Got it");
        Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"text":"SECRET","color":"1F4E79"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/watermark[1]" })));

        Assert.Equal("watermark", got["type"]!.GetValue<string>());
        Assert.Equal("SECRET", got["properties"]!["text"]!.GetValue<string>());
        Assert.Equal("1F4E79", got["properties"]!["color"]!.GetValue<string>());
        Assert.True(got["properties"]!["diagonal"]!.GetValue<bool>());
        Assert.Equal(1, got["properties"]!["headers"]!.GetValue<int>());
    }

    [Fact]
    public void Remove_watermark_keeps_the_header()
    {
        var file = CreateDoc(title: "Cleanup");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"Kept"}}]""");
        Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"text":"GONE"}}]""");

        Edit(file, """[{"op":"remove","path":"/watermark[1]"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var part = Assert.Single(doc.MainDocumentPart!.HeaderParts);
            Assert.Empty(part.Header!.Descendants<V.Shape>());
            Assert.Contains("Kept", part.Header.InnerText, StringComparison.Ordinal);
        }

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/watermark[1]" })));
        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Second_watermark_is_invalid_args()
    {
        var file = CreateDoc(title: "One only");
        Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"text":"DRAFT"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"text":"FINAL"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("/watermark[1]", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Watermark_requires_text_and_targets_body()
    {
        var file = CreateDoc(title: "Strict");

        var noText = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, noText.Code);

        var wrongPath = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"watermark","props":{"text":"X"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, wrongPath.Code);
        Assert.Contains("/body", wrongPath.Candidates!);
    }

    [Fact]
    public void Tracked_watermark_add_is_unsupported()
    {
        var file = CreateDoc(title: "Tracked");

        var ex = Assert.Throws<AiofficeException>(() => Edit(
            file,
            """[{"op":"add","path":"/body","type":"watermark","props":{"text":"DRAFT"}}]""",
            new JsonObject { ["track"] = true }));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    // ----------------------------------------------- v1.7 picture watermark

    [Fact]
    public void Picture_watermark_embeds_an_image_in_the_header_and_validates()
    {
        var file = CreateDoc(title: "Logo doc");
        WritePng("logo.png", width: 200, height: 100);

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"image":"logo.png"}}]""");
        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("/watermark[1]", summary["path"]!.GetValue<string>());
        Assert.Equal("logo.png", summary["image"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var part = Assert.Single(doc.MainDocumentPart!.HeaderParts);
            var shape = Assert.Single(part.Header!.Descendants<V.Shape>());
            Assert.StartsWith("PowerPlusWaterMarkObject", shape.Id!.Value!, StringComparison.Ordinal);
            var imageData = Assert.Single(shape.Descendants<V.ImageData>());
            // Washed out by default: the grayscale gain/blacklevel are present.
            Assert.False(string.IsNullOrEmpty(imageData.Gain?.Value));
            // The referenced image part is embedded in the header.
            Assert.Single(part.ImageParts);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Picture_watermark_without_washout_omits_the_grayscale_gain()
    {
        var file = CreateDoc(title: "Crisp logo");
        WritePng("logo.png");

        Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"image":"logo.png","washout":false}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var imageData = Assert.Single(doc.MainDocumentPart!.HeaderParts.Single().Header!.Descendants<V.ImageData>());
        Assert.True(string.IsNullOrEmpty(imageData.Gain?.Value));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_picture_watermark_reports_kind_picture()
    {
        var file = CreateDoc(title: "Read it");
        WritePng("logo.png");
        Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"image":"logo.png"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/watermark[1]" })));
        Assert.Equal("picture", got["properties"]!["kind"]!.GetValue<string>());
        Assert.True(got["properties"]!["washout"]!.GetValue<bool>());
    }

    [Fact]
    public void Picture_watermark_image_must_resolve_in_the_sandbox()
    {
        var file = CreateDoc(title: "Escape attempt");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"image":"../../../etc/passwd"}}]"""));

        // Escaping the workspace is denied (sandbox), not a generic not-found.
        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
    }

    [Fact]
    public void Picture_watermark_missing_file_is_file_not_found()
    {
        var file = CreateDoc(title: "Missing");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"image":"nope.png"}}]"""));
        Assert.Equal(ErrorCodes.FileNotFound, ex.Code);
    }

    [Fact]
    public void Watermark_needs_text_or_image()
    {
        var file = CreateDoc(title: "Neither");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("image", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Remove_picture_watermark_clears_the_shape()
    {
        var file = CreateDoc(title: "Cleanup pic");
        WritePng("logo.png");
        Edit(file, """[{"op":"add","path":"/body","type":"watermark","props":{"image":"logo.png"}}]""");

        Edit(file, """[{"op":"remove","path":"/watermark[1]"}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Empty(doc.MainDocumentPart!.HeaderParts.Single().Header!.Descendants<V.Shape>());
        AssertValidatesClean(file);
    }
}
