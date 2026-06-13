using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

public sealed class AuditTests : IDisposable
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

    private AuditResult Audit(string category = "all", string minSeverity = "info") =>
        _handler.Audit(_ws.Ctx("deck.pptx"), new AuditOptions { Category = category, MinSeverity = minSeverity });

    private static bool Has(AuditResult result, string code) => result.Findings.Any(f => f.Code == code);

    private static AuditFinding Finding(AuditResult result, string code) => result.Findings.First(f => f.Code == code);

    private string AddShape(JsonObject props)
    {
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "shape", props: props));
        return data["results"]![0]!["target"]!.GetValue<string>();
    }

    private string WriteImage()
    {
        var imagePath = _ws.PathOf("logo.png");
        File.WriteAllBytes(imagePath, TestImages.Png(32, 32));
        return "logo.png";
    }

    // ---- clean deck ---------------------------------------------------------

    [Fact]
    public void CleanTitledSlide_HasNoFindings()
    {
        Create();
        // Slide 2: a real title placeholder plus one legible, high-contrast shape.
        Edit(TestEnv.Op("add", "/slide[1]", type: "slide", position: "after",
            props: TestEnv.Props(("title", JsonValue.Create("Agenda")))));
        Edit(TestEnv.Op("add", "/slide[2]", type: "shape", props: TestEnv.Props(
            ("text", JsonValue.Create("Body copy")),
            ("fontSize", JsonValue.Create(18)),
            ("color", JsonValue.Create("000000")),
            ("fill", JsonValue.Create("FFFFFF")),
            ("x", JsonValue.Create(2)),
            ("y", JsonValue.Create(5)))));

        var slide2 = Audit().Findings.Where(f => f.Path != null && f.Path.StartsWith("/slide[2]", StringComparison.Ordinal));
        Assert.Empty(slide2);
    }

    [Fact]
    public void FreshBlankSlide_FlagsMissingTitleOnly()
    {
        Create();
        var result = Audit("accessibility");
        Assert.True(Has(result, "a11y_no_slide_title"));
        Assert.False(Has(result, "a11y_no_alt_text"));
        Assert.False(Has(result, "a11y_tiny_font"));
    }

    // ---- alt text -----------------------------------------------------------

    [Fact]
    public void Picture_WithoutAltText_FiresAndIsAutofixable()
    {
        Create();
        var src = WriteImage();
        Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", JsonValue.Create(src)))));

        var result = Audit();
        Assert.True(Has(result, "a11y_no_alt_text"));
        Assert.True(Finding(result, "a11y_no_alt_text").Autofixable);
    }

    [Fact]
    public void Picture_WithAltText_IsSilent()
    {
        Create();
        var src = WriteImage();
        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", JsonValue.Create(src)))));
        var path = data["results"]![0]!["target"]!.GetValue<string>();
        Edit(TestEnv.Op("set", path, props: TestEnv.Props(("altText", JsonValue.Create("A logo")))));

        Assert.False(Has(Audit(), "a11y_no_alt_text"));
    }

    // ---- slide title --------------------------------------------------------

    [Fact]
    public void TitledSlide_IsSilentOnTitle()
    {
        Create();
        Edit(TestEnv.Op("add", "/slide[1]", type: "slide", position: "after",
            props: TestEnv.Props(("title", JsonValue.Create("Real Title")))));

        var result = _handler.Audit(_ws.Ctx("deck.pptx"), new AuditOptions { Category = "accessibility" });
        var slide2 = result.Findings.Where(f => f.Code == "a11y_no_slide_title" && f.Path == "/slide[2]");
        Assert.Empty(slide2);
    }

    // ---- tiny font ----------------------------------------------------------

    [Fact]
    public void TinyFont_BelowEightPt_IsError()
    {
        Create();
        AddShape(TestEnv.Props(("text", JsonValue.Create("microcopy")), ("fontSize", JsonValue.Create(6))));

        var f = Finding(Audit(), "a11y_tiny_font");
        Assert.Equal("error", f.Severity);
    }

    [Fact]
    public void TinyFont_BetweenEightAndTwelve_IsWarning()
    {
        Create();
        AddShape(TestEnv.Props(("text", JsonValue.Create("small")), ("fontSize", JsonValue.Create(10))));

        var f = Finding(Audit(), "a11y_tiny_font");
        Assert.Equal("warning", f.Severity);
    }

    [Fact]
    public void Font_AtTwelvePt_IsSilent()
    {
        Create();
        AddShape(TestEnv.Props(("text", JsonValue.Create("ok")), ("fontSize", JsonValue.Create(12))));

        Assert.False(Has(Audit(), "a11y_tiny_font"));
    }

    // ---- contrast -----------------------------------------------------------

    [Fact]
    public void LowContrast_GreyOnWhite_Fires()
    {
        Create();
        // Light grey text (#BBBBBB) on a white fill is far below 4.5:1.
        AddShape(TestEnv.Props(
            ("text", JsonValue.Create("faint")),
            ("fontSize", JsonValue.Create(18)),
            ("color", JsonValue.Create("BBBBBB")),
            ("fill", JsonValue.Create("FFFFFF"))));

        Assert.True(Has(Audit(), "a11y_low_contrast"));
    }

    [Fact]
    public void HighContrast_BlackOnWhite_IsSilent()
    {
        Create();
        AddShape(TestEnv.Props(
            ("text", JsonValue.Create("clear")),
            ("fontSize", JsonValue.Create(18)),
            ("color", JsonValue.Create("000000")),
            ("fill", JsonValue.Create("FFFFFF"))));

        Assert.False(Has(Audit(), "a11y_low_contrast"));
    }

    // ---- off canvas ---------------------------------------------------------

    [Fact]
    public void OffCanvasShape_Fires()
    {
        Create();
        AddShape(TestEnv.Props(
            ("text", JsonValue.Create("hidden")),
            ("x", JsonValue.Create("-60cm")),
            ("y", JsonValue.Create("-60cm")),
            ("w", JsonValue.Create(5)),
            ("h", JsonValue.Create(3))));

        Assert.True(Has(Audit("quality"), "quality_off_canvas"));
    }

    [Fact]
    public void OnCanvasShape_IsSilent()
    {
        Create();
        AddShape(TestEnv.Props(("text", JsonValue.Create("visible")), ("x", JsonValue.Create(2)), ("y", JsonValue.Create(2))));

        Assert.False(Has(Audit("quality"), "quality_off_canvas"));
    }

    // ---- reading order ------------------------------------------------------

    [Fact]
    public void ReadingOrderMismatch_FiresInfo()
    {
        Create();
        // Bottom-left shape added first, top shape added second: visual order
        // (top → bottom) disagrees with document order.
        AddShape(TestEnv.Props(("text", JsonValue.Create("bottom")), ("x", JsonValue.Create(2)), ("y", JsonValue.Create(15))));
        AddShape(TestEnv.Props(("text", JsonValue.Create("top")), ("x", JsonValue.Create(2)), ("y", JsonValue.Create(1))));

        var f = Finding(Audit("accessibility"), "a11y_reading_order");
        Assert.Equal("info", f.Severity);
        Assert.False(f.Autofixable);
    }

    [Fact]
    public void ReadingOrderInVisualOrder_IsSilent()
    {
        Create();
        AddShape(TestEnv.Props(("text", JsonValue.Create("top")), ("x", JsonValue.Create(2)), ("y", JsonValue.Create(1))));
        AddShape(TestEnv.Props(("text", JsonValue.Create("bottom")), ("x", JsonValue.Create(2)), ("y", JsonValue.Create(15))));

        Assert.False(Has(Audit("accessibility"), "a11y_reading_order"));
    }

    // ---- duplicate ids + empty placeholder (raw XML injection) --------------

    [Fact]
    public void DuplicateShapeIds_FireError()
    {
        var file = Create();
        InjectShape(file, id: 42, placeholder: null, text: "one");
        InjectShape(file, id: 42, placeholder: null, text: "two");

        var f = Finding(Audit("quality"), "quality_duplicate_id");
        Assert.Equal("error", f.Severity);
        Assert.False(f.Autofixable);
    }

    [Fact]
    public void EmptyBodyPlaceholder_FiresAndAutofixRemovesIt()
    {
        var file = Create();
        InjectShape(file, id: 50, placeholder: P.PlaceholderValues.Body, text: "");

        var before = Audit("quality");
        var finding = Finding(before, "quality_empty_placeholder");
        Assert.True(finding.Autofixable);

        var removed = _handler.Fix(_ws.Ctx("deck.pptx"), [finding.Id]);
        Assert.Equal(1, removed);
        Assert.False(Has(Audit("quality"), "quality_empty_placeholder"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- severity + category filtering --------------------------------------

    [Fact]
    public void SeverityFilter_DropsInfoFindings()
    {
        Create();
        AddShape(TestEnv.Props(("text", JsonValue.Create("bottom")), ("x", JsonValue.Create(2)), ("y", JsonValue.Create(15))));
        AddShape(TestEnv.Props(("text", JsonValue.Create("top")), ("x", JsonValue.Create(2)), ("y", JsonValue.Create(1))));

        Assert.True(Has(Audit(minSeverity: "info"), "a11y_reading_order"));
        Assert.False(Has(Audit(minSeverity: "warning"), "a11y_reading_order"));
    }

    [Fact]
    public void CategoryFilter_QualityExcludesAccessibility()
    {
        Create(); // blank slide → a11y_no_slide_title (accessibility)
        var quality = Audit("quality");
        Assert.DoesNotContain(quality.Findings, f => f.Category == "accessibility");
    }

    [Fact]
    public void Summary_CountsBySeverity()
    {
        Create();
        AddShape(TestEnv.Props(("text", JsonValue.Create("tiny")), ("fontSize", JsonValue.Create(6))));
        var result = Audit();
        Assert.Equal(result.Findings.Count(f => f.Severity == "error"), result.Summary.Errors);
        Assert.Equal(result.Findings.Count(f => f.Severity == "warning"), result.Summary.Warnings);
        Assert.Equal(result.Findings.Count(f => f.Severity == "info"), result.Summary.Infos);
    }

    // ---- --fix --------------------------------------------------------------

    [Fact]
    public void Fix_SetsAltTextOnPictures()
    {
        Create();
        var src = WriteImage();
        Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", JsonValue.Create(src)))));

        var fixed_ = _handler.Fix(_ws.Ctx("deck.pptx"), []);
        Assert.True(fixed_ >= 1);
        Assert.False(Has(Audit(), "a11y_no_alt_text"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Fix_AddsSlideTitle()
    {
        Create();
        Assert.True(Has(Audit("accessibility"), "a11y_no_slide_title"));

        _handler.Fix(_ws.Ctx("deck.pptx"), []);
        Assert.False(Has(Audit("accessibility"), "a11y_no_slide_title"));
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Fix_LeavesReportOnlyFindingsAlone()
    {
        Create();
        AddShape(TestEnv.Props(("text", JsonValue.Create("tiny")), ("fontSize", JsonValue.Create(6))));

        _handler.Fix(_ws.Ctx("deck.pptx"), []);
        // The tiny font is report-only; it survives a fix pass.
        Assert.True(Has(Audit(), "a11y_tiny_font"));
    }

    [Fact]
    public void Fix_TargetsSpecificFindingId()
    {
        Create();
        var src = WriteImage();
        Edit(TestEnv.Op("add", "/slide[1]", type: "image", props: TestEnv.Props(("src", JsonValue.Create(src)))));
        var altFinding = Finding(Audit(), "a11y_no_alt_text");

        var fixed_ = _handler.Fix(_ws.Ctx("deck.pptx"), [altFinding.Id]);
        Assert.Equal(1, fixed_);
        // The blank-slide title finding was NOT targeted, so it remains.
        Assert.True(Has(Audit("accessibility"), "a11y_no_slide_title"));
    }

    // ---- raw injection helper ----------------------------------------------

    /// <summary>Appends a shape directly to slide 1's spTree (used to craft cases the op surface can't).</summary>
    private static void InjectShape(string file, uint id, P.PlaceholderValues? placeholder, string text)
    {
        using var doc = PresentationDocument.Open(file, true);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        var tree = slidePart.Slide!.CommonSlideData!.ShapeTree!;

        var appNonVisual = new P.ApplicationNonVisualDrawingProperties();
        if (placeholder is { } ph)
        {
            appNonVisual.Append(new P.PlaceholderShape { Type = ph });
        }

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = $"Injected {id}" },
                new P.NonVisualShapeDrawingProperties(),
                appNonVisual),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 1_000_000L, Y = 1_000_000L },
                    new A.Extents { Cx = 2_000_000L, Cy = 1_000_000L }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.RunProperties { Language = "en-US" }, new A.Text(text)))));
        tree.Append(shape);
        slidePart.Slide.Save();
    }
}
