using System.Text.Json.Nodes;
using System.Xml.Linq;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// M10 pptx equations: a LaTeX equation rendered as native OMML inside a slide
/// text box (an mc:AlternateContent / a14:m / m:oMathPara / m:oMath, with a plain
/// fallback run). The tests cover add (on a slide and on a shape), the latex
/// round-trip via get, the equation_partial warning for unknown commands, remove,
/// and validator-cleanliness — all platform-independent. The OMML comes from the
/// shared AIOffice.Core.Equations converter the Word handler also uses.
/// </summary>
public sealed class EquationTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() =>
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Math"))));

    private Envelope EditRaw(params EditOp[] ops) => _handler.Edit(_ws.Ctx("deck.pptx"), ops);

    private JsonObject Edit(params EditOp[] ops) => TestEnv.AssertOk(EditRaw(ops));

    private JsonObject Get(string path) => TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", path))));

    private static readonly XNamespace MathNs = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>Counts m:oMath / m:f elements in the first slide's raw XML (the OMML sits inside an mc:AlternateContent the typed model does not descend into).</summary>
    private (int OMath, int Fractions) MathCount()
    {
        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var slidePart = doc.PresentationPart!.SlideParts.Single();
        var xml = XElement.Parse(slidePart.Slide!.OuterXml);
        return (
            xml.Descendants(MathNs + "oMath").Count(),
            xml.Descendants(MathNs + "f").Count());
    }

    // ---- add ----------------------------------------------------------------

    [Fact]
    public void AddEquation_OnSlide_CreatesTextBox_WithOmmlMath_AndValidates()
    {
        Create();

        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "equation", props: TestEnv.Props(
            ("latex", "x = \\frac{1}{2}"))));
        var canonical = data["results"]![0]!["target"]!.GetValue<string>();
        Assert.Matches(@"^/slide\[1\]/shape\[@id=[0-9]+\]/omath\[1\]$", canonical);

        // The real OMML element model is present (m:oMath -> m:f fraction).
        var (oMath, fractions) = MathCount();
        Assert.Equal(1, oMath);
        Assert.Equal(1, fractions);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddEquation_WithPlacementAndFontSize_PlacesANewTextBox()
    {
        Create();

        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "equation", props: TestEnv.Props(
            ("latex", "E = mc^2"),
            ("x", "1cm"),
            ("y", "4cm"),
            ("w", "10cm"),
            ("h", "3cm"),
            ("fontSize", System.Text.Json.Nodes.JsonValue.Create(32)))));
        var canonical = data["results"]![0]!["target"]!.GetValue<string>();
        Assert.Matches(@"^/slide\[1\]/shape\[@id=[0-9]+\]/omath\[1\]$", canonical);

        Assert.Equal(1, MathCount().OMath);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddEquation_OnExistingShape_AppendsToThatTextBox()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "textbox", props: TestEnv.Props(("text", "Given:"))));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();

        var data = Edit(TestEnv.Op("add", shapePath, type: "equation", props: TestEnv.Props(("latex", "E = mc^2"))));
        var canonical = data["results"]![0]!["target"]!.GetValue<string>();
        Assert.Equal(shapePath + "/omath[1]", canonical);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- get / latex round-trip --------------------------------------------

    [Fact]
    public void GetEquation_ReportsTheOriginalLatex()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "equation", props: TestEnv.Props(
            ("latex", "\\sqrt{a^2 + b^2}"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        var got = Get(path);
        Assert.Equal("equation", got["type"]!.GetValue<string>());
        Assert.Equal("\\sqrt{a^2 + b^2}", got["latex"]!.GetValue<string>());
    }

    [Fact]
    public void Latex_SurvivesAnOpenSaveRoundTrip()
    {
        Create();
        const string latex = "\\sum_{i=1}^{n} i = \\frac{n(n+1)}{2}";
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "equation", props: TestEnv.Props(("latex", latex))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        // A no-op-ish edit (add an unrelated text box) forces an open+save cycle.
        Edit(TestEnv.Op("add", "/slide[1]", type: "textbox", props: TestEnv.Props(("text", "x"))));

        var got = Get(path);
        Assert.Equal(latex, got["latex"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- partial warning ----------------------------------------------------

    [Fact]
    public void UnknownCommands_RaiseEquationPartial_ButStillSucceedAndValidate()
    {
        Create();

        var envelope = EditRaw(TestEnv.Op("add", "/slide[1]", type: "equation", props: TestEnv.Props(
            ("latex", "\\foobar + x"))));
        TestEnv.AssertOk(envelope);

        var partial = Assert.Single(envelope.Meta.Warnings!, w => w.Code == "equation_partial");
        Assert.Contains("\\foobar", partial.Message);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void KnownEquation_RaisesNoWarning()
    {
        Create();
        var envelope = EditRaw(TestEnv.Op("add", "/slide[1]", type: "equation", props: TestEnv.Props(
            ("latex", "a + b = c"))));
        TestEnv.AssertOk(envelope);
        Assert.Null(envelope.Meta.Warnings);
    }

    // ---- remove -------------------------------------------------------------

    [Fact]
    public void RemoveEquation_DropsIt_AndLeavesAValidDeck()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "equation", props: TestEnv.Props(("latex", "E = mc^2"))));
        var path = added["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("remove", path));

        Assert.Equal(0, MathCount().OMath);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- errors -------------------------------------------------------------

    [Fact]
    public void AddEquation_WithoutLatex_IsInvalidArgs()
    {
        Create();
        var envelope = EditRaw(TestEnv.Op("add", "/slide[1]", type: "equation", props: TestEnv.Props(("x", "1cm"))));
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void GetMissingEquation_IsInvalidPath()
    {
        Create();
        var added = Edit(TestEnv.Op("add", "/slide[1]", type: "textbox", props: TestEnv.Props(("text", "no math"))));
        var shapePath = added["results"]![0]!["target"]!.GetValue<string>();

        var envelope = _handler.Get(_ws.Ctx("deck.pptx", ("path", shapePath + "/omath[1]")));
        TestEnv.AssertFail(envelope, ErrorCodes.InvalidPath);
    }
}
