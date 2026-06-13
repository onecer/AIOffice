using System.Globalization;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using AIOffice.Core;
using AIOffice.Core.Equations;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx;

/// <summary>
/// The M10 pptx equation surface: a LaTeX equation rendered as native OMML inside
/// a slide text box. The math is stored exactly the way PowerPoint writes it — an
/// <c>mc:AlternateContent</c> whose <c>mc:Choice Requires="a14"</c> carries an
/// <c>a14:m/m:oMathPara/m:oMath</c>, with an <c>mc:Fallback</c> plain run so
/// viewers without the math extension still show the source. The original LaTeX is
/// stored on the <c>m:oMath</c> as an <c>mc:Ignorable</c> vendor attribute so
/// read-back is faithful and the OpenXmlValidator passes with zero errors.
///
/// The OMML itself comes from the shared <see cref="OmmlMath"/> producer in
/// <c>AIOffice.Core.Equations</c> — the same converter the Word handler consumes,
/// so a given LaTeX string renders identically in both formats. xlsx has no math
/// object model (Excel has cell formulas, not equations), so equations are N/A
/// there; that is documented as an <c>unsupported_feature</c> on the Excel side.
/// </summary>
internal static class PptxEquations
{
    private static readonly XNamespace ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
    private static readonly XNamespace Aio = "urn:aioffice:equation";

    /// <summary>The local name of the vendor latex attribute stored on each m:oMath (mc:Ignorable).</summary>
    private const string LatexAttribute = "latex";

    private static readonly IReadOnlyList<string> AddPropKeys = ["latex", "x", "y", "w", "h", "fontSize"];

    /// <summary>One resolved equation on a slide: its host shape, the a:p that carries it and its 1-based index.</summary>
    internal sealed record EquationRef(ShapeView Shape, A.Paragraph Paragraph, XElement OMath, int Index)
    {
        public string? Latex => ReadLatex(OMath);
    }

    // ---- add ----------------------------------------------------------------

    /// <summary>
    /// Inserts a LaTeX equation as OMML into a slide text body and returns
    /// (canonicalPath, unknownTokens). When <paramref name="address"/> targets a
    /// shape, the equation is appended to that shape's text body; when it targets a
    /// slide, a new text box is created to hold it. Unknown LaTeX commands degrade
    /// to literal runs and are returned so the caller raises an equation_partial
    /// warning — exactly the docx contract.
    /// </summary>
    public static (string Path, IReadOnlyList<string> UnknownTokens) Add(
        PresentationPart presentation, PptxAddress address, JsonObject? props)
    {
        props ??= [];
        foreach (var (key, _) in props)
        {
            if (!AddPropKeys.Contains(key, StringComparer.Ordinal))
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown equation prop '{key}'.",
                    "Equation props: latex (required), x, y, w, h, fontSize. " +
                    "latex is the LaTeX source, e.g. {\"latex\":\"E = mc^2\"}; x/y/w/h place a new text box.",
                    candidates: AddPropKeys);
            }
        }

        if (!props.TryGetPropertyValue("latex", out var latexNode) || latexNode is null ||
            J.ScalarText(latexNode).Trim().Length == 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add equation requires props.latex with the equation source.",
                "Example: {\"op\":\"add\",\"path\":\"/slide[1]\",\"type\":\"equation\",\"props\":{\"latex\":\"E = mc^2\"}}. " +
                "Target /slide[i]/shape[@id=N] to append to an existing text box.");
        }

        var latex = J.ScalarText(latexNode);
        var (omath, unknown) = OmmlMath.FromLatex(latex);
        StoreLatex(omath, latex);

        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);

        var (shape, fontSize) = ResolveOrCreateHost(slidePart, address, props);
        var body = RequireTextBody(shape.Element);

        var paragraph = BuildEquationParagraph(omath, latex, fontSize);
        body.Append(paragraph);

        var index = CountEquations(body);
        return (address.CanonicalOMathPath(shape.Id, index), unknown);
    }

    /// <summary>
    /// Resolves the host shape: an existing shape when the address names one, or a
    /// freshly created text box on the slide otherwise. Returns the shape and an
    /// optional font size (hundredths of a point) for the equation fallback run.
    /// </summary>
    private static (ShapeView Shape, int? FontSize) ResolveOrCreateHost(
        SlidePart slidePart, PptxAddress address, JsonObject props)
    {
        var fontSize = props.TryGetPropertyValue("fontSize", out var sizeNode) && sizeNode is not null
            ? Units.ParseFontSizeHundredths("fontSize", sizeNode)
            : (int?)null;

        if (address.HasShape)
        {
            var resolved = PptxDoc.ResolveShape(slidePart, address);
            if (resolved.Element is not P.Shape)
            {
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"An equation attaches to a text shape, not a {resolved.Kind} at {resolved.CanonicalPath(address.SlideIndex)}.",
                    "Target a text box (shape) or the slide itself (/slide[i]) to add an equation.");
            }

            return (resolved, fontSize);
        }

        if (address.IsChart || address.IsTable || address.IsAnimation || address.IsComment ||
            address.IsEmbed || address.IsSmartArt || address.IsNotes)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add equation targets a slide or a text shape, not '{address.Raw}'.",
                "Use /slide[i] (creates a text box) or /slide[i]/shape[@id=N] (appends to an existing text box).");
        }

        // No shape addressed: create a text box to hold the equation.
        var tree = PptxDoc.RequireShapeTree(slidePart);
        var id = AddEquationTextBox(tree, props);
        var view = PptxDoc.Shapes(tree).First(s => s.Id == id);
        return (view, fontSize);
    }

    /// <summary>Creates an empty text box sized for an equation and returns its shape id.</summary>
    private static uint AddEquationTextBox(P.ShapeTree tree, JsonObject props)
    {
        var id = PptxDoc.NextShapeId(tree);

        var x = props.TryGetPropertyValue("x", out var xNode) && xNode is not null
            ? Units.ParseLengthEmu("x", xNode) : Units.CmToEmu(2.5);
        var y = props.TryGetPropertyValue("y", out var yNode) && yNode is not null
            ? Units.ParseLengthEmu("y", yNode) : Units.CmToEmu(2.5);
        var w = props.TryGetPropertyValue("w", out var wNode) && wNode is not null
            ? Units.ParseLengthEmu("w", wNode) : Units.CmToEmu(10);
        var h = props.TryGetPropertyValue("h", out var hNode) && hNode is not null
            ? Units.ParseLengthEmu("h", hNode) : Units.CmToEmu(3);

        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = Units.Inv($"Equation {id}") },
                new P.NonVisualShapeDrawingProperties { TextBox = true },
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = w, Cy = h }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(
                new A.BodyProperties { Wrap = A.TextWrappingValues.Square },
                new A.ListStyle()));

        tree.Append(shape);
        return id;
    }

    // ---- get / list ---------------------------------------------------------

    /// <summary>The metadata of one addressed equation: its canonical path, the stored latex, and display flag.</summary>
    public static object Detail(PresentationPart presentation, PptxAddress address)
    {
        var equation = Resolve(presentation, address);
        return new
        {
            Path = address.CanonicalOMathPath(equation.Shape.Id, equation.Index),
            Type = "equation",
            Latex = equation.Latex,
            Display = true,
        };
    }

    // ---- remove -------------------------------------------------------------

    /// <summary>Removes the addressed equation (the whole a:p carrying it). Returns the canonical path it occupied.</summary>
    public static string Remove(PresentationPart presentation, PptxAddress address)
    {
        var equation = Resolve(presentation, address);
        var canonical = address.CanonicalOMathPath(equation.Shape.Id, equation.Index);
        var body = equation.Paragraph.Parent as P.TextBody;
        equation.Paragraph.Remove();

        // A p:txBody must keep at least one a:p; if the equation was the last
        // paragraph, leave an empty one behind so the shape stays schema-valid.
        if (body is not null && !body.Elements<A.Paragraph>().Any())
        {
            body.Append(new A.Paragraph());
        }

        return canonical;
    }

    // ---- resolution ---------------------------------------------------------

    /// <summary>Resolves /slide[i]/shape[@id=N]/omath[k] to a live equation or throws invalid_path with candidates.</summary>
    private static EquationRef Resolve(PresentationPart presentation, PptxAddress address)
    {
        var slidePart = PptxDoc.ResolveSlide(presentation, address.SlideIndex, address.Raw);
        var shape = PptxDoc.ResolveShape(slidePart, address);
        if (shape.Element is not P.Shape || shape.Element.GetFirstChild<P.TextBody>() is not { } textBody)
        {
            throw EquationNotFound(address, shape, 0);
        }

        var equations = EquationsIn(shape, textBody);
        var index = address.OMathIndex ?? 1;
        if (index >= 1 && index <= equations.Count)
        {
            return equations[index - 1];
        }

        throw EquationNotFound(address, shape, equations.Count);
    }

    /// <summary>Every equation in a shape's text body, 1-based in document order.</summary>
    private static List<EquationRef> EquationsIn(ShapeView shape, P.TextBody body)
    {
        var result = new List<EquationRef>();
        var index = 0;
        foreach (var paragraph in body.Elements<A.Paragraph>())
        {
            if (FindOMath(paragraph) is { } omath)
            {
                index++;
                result.Add(new EquationRef(shape, paragraph, omath, index));
            }
        }

        return result;
    }

    private static int CountEquations(P.TextBody body) =>
        body.Elements<A.Paragraph>().Count(p => FindOMath(p) is not null);

    /// <summary>The m:oMath element a math paragraph carries (inside its AlternateContent/a14:m), or null.</summary>
    private static XElement? FindOMath(A.Paragraph paragraph)
    {
        var xml = XElement.Parse(paragraph.OuterXml);
        return xml.Descendants(OmmlMath.M + "oMath").FirstOrDefault();
    }

    private static AiofficeException EquationNotFound(PptxAddress address, ShapeView shape, int count) => new(
        ErrorCodes.InvalidPath,
        count == 0
            ? Units.Inv($"{shape.CanonicalPath(address.SlideIndex)} has no equations.")
            : Units.Inv($"{address.Raw} does not exist; the shape has {count} equation(s)."),
        "Address an equation as /slide[i]/shape[@id=N]/omath[k] (1-based). " +
        "Add one with {\"op\":\"add\",\"type\":\"equation\",\"props\":{\"latex\":\"…\"}}.",
        candidates: [.. Enumerable.Range(1, count).Select(n => address.CanonicalOMathPath(shape.Id, n))]);

    // ---- xml construction ---------------------------------------------------

    /// <summary>
    /// Builds the <c>a:p</c> that carries one equation: an mc:AlternateContent with
    /// an a14 math choice (the real OMML) and a plain-run fallback (the latex
    /// source). Loaded into a <see cref="A.Paragraph"/> from its serialized form.
    /// </summary>
    private static A.Paragraph BuildEquationParagraph(XElement omath, string latex, int? fontSize)
    {
        var omathPara = new XElement(OmmlMath.M + "oMathPara", omath);

        var fallbackRun = new XElement(
            ANs + "r",
            fontSize is { } size
                ? new XElement(ANs + "rPr",
                    new XAttribute("lang", "en-US"),
                    new XAttribute("sz", size.ToString(CultureInfo.InvariantCulture)))
                : null,
            new XElement(ANs + "t", latex));

        var paragraph = new XElement(
            ANs + "p",
            new XElement(
                MC + "AlternateContent",
                new XAttribute(XNamespace.Xmlns + "mc", MC.NamespaceName),
                new XElement(
                    MC + "Choice",
                    new XAttribute("Requires", "a14"),
                    new XAttribute(XNamespace.Xmlns + "a14", A14.NamespaceName),
                    new XElement(A14 + "m", omathPara)),
                new XElement(MC + "Fallback", fallbackRun)));

        return new A.Paragraph(paragraph.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>Stores the original LaTeX on an m:oMath as an mc:Ignorable vendor attribute for faithful read-back.</summary>
    private static void StoreLatex(XElement omath, string latex)
    {
        omath.SetAttributeValue(XNamespace.Xmlns + "aio", Aio.NamespaceName);
        omath.SetAttributeValue(Aio + LatexAttribute, latex);
        omath.SetAttributeValue(MC + "Ignorable", "aio");
    }

    /// <summary>The stored LaTeX of an equation, or null when it came from another tool.</summary>
    private static string? ReadLatex(XElement omath) => omath.Attribute(Aio + LatexAttribute)?.Value;

    /// <summary>The text body of a shape, or invalid_args when the shape has none.</summary>
    private static P.TextBody RequireTextBody(OpenXmlCompositeElement shape) =>
        shape.GetFirstChild<P.TextBody>()
        ?? throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            "The target shape has no text body to hold an equation.",
            "Add the equation to /slide[i] (a new text box) instead.");
}
