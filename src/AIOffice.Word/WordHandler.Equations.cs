using System.Text.Json.Nodes;
using AIOffice.Core;
using AIOffice.Word.Equations;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>
    /// The vendor namespace under which the original LaTeX source is stored on
    /// every emitted <c>m:oMath</c> (as an <c>mc:Ignorable</c> foreign attribute,
    /// so Word ignores it, the validator passes, and read-back is faithful).
    /// </summary>
    private const string EquationNamespace = "urn:aioffice:equation";
    private const string EquationPrefix = "aio";
    private const string EquationLatexAttribute = "latex";
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string MarkupCompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    /// <summary>
    /// <c>{"op":"add","path":"/body/p[2]","type":"equation","props":{"latex":"E=mc^2","display":false}}</c>.
    /// Inline (display:false) appends an <c>m:oMath</c> to the target paragraph;
    /// display:true emits its own centered equation paragraph (<c>m:oMathPara</c>),
    /// placed before/after/inside like other block adds. The original LaTeX is
    /// stored on the equation for faithful read-back; unrecognized commands degrade
    /// to literal runs and raise an <c>equation_partial</c> meta warning.
    /// </summary>
    private static object ApplyAddEquation(WordprocessingDocument doc, EditOp op, EditSession session)
    {
        var props = op.Props ?? throw EquationNeedsLatex();
        var latex = props["latex"] is { } latexNode ? NodeToString(latexNode) : throw EquationNeedsLatex();
        if (string.IsNullOrWhiteSpace(latex))
        {
            throw EquationNeedsLatex();
        }

        var display = props["display"] is { } displayNode && IsTrue(displayNode);

        var parse = LatexParser.Parse(latex);
        if (parse.UnknownTokens.Count > 0)
        {
            session.Warnings.Add(new Warning(
                "equation_partial",
                $"The equation rendered, but these LaTeX tokens were not recognized and appear literally: " +
                $"{string.Join(", ", parse.UnknownTokens.Distinct())}. " +
                "Check the spelling, or split the unsupported part into a \\text{…} run."));
        }

        // The math + mc + vendor namespaces are declared once on the document root,
        // in the order the SDK normalizes to, so the very first save already matches
        // the reopen-and-resave form — the round-trip law stays byte-exact.
        EnsureEquationNamespaces(doc);

        var oMath = OmmlEmitter.ToOfficeMath(parse.Root);
        StoreLatex(oMath, latex);

        if (display)
        {
            return AddDisplayEquation(doc, op, oMath, latex);
        }

        return AddInlineEquation(doc, op, oMath, latex);
    }

    /// <summary>Appends an inline <c>m:oMath</c> to the resolved paragraph; returns its /omath[j] path.</summary>
    private static object AddInlineEquation(WordprocessingDocument doc, EditOp op, M.OfficeMath oMath, string latex)
    {
        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));
        if (anchor.Element is not Paragraph paragraph)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"An inline equation attaches to a paragraph, not '{anchor.Type}' at {anchor.CanonicalPath}.",
                "Target a paragraph (e.g. /body/p[2]) for an inline equation, or pass display:true to add a centered equation block.");
        }

        // m:oMath sits after the paragraph's runs; pPr (if any) must stay first.
        paragraph.AppendChild(oMath);

        var index = paragraph.ChildElements.OfType<M.OfficeMath>().Count();
        return new
        {
            op = "add",
            type = "equation",
            path = $"{anchor.CanonicalPath}/omath[{index}]",
            display = false,
            latex,
        };
    }

    /// <summary>
    /// Emits a centered display equation as its own body paragraph carrying an
    /// <c>m:oMathPara</c>. Placement honors position before/after/inside exactly
    /// like other block adds (default: a new paragraph after the anchor, or inside
    /// a container anchor).
    /// </summary>
    private static object AddDisplayEquation(WordprocessingDocument doc, EditOp op, M.OfficeMath oMath, string latex)
    {
        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));

        var mathParagraph = new M.Paragraph(
            new M.ParagraphProperties(new M.Justification { Val = M.JustificationValues.Center }),
            oMath);
        var paragraph = new Paragraph(mathParagraph);

        var isContainerAnchor = anchor.Type is "body" or "tc" or "header" or "footer";
        var position = op.Position;
        if (position is not (null or "before" or "after" or "inside"))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add position '{position}' is not valid for a display equation.",
                "Use position before, after or inside.",
                candidates: ["before", "after", "inside"]);
        }

        var pos = position ?? (isContainerAnchor ? "inside" : "after");
        var valid = (anchor.Type, pos) switch
        {
            ("body" or "tc" or "header" or "footer", "inside") => true,
            ("p" or "table", "before" or "after") => true,
            _ => false,
        };

        if (!valid)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Cannot add a display equation {pos} {anchor.CanonicalPath} ({anchor.Type}).",
                "Add a display equation inside /body, a tc, a header or footer, or before/after an existing p/table.");
        }

        var canonical = Insert(doc, anchor, paragraph, pos);
        return new
        {
            op = "add",
            type = "equation",
            path = canonical + "/omath[1]",
            display = true,
            latex,
        };
    }

    // ----------------------------------------------------------- get / remove

    /// <summary>
    /// Resolves an inline-equation path <c>/…/p[i]/omath[j]</c> (1-based) to its
    /// <c>m:oMath</c>, with invalid_path + candidate paths on miss. A display
    /// equation's <c>m:oMathPara</c> wrapper is transparent: <c>omath[1]</c>
    /// addresses the inner <c>m:oMath</c>.
    /// </summary>
    private static (M.OfficeMath OMath, Paragraph Paragraph, string Canonical) ResolveEquation(
        WordprocessingDocument doc, DocPath path)
    {
        if (path.Segments.Count < 2 || path.Segments[^1].Name != "omath")
        {
            throw EquationPathInvalid(path.ToCanonicalString(), "Equation paths end with /omath[j].");
        }

        var parentPath = new DocPath { Segments = [.. path.Segments.Take(path.Segments.Count - 1)] };
        var parent = WordAddress.Resolve(doc, parentPath);
        if (parent.Element is not Paragraph paragraph)
        {
            throw EquationPathInvalid(path.ToCanonicalString(), $"{parent.CanonicalPath} is a {parent.Type}, not a paragraph.");
        }

        var equations = ParagraphEquations(paragraph);
        var index = path.Segments[^1].Index ?? 1;
        if (index < 1 || index > equations.Count)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"{parent.CanonicalPath}/omath[{index}] does not exist; the paragraph has {equations.Count} equation(s).",
                equations.Count == 0
                    ? "This paragraph has no equations. Add one with {\"op\":\"add\",\"type\":\"equation\",\"props\":{\"latex\":\"…\"}}."
                    : "Use a 1-based omath index within range.",
                candidates: [.. Enumerable.Range(1, equations.Count).Select(n => $"{parent.CanonicalPath}/omath[{n}]")]);
        }

        return (equations[index - 1], paragraph, $"{parent.CanonicalPath}/omath[{index}]");
    }

    /// <summary>Every <c>m:oMath</c> directly under a paragraph, including the one inside a display <c>m:oMathPara</c>.</summary>
    private static List<M.OfficeMath> ParagraphEquations(Paragraph paragraph)
    {
        var result = new List<M.OfficeMath>();
        foreach (var child in paragraph.ChildElements)
        {
            switch (child)
            {
                case M.OfficeMath inline:
                    result.Add(inline);
                    break;
                case M.Paragraph mathPara:
                    result.AddRange(mathPara.ChildElements.OfType<M.OfficeMath>());
                    break;
                default:
                    break;
            }
        }

        return result;
    }

    /// <summary>get /…/omath[j]: the stored LaTeX source and whether it is a display equation.</summary>
    private static Dictionary<string, object?> GetEquationProperties(WordprocessingDocument doc, DocPath path)
    {
        var (oMath, _, canonical) = ResolveEquation(doc, path);
        return new Dictionary<string, object?>
        {
            ["path"] = canonical,
            ["latex"] = ReadStoredLatex(oMath),
            ["display"] = oMath.Parent is M.Paragraph,
        };
    }

    /// <summary>remove /…/omath[j]: drops an inline equation, or the whole display paragraph it owns.</summary>
    private static object ApplyRemoveEquation(WordprocessingDocument doc, EditOp op)
    {
        var path = DocPath.Parse(op.Path);
        var (oMath, paragraph, canonical) = ResolveEquation(doc, path);

        if (oMath.Parent is M.Paragraph mathPara)
        {
            // Display equation: remove the math paragraph; if that empties the body
            // paragraph entirely, drop the paragraph too (unless it is the last block).
            mathPara.Remove();
            if (!paragraph.ChildElements.Any(c => c is not ParagraphProperties) &&
                paragraph.Parent is Body { ChildElements.Count: > 1 })
            {
                paragraph.Remove();
            }
        }
        else
        {
            oMath.Remove();
        }

        return new { op = "remove", path = canonical, type = "equation" };
    }

    // ------------------------------------------------------------- latex store

    /// <summary>
    /// Declares the math (m:), markup-compatibility (mc:) and vendor (aio:)
    /// namespaces on the root <c>w:document</c>. Declaring them up front, in the
    /// SDK's normalized order, keeps the no-edit reopen byte-identical (the SDK
    /// otherwise hoists m:/mc: to the root on the next save and reorders them).
    /// </summary>
    private static void EnsureEquationNamespaces(WordprocessingDocument doc)
    {
        if (doc.MainDocumentPart?.Document is not { } document)
        {
            return;
        }

        var declared = document.NamespaceDeclarations.Select(d => d.Value).ToHashSet(StringComparer.Ordinal);
        if (!declared.Contains(MathNamespace))
        {
            document.AddNamespaceDeclaration("m", MathNamespace);
        }

        if (!declared.Contains(EquationNamespace))
        {
            document.AddNamespaceDeclaration(EquationPrefix, EquationNamespace);
        }

        if (!declared.Contains(MarkupCompatibilityNamespace))
        {
            document.AddNamespaceDeclaration("mc", MarkupCompatibilityNamespace);
        }
    }

    /// <summary>
    /// Stores the LaTeX source on an <c>m:oMath</c> as an mc:Ignorable foreign
    /// attribute. The referenced namespaces are declared on the document root by
    /// <see cref="EnsureEquationNamespaces"/>.
    /// </summary>
    private static void StoreLatex(M.OfficeMath oMath, string latex)
    {
        oMath.SetAttribute(new OpenXmlAttribute(EquationPrefix, EquationLatexAttribute, EquationNamespace, latex));
        oMath.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = EquationPrefix };
    }

    /// <summary>The stored LaTeX, or null when the equation was authored outside aioffice.</summary>
    private static string? ReadStoredLatex(M.OfficeMath oMath) =>
        oMath.GetAttributes().FirstOrDefault(a =>
            a.LocalName == EquationLatexAttribute && a.NamespaceUri == EquationNamespace).Value;

    /// <summary>
    /// A read-back rendering of one equation. When the original LaTeX is stored,
    /// it is returned verbatim; otherwise the equation came from another tool and
    /// we fall back to its OMML plain text (the math glyphs), which still reads
    /// sensibly even though it is not source LaTeX.
    /// </summary>
    internal static string EquationLatex(M.OfficeMath oMath) =>
        ReadStoredLatex(oMath) ?? oMath.InnerText;

    private static AiofficeException EquationNeedsLatex() => new(
        ErrorCodes.InvalidArgs,
        "add --type equation needs props.latex with the equation source.",
        "Example: {\"op\":\"add\",\"path\":\"/body/p[1]\",\"type\":\"equation\",\"props\":{\"latex\":\"E = mc^2\"}}. " +
        "Pass display:true for a centered equation block.");

    private static AiofficeException EquationPathInvalid(string path, string detail) => new(
        ErrorCodes.InvalidPath,
        $"'{path}' is not an equation path: {detail}",
        "Address an inline equation as /body/p[i]/omath[j] (1-based). Run query or read --view text to find equations.");

    private static bool IsTrue(JsonNode node)
    {
        if (node is JsonValue v)
        {
            if (v.TryGetValue<bool>(out var b))
            {
                return b;
            }

            if (v.TryGetValue<string>(out var s))
            {
                return s.Equals("true", StringComparison.OrdinalIgnoreCase) || s == "1";
            }
        }

        return false;
    }
}
