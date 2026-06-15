using System.Text.Json.Nodes;
using AIOffice.Core;
using AIOffice.Core.Equations;
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

        // (1.7) props.number turns a display equation into a numbered display
        // equation: number:true auto-increments; number:"(1.1)" sets the label
        // verbatim. A number on an inline equation is a usage error (inline
        // equations are not the numbered-equation pattern).
        var numberRequest = ParseEquationNumberRequest(props);
        if (numberRequest is not null && !display)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "props.number applies to a display equation; pass display:true alongside it.",
                "Example: {\"op\":\"add\",\"path\":\"/body\",\"type\":\"equation\"," +
                "\"props\":{\"latex\":\"E=mc^2\",\"display\":true,\"number\":true}}.");
        }

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
            return AddDisplayEquation(doc, op, oMath, latex, numberRequest, session);
        }

        return AddInlineEquation(doc, op, oMath, latex);
    }

    /// <summary>
    /// The numbering request behind props.number: <c>Auto</c> (number:true →
    /// next sequence value), an explicit <c>Label</c> (number:"(1.1)"), or null
    /// when the key is absent / false.
    /// </summary>
    private sealed record EquationNumberRequest(bool Auto, string? Label);

    private static EquationNumberRequest? ParseEquationNumberRequest(JsonObject props)
    {
        if (!props.TryGetPropertyValue("number", out var node) || node is null)
        {
            return null;
        }

        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var flag))
            {
                return flag ? new EquationNumberRequest(Auto: true, Label: null) : null;
            }

            if (value.TryGetValue<string>(out var text))
            {
                var trimmed = text.Trim();
                return trimmed switch
                {
                    "" or "false" or "0" => null,
                    "true" or "1" => new EquationNumberRequest(Auto: true, Label: null),
                    _ => new EquationNumberRequest(Auto: false, Label: trimmed),
                };
            }

            if (value.TryGetValue<double>(out var num))
            {
                // A bare number is taken as the literal label, e.g. number:7 → "7".
                return new EquationNumberRequest(
                    Auto: false,
                    Label: num.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }

        return new EquationNumberRequest(Auto: true, Label: null);
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
    /// a container anchor). When <paramref name="numberRequest"/> is set, the
    /// numbered-equation pattern is used instead: tab-aligned with a right-aligned
    /// number label and the number cached on the equation for read-back.
    /// </summary>
    private static object AddDisplayEquation(
        WordprocessingDocument doc,
        EditOp op,
        M.OfficeMath oMath,
        string latex,
        EquationNumberRequest? numberRequest,
        EditSession session)
    {
        var anchor = WordAddress.Resolve(doc, DocPath.Parse(op.Path));

        string? number = null;
        Paragraph paragraph;
        if (numberRequest is not null)
        {
            number = numberRequest.Auto ? NextEquationNumber(doc) : numberRequest.Label!;
            StoreEquationNumber(oMath, number);
            paragraph = BuildNumberedEquationParagraph(oMath, number);
            session.Warnings.Add(new Warning(
                WarningCodes.EquationNumbersCached,
                $"The display equation was numbered {number} from the cached sequence. " +
                "If the document uses Word SEQ fields elsewhere, update fields (F9) so all numbers agree."));
        }
        else
        {
            paragraph = new Paragraph(new M.Paragraph(
                new M.ParagraphProperties(new M.Justification { Val = M.JustificationValues.Center }),
                oMath));
        }

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
        return number is null
            ? new
            {
                op = "add",
                type = "equation",
                path = canonical + "/omath[1]",
                display = true,
                latex,
            }
            : (object)new
            {
                op = "add",
                type = "equation",
                path = canonical + "/omath[1]",
                display = true,
                latex,
                number,
            };
    }

    /// <summary>
    /// The standard Word numbered-equation paragraph: the equation sits centered
    /// (a center tab at the column mid-point) and the number label is right-aligned
    /// (a right tab at the right margin), the tab-stop pattern Word's own
    /// "Insert ▸ Equation ▸ numbered" produces. The leading tab pushes the equation
    /// to the center stop; the trailing tab pushes the number to the right stop.
    /// </summary>
    private static Paragraph BuildNumberedEquationParagraph(M.OfficeMath oMath, string number)
    {
        // Tab stops: center at half the typical text width, right at the full width
        // (twips; 9026 ≈ 6.27in default text column). These match Word's defaults so
        // the layout reads correctly on open without the caller setting margins.
        const int CenterTwips = 4513;
        const int RightTwips = 9026;
        var pPr = new ParagraphProperties(new Tabs(
            new TabStop { Val = TabStopValues.Center, Position = CenterTwips },
            new TabStop { Val = TabStopValues.Right, Position = RightTwips }));

        return new Paragraph(
            pPr,
            new Run(new TabChar()),
            oMath,
            new Run(new TabChar(), NewText(number)));
    }

    // ------------------------------------------------------- equation numbering

    /// <summary>The vendor attribute holding a numbered equation's label (read-back + addressing).</summary>
    private const string EquationNumberAttribute = "num";

    /// <summary>Stores the numbered equation's label on its <c>m:oMath</c> (alongside the LaTeX).</summary>
    private static void StoreEquationNumber(M.OfficeMath oMath, string number) =>
        oMath.SetAttribute(new OpenXmlAttribute(EquationPrefix, EquationNumberAttribute, EquationNamespace, number));

    /// <summary>The stored equation number, or null for an unnumbered equation.</summary>
    private static string? ReadStoredEquationNumber(M.OfficeMath oMath) =>
        oMath.GetAttributes().FirstOrDefault(a =>
            a.LocalName == EquationNumberAttribute && a.NamespaceUri == EquationNamespace).Value;

    /// <summary>
    /// The next auto sequence number: one past the highest numeric label already
    /// stored on a numbered equation in the document body. Honors the documented
    /// "(1.1)"/"7"/"(3)" label shapes by reading the trailing integer; an
    /// all-non-numeric run of prior labels still advances from the count.
    /// </summary>
    private static string NextEquationNumber(WordprocessingDocument doc)
    {
        var max = 0;
        var seen = 0;
        if (doc.MainDocumentPart?.Document?.Body is { } body)
        {
            foreach (var existing in body.Descendants<M.OfficeMath>())
            {
                if (ReadStoredEquationNumber(existing) is not { } label)
                {
                    continue;
                }

                seen++;
                if (TrailingInteger(label) is { } value && value > max)
                {
                    max = value;
                }
            }
        }

        var next = Math.Max(max, seen) + 1;
        return $"({next})";
    }

    /// <summary>The trailing integer in a label ("(1.4)" → 4, "7" → 7), or null when there is none.</summary>
    private static int? TrailingInteger(string label)
    {
        var end = label.Length;
        while (end > 0 && !char.IsAsciiDigit(label[end - 1]))
        {
            end--;
        }

        var start = end;
        while (start > 0 && char.IsAsciiDigit(label[start - 1]))
        {
            start--;
        }

        return start < end && int.TryParse(label.AsSpan(start, end - start), out var value) ? value : null;
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

    /// <summary>The numbered display equation in a paragraph (an inline <c>m:oMath</c> carrying a number), or null.</summary>
    internal static M.OfficeMath? NumberedEquationInParagraph(Paragraph paragraph) =>
        paragraph.ChildElements.OfType<M.OfficeMath>().FirstOrDefault(m => ReadStoredEquationNumber(m) is not null);

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

    /// <summary>get /…/omath[j]: the stored LaTeX source, whether it is a display equation, and its number when numbered.</summary>
    private static Dictionary<string, object?> GetEquationProperties(WordprocessingDocument doc, DocPath path)
    {
        var (oMath, _, canonical) = ResolveEquation(doc, path);
        var shape = new Dictionary<string, object?>
        {
            ["path"] = canonical,
            ["latex"] = ReadStoredLatex(oMath),
            // A numbered equation lives in a tab-aligned body paragraph, not an
            // m:oMathPara, but it is still a display equation for read-back purposes.
            ["display"] = oMath.Parent is M.Paragraph || ReadStoredEquationNumber(oMath) is not null,
        };

        if (ReadStoredEquationNumber(oMath) is { } number)
        {
            shape["number"] = number;
        }

        return shape;
    }

    /// <summary>
    /// get /equation[@num=…]: resolves a numbered display equation by its label and
    /// returns the same equation shape as the omath path. The selector value matches
    /// the stored label either verbatim ("(1.1)") or by its numeric core ("1.1"), so
    /// agents can address it without re-quoting the surrounding parentheses.
    /// </summary>
    private static Dictionary<string, object?> GetNumberedEquationProperties(WordprocessingDocument doc, DocPath path)
    {
        var segment = path.Segments[0];
        if (path.Segments.Count != 1 || segment.IdAttribute != "num" || segment.Id is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"'{path.ToCanonicalString()}' is not a numbered-equation path.",
                "Address a numbered equation as /equation[@num=(1.1)] or by its number /equation[@num=1.1].");
        }

        var wanted = segment.Id;
        var matches = new List<(M.OfficeMath OMath, Paragraph Paragraph, string Label)>();
        if (doc.MainDocumentPart?.Document?.Body is { } body)
        {
            foreach (var paragraph in body.Descendants<Paragraph>())
            {
                foreach (var oMath in paragraph.ChildElements.OfType<M.OfficeMath>())
                {
                    if (ReadStoredEquationNumber(oMath) is { } label && EquationNumberMatches(label, wanted))
                    {
                        matches.Add((oMath, paragraph, label));
                    }
                }
            }
        }

        if (matches.Count == 0)
        {
            var available = NumberedEquationLabels(doc);
            throw new AiofficeException(
                ErrorCodes.InvalidPath,
                $"No numbered equation has number '{wanted}'.",
                available.Count > 0
                    ? "Use one of the numbers that exist, or read --view text to find equations."
                    : "This document has no numbered equations. Add one with display:true and number:true.",
                candidates: [.. available.Select(n => $"/equation[@num={n}]")]);
        }

        var (matchedMath, _, _) = matches[0];
        var omathPath = EquationOmathPath(doc, matchedMath);
        return new Dictionary<string, object?>
        {
            ["path"] = omathPath,
            ["latex"] = ReadStoredLatex(matchedMath),
            ["display"] = true,
            ["number"] = ReadStoredEquationNumber(matchedMath),
        };
    }

    /// <summary>True when a stored label answers to the requested selector value, verbatim or by numeric core.</summary>
    private static bool EquationNumberMatches(string label, string wanted)
    {
        if (string.Equals(label, wanted, StringComparison.Ordinal))
        {
            return true;
        }

        // Compare the digits-and-dots cores so "(1.1)" answers to "1.1" and "(7)" to "7".
        return string.Equals(NumericCore(label), NumericCore(wanted), StringComparison.Ordinal)
            && NumericCore(label).Length > 0;
    }

    /// <summary>The leading/trailing-bracket-stripped digits-and-dots core of a label ("(1.1)" → "1.1").</summary>
    private static string NumericCore(string label)
    {
        var core = new string([.. label.Where(c => char.IsAsciiDigit(c) || c == '.')]);
        return core.Trim('.');
    }

    /// <summary>Every numbered equation's label, in document order, for invalid-path candidates.</summary>
    private static List<string> NumberedEquationLabels(WordprocessingDocument doc)
    {
        var labels = new List<string>();
        if (doc.MainDocumentPart?.Document?.Body is { } body)
        {
            foreach (var oMath in body.Descendants<M.OfficeMath>())
            {
                if (ReadStoredEquationNumber(oMath) is { } label)
                {
                    labels.Add(label);
                }
            }
        }

        return labels;
    }

    /// <summary>The canonical <c>/…/p[i]/omath[j]</c> path for an equation element (matches the get/remove surface).</summary>
    private static string EquationOmathPath(WordprocessingDocument doc, M.OfficeMath oMath)
    {
        var paragraph = oMath.Ancestors<Paragraph>().FirstOrDefault();
        if (paragraph is null)
        {
            return "/body";
        }

        var node = WordAddress.EnumerateAll(doc).FirstOrDefault(n => ReferenceEquals(n.Element, paragraph));
        var paragraphPath = node?.CanonicalPath ?? "/body";
        var index = ParagraphEquations(paragraph).FindIndex(m => ReferenceEquals(m, oMath)) + 1;
        return $"{paragraphPath}/omath[{Math.Max(index, 1)}]";
    }

    /// <summary>remove /…/omath[j]: drops an inline equation, or the whole display paragraph it owns.</summary>
    private static object ApplyRemoveEquation(WordprocessingDocument doc, EditOp op)
    {
        var path = DocPath.Parse(op.Path);
        var (oMath, paragraph, canonical) = ResolveEquation(doc, path);

        var numbered = ReadStoredEquationNumber(oMath) is not null;
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
        else if (numbered && ReferenceEquals(oMath.Parent, paragraph))
        {
            // Numbered equation: the equation owns its tab-aligned paragraph (leading
            // tab + m:oMath + trailing tab + number). Drop the whole paragraph unless
            // it is the body's last block, in which case clear it to an empty one.
            if (paragraph.Parent is Body { ChildElements.Count: > 1 })
            {
                paragraph.Remove();
            }
            else
            {
                foreach (var child in paragraph.ChildElements.Where(c => c is not ParagraphProperties).ToList())
                {
                    child.Remove();
                }
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
