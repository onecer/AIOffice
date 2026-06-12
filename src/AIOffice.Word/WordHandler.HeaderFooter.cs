using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    // The M2-era unsupported_feature refusal for first-page/even-odd types is
    // gone: M5 implements them for real (w:titlePg / w:evenAndOddHeaders), so
    // refusing would now be dishonest about a capability the handler has.

    /// <summary>props.type spellings accepted per variant (path variants are exact).</summary>
    private static readonly Dictionary<string, string> HeaderFooterTypeSpellings = new(StringComparer.OrdinalIgnoreCase)
    {
        ["default"] = "default",
        ["first"] = "firstPage",
        ["firstpage"] = "firstPage",
        ["first-page"] = "firstPage",
        ["even"] = "even",
    };

    /// <summary>
    /// <c>{"op":"add","path":"/header[1]|/header[firstPage]|/header[even]","type":"header","props":{…}}</c>:
    /// creates the header (or footer) of one variant with one paragraph. The
    /// default variant applies to all pages; firstPage wires w:titlePg on the
    /// body sectPr; even wires w:evenAndOddHeaders in settings (the default
    /// variant then serves odd pages). Adding a variant that already exists is
    /// <c>invalid_args</c> (edit its paragraphs instead).
    /// </summary>
    private static object ApplyAddHeaderFooter(WordprocessingDocument doc, string file, EditOp op)
    {
        var kind = op.Type!; // "header" or "footer", routed here by ApplyOp

        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author"); // batch attribution metadata, not a paragraph prop
        var variant = RequestedHeaderFooterVariant(op, props, kind);

        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            $"The document has no main part: {file}",
            "The package is missing its main document part. Re-export the file from Word.");

        if (WordAddress.FindReferencedPart(doc, kind, variant) is not null)
        {
            var existingPath = variant == "default" ? $"/{kind}[1]" : $"/{kind}[{variant}]";
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"{existingPath} already exists; add --type {kind} only creates a variant the document lacks.",
                $"Edit its paragraphs instead, e.g. {{\"op\":\"set\",\"path\":\"{existingPath}/p[1]\",\"props\":{{\"text\":\"…\"}}}}, " +
                $"or insert one with {{\"op\":\"add\",\"path\":\"{existingPath}\",\"type\":\"p\"}}.");
        }

        var body = GetBody(doc, file);
        var paragraph = BuildParagraph(doc, props); // validates props before the part lands

        string relId;
        int index;
        if (kind == "header")
        {
            var part = main.AddNewPart<HeaderPart>();
            part.Header = new Header(paragraph);
            relId = main.GetIdOfPart(part);
            index = main.HeaderParts.ToList().IndexOf(part) + 1;
        }
        else
        {
            var part = main.AddNewPart<FooterPart>();
            part.Footer = new Footer(paragraph);
            relId = main.GetIdOfPart(part);
            index = main.FooterParts.ToList().IndexOf(part) + 1;
        }

        // r:id needs the relationships namespace in scope at the root (like Word
        // writes it); declaring it inline on the reference element would break the
        // round-trip law, because a reopen+save hoists it to the root.
        const string RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var document = main.Document!;
        if (document.LookupNamespace("r") is null)
        {
            document.AddNamespaceDeclaration("r", RelationshipsNs);
        }

        // The reference lives on the body-level sectPr (created last-in-body when
        // absent); hdr/ftr references must precede every other sectPr child.
        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectPr is null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }

        var referenceType = WordAddress.VariantReferenceType(variant);
        OpenXmlElement reference = kind == "header"
            ? new HeaderReference { Type = referenceType, Id = relId }
            : new FooterReference { Type = referenceType, Id = relId };
        sectPr.InsertAt(reference, 0);

        // The variant only shows once its section/document switch is on.
        if (variant == "firstPage" && sectPr.GetFirstChild<TitlePage>() is null)
        {
            InsertSectionChild(sectPr, new TitlePage());
        }

        if (variant == "even")
        {
            EnsureEvenAndOddHeaders(doc);
        }

        var canonical = $"/{kind}[{index}]";
        return new { op = "add", type = kind, variant, path = canonical, paragraph = $"{canonical}/p[1]" };
    }

    /// <summary>
    /// The variant requested by an add --type header|footer op: the path's named
    /// variant (/header[firstPage]), props.type (legacy spelling tolerated), or
    /// default. The path must be the bare root: /header, /header[1] or a named
    /// variant — numeric indices &gt; 1 are part-order artifacts, not variants.
    /// </summary>
    private static string RequestedHeaderFooterVariant(EditOp op, JsonObject props, string kind)
    {
        var path = DocPath.Parse(op.Path);
        var root = path.Segments[0];
        var rootOk = path.Segments.Count == 1
            && root.Kind == PathSegmentKind.Element
            && string.Equals(root.Name, kind, StringComparison.Ordinal)
            && root.Id is null
            && (root.Variant is not null || (root.Index ?? 1) == 1);

        if (!rootOk)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add --type {kind} targets /{kind}[1], /{kind}[firstPage] or /{kind}[even], not '{op.Path}'.",
                $"Use {{\"op\":\"add\",\"path\":\"/{kind}[1]\",\"type\":\"{kind}\",\"props\":{{\"text\":\"…\"}}}} " +
                "(or a named variant path for first-page/even-page content).",
                candidates: [$"/{kind}[1]", $"/{kind}[firstPage]", $"/{kind}[even]"]);
        }

        string? fromProps = null;
        if (props.TryGetPropertyValue("type", out var typeNode))
        {
            props.Remove("type");
            var value = NodeToString(typeNode);
            if (!HeaderFooterTypeSpellings.TryGetValue(value, out fromProps))
            {
                var isOdd = value.ToLowerInvariant() is "odd" or "evenodd" or "even-odd";
                throw new AiofficeException(
                    ErrorCodes.InvalidArgs,
                    $"Unknown {kind} type '{value}'.",
                    isOdd
                        ? $"There is no odd-page {kind} type in OOXML: once the even variant exists, the default {kind} serves the odd pages. " +
                          $"Create /{kind}[even] and put the odd-page content in /{kind}[1]."
                        : $"Use type default, firstPage or even (or address the variant in the path, e.g. /{kind}[firstPage]).",
                    candidates: ["default", "firstPage", "even"]);
            }
        }

        var fromPath = root.Variant;
        if (fromPath is not null && fromProps is not null && fromPath != fromProps)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"The path says /{kind}[{fromPath}] but props.type says '{fromProps}'.",
                "Drop props.type (the path variant wins) or make them agree.");
        }

        return fromPath ?? fromProps ?? "default";
    }

    // ---------------------------------------------------------------- settings

    /// <summary>
    /// CT_Settings child order (the slice around w:evenAndOddHeaders plus the
    /// common trailing elements real documents carry), so the flag lands at its
    /// schema position even in foreign files.
    /// </summary>
    private static readonly Type[] SettingsOrder =
    [
        typeof(WriteProtection), typeof(View), typeof(Zoom), typeof(RemovePersonalInformation),
        typeof(HideSpellingErrors), typeof(HideGrammaticalErrors), typeof(ProofState), typeof(AttachedTemplate),
        typeof(LinkStyles), typeof(StylePaneFormatFilter), typeof(DocumentType), typeof(MailMerge),
        typeof(RevisionView), typeof(DocumentProtection), typeof(AutoFormatOverride),
        typeof(DefaultTabStop), typeof(AutoHyphenation), typeof(ConsecutiveHyphenLimit), typeof(HyphenationZone),
        typeof(DefaultTableStyle), typeof(EvenAndOddHeaders), typeof(BookFoldReversePrinting), typeof(BookFoldPrinting),
        typeof(BookFoldPrintingSheets), typeof(DrawingGridHorizontalSpacing), typeof(DrawingGridVerticalSpacing),
        typeof(DisplayHorizontalDrawingGrid), typeof(DisplayVerticalDrawingGrid), typeof(NoPunctuationKerning),
        typeof(CharacterSpacingControl), typeof(SavePreviewPicture), typeof(UpdateFieldsOnOpen), typeof(Compatibility),
        typeof(Rsids), typeof(DocumentFormat.OpenXml.Math.MathProperties), typeof(ThemeFontLanguages),
        typeof(ColorSchemeMapping), typeof(ShapeDefaults), typeof(DecimalSymbol), typeof(ListSeparator),
    ];

    /// <summary>Turns on w:evenAndOddHeaders in settings (created on demand), idempotently.</summary>
    private static void EnsureEvenAndOddHeaders(WordprocessingDocument doc)
    {
        var settings = EnsureSettingsRoot(doc);
        if (settings.GetFirstChild<EvenAndOddHeaders>() is not null)
        {
            return;
        }

        var flag = new EvenAndOddHeaders();
        var rank = Array.IndexOf(SettingsOrder, typeof(EvenAndOddHeaders));
        var before = settings.ChildElements.FirstOrDefault(existing =>
        {
            var existingRank = Array.IndexOf(SettingsOrder, existing.GetType());
            return existingRank > rank; // unknown (-1) children sort first and never push us back
        });

        if (before is null)
        {
            settings.AppendChild(flag);
        }
        else
        {
            settings.InsertBefore(flag, before);
        }
    }

    private static Settings EnsureSettingsRoot(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            "The document has no main part.",
            "Re-export the file from Word.");

        var part = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
        return part.Settings ??= new Settings();
    }

    // --------------------------------------------------------------------- get

    /// <summary>get on a /header[i] | /footer[i] root: kind, variant and content shape.</summary>
    private static Dictionary<string, object?> HeaderFooterProperties(WordprocessingDocument doc, ResolvedNode node)
    {
        var kind = node.Element is Header ? "header" : "footer";
        var part = OwningPart(doc, node.Element);
        return new Dictionary<string, object?>
        {
            ["kind"] = kind,
            ["variant"] = part is null ? null : WordAddress.HeaderFooterVariantOf(doc, part) ?? "unreferenced",
            ["paragraphs"] = node.Element.ChildElements.OfType<Paragraph>().Count(),
            ["text"] = node.Element.InnerText,
        };
    }
}
