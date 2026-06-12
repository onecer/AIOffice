using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    /// <summary>Spellings of the M2-deferred header/footer types we refuse honestly.</summary>
    private static readonly string[] DeferredHeaderFooterTypes =
        ["first", "firstpage", "first-page", "even", "odd", "evenodd", "even-odd"];

    /// <summary>
    /// <c>{"op":"add","path":"/header[1]","type":"header","props":{…}}</c>: creates
    /// the default-type header (or footer) with one paragraph when the document has
    /// none. Adding when one exists is <c>invalid_args</c> (edit /header[1]/p[i]
    /// instead); first-page/even-odd types are <c>unsupported_feature</c> until M2.
    /// </summary>
    private static object ApplyAddHeaderFooter(WordprocessingDocument doc, string file, EditOp op)
    {
        var kind = op.Type!; // "header" or "footer", routed here by ApplyOp
        RequireHeaderFooterRootPath(op, kind);

        var props = op.Props?.DeepClone().AsObject() ?? [];
        RejectDeferredHeaderFooterType(props, kind);

        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            $"The document has no main part: {file}",
            "The package is missing its main document part. Re-export the file from Word.");

        var exists = kind == "header" ? main.HeaderParts.Any() : main.FooterParts.Any();
        if (exists)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"/{kind}[1] already exists; add --type {kind} only creates one when the document has none.",
                $"Edit its paragraphs instead, e.g. {{\"op\":\"set\",\"path\":\"/{kind}[1]/p[1]\",\"props\":{{\"text\":\"…\"}}}}, " +
                $"or insert one with {{\"op\":\"add\",\"path\":\"/{kind}[1]\",\"type\":\"p\"}}.");
        }

        var body = GetBody(doc, file);
        var paragraph = BuildParagraph(doc, props); // validates props before the part lands

        string relId;
        if (kind == "header")
        {
            var part = main.AddNewPart<HeaderPart>();
            part.Header = new Header(paragraph);
            relId = main.GetIdOfPart(part);
        }
        else
        {
            var part = main.AddNewPart<FooterPart>();
            part.Footer = new Footer(paragraph);
            relId = main.GetIdOfPart(part);
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

        OpenXmlElement reference = kind == "header"
            ? new HeaderReference { Type = HeaderFooterValues.Default, Id = relId }
            : new FooterReference { Type = HeaderFooterValues.Default, Id = relId };
        sectPr.InsertAt(reference, 0);

        return new { op = "add", type = kind, path = $"/{kind}[1]", paragraph = $"/{kind}[1]/p[1]" };
    }

    /// <summary>add --type header|footer must target exactly /header[1] | /footer[1].</summary>
    private static void RequireHeaderFooterRootPath(EditOp op, string kind)
    {
        var path = DocPath.Parse(op.Path);
        var root = path.Segments[0];
        var ok = path.Segments.Count == 1
            && root.Kind == PathSegmentKind.Element
            && string.Equals(root.Name, kind, StringComparison.Ordinal)
            && (root.Index ?? 1) == 1;

        if (!ok)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add --type {kind} targets /{kind}[1], not '{op.Path}'.",
                $"Use {{\"op\":\"add\",\"path\":\"/{kind}[1]\",\"type\":\"{kind}\",\"props\":{{\"text\":\"…\"}}}}; " +
                $"M1 supports one default-type {kind} per document.",
                candidates: [$"/{kind}[1]"]);
        }
    }

    /// <summary>
    /// Honesty gate: props.type first/even-odd is a real Word capability we defer
    /// to M2, so it fails as <c>unsupported_feature</c> naming the default-type
    /// workaround. The key is consumed so the rest of props builds the paragraph.
    /// </summary>
    private static void RejectDeferredHeaderFooterType(JsonObject props, string kind)
    {
        if (props["type"] is not { } typeNode)
        {
            return;
        }

        props.Remove("type");
        var value = NodeToString(typeNode);
        if (value.Equals("default", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        if (DeferredHeaderFooterTypes.Contains(value.ToLowerInvariant()))
        {
            throw new AiofficeException(
                ErrorCodes.UnsupportedFeature,
                $"First-page and even/odd {kind}s are not supported yet (planned for M2); got props.type '{value}'.",
                $"Create the default-type {kind} instead (omit props.type or pass \"default\"); it applies to all pages.",
                candidates: ["default"]);
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Unknown {kind} type '{value}'.",
            $"Omit props.type or pass \"default\". First-page/even-odd {kind}s arrive in M2.",
            candidates: ["default"]);
    }
}
