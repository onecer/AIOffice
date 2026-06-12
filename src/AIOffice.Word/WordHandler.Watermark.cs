using System.Globalization;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;

namespace AIOffice.Word;

public sealed partial class WordHandler
{
    private const string WatermarkShapeIdPrefix = "PowerPlusWaterMarkObject";

    /// <summary>
    /// <c>{"op":"add","path":"/body","type":"watermark","props":{"text":"DRAFT","color":"C0C0C0"?,"diagonal":true?}}</c>:
    /// the classic VML WordArt watermark (text-path shape 136, the one Word's
    /// Watermark gallery inserts), placed in every existing header — creating
    /// the default header first when the document has none. VML in headers is
    /// valid OOXML, so the validator stays clean.
    /// </summary>
    private static object ApplyAddWatermark(WordprocessingDocument doc, string file, EditOp op)
    {
        var path = DocPath.Parse(op.Path);
        var root = path.Segments[0];
        if (path.Segments.Count != 1 || root.Name != "body" || root.Index is not null || root.Id is not null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"add --type watermark targets /body (it lands in the headers), not '{op.Path}'.",
                "Use {\"op\":\"add\",\"path\":\"/body\",\"type\":\"watermark\",\"props\":{\"text\":\"DRAFT\"}}.",
                candidates: ["/body"]);
        }

        var props = op.Props?.DeepClone().AsObject() ?? [];
        props.Remove("author");
        var text = props["text"] is { } textNode ? NodeToString(textNode) : null;
        if (string.IsNullOrEmpty(text))
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "add --type watermark needs props.text.",
                "Example: {\"op\":\"add\",\"path\":\"/body\",\"type\":\"watermark\",\"props\":{\"text\":\"DRAFT\",\"diagonal\":true}}.");
        }

        var color = props["color"] is { } colorNode
            ? WordFormatting.ParseHexColor(NodeToString(colorNode))
            : "C0C0C0";
        var diagonal = props["diagonal"] is not { } diagonalNode
            || WordFormatting.ParseBool("diagonal", NodeToString(diagonalNode));

        if (FindWatermarkShapes(doc).Count > 0)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                "/watermark[1] already exists; a document carries one watermark.",
                "Remove it first ({\"op\":\"remove\",\"path\":\"/watermark[1]\"}) and add the new one in the same batch.");
        }

        var main = doc.MainDocumentPart ?? throw new AiofficeException(
            ErrorCodes.FormatCorrupt,
            $"The document has no main part: {file}",
            "The package is missing its main document part. Re-export the file from Word.");

        if (!main.HeaderParts.Any())
        {
            CreateDefaultHeader(doc, file);
        }

        var headers = 0;
        var shapeIndex = 0;
        foreach (var part in main.HeaderParts)
        {
            var header = part.Header ??= new Header(new Paragraph());

            // A reopen+save hoists the VML namespaces to the w:hdr root; declaring
            // them there from the start keeps the round-trip law byte-identical.
            if (header.LookupNamespace("v") is null)
            {
                header.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            }

            if (header.LookupNamespace("o") is null)
            {
                header.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            }

            var paragraph = header.Elements<Paragraph>().FirstOrDefault();
            if (paragraph is null)
            {
                paragraph = new Paragraph();
                header.AppendChild(paragraph);
            }

            shapeIndex++;
            paragraph.AppendChild(new Run(BuildWatermarkPicture(text, color, diagonal, shapeIndex)));
            headers++;
        }

        return new { op = "add", type = "watermark", path = "/watermark[1]", text, headers };
    }

    /// <summary>remove /watermark[1]: drops the w:pict run from every header; the headers themselves stay.</summary>
    private static object ApplyRemoveWatermark(WordprocessingDocument doc, EditOp op)
    {
        var shapes = ResolveWatermark(doc, DocPath.Parse(op.Path));
        var headers = 0;
        foreach (var shape in shapes)
        {
            DocumentFormat.OpenXml.OpenXmlElement? host =
                shape.Ancestors<Run>().FirstOrDefault()
                ?? (DocumentFormat.OpenXml.OpenXmlElement?)shape.Ancestors<Picture>().FirstOrDefault();
            (host ?? shape).Remove();
            headers++;
        }

        return new { op = "remove", path = "/watermark[1]", type = "watermark", headers };
    }

    // ------------------------------------------------------------------ read

    /// <summary>get /watermark[1] data.</summary>
    private static Dictionary<string, object?> GetWatermarkProperties(WordprocessingDocument doc, DocPath path)
    {
        var shapes = ResolveWatermark(doc, path);
        var shape = shapes[0];
        var style = shape.Style?.Value ?? string.Empty;

        return new Dictionary<string, object?>
        {
            ["text"] = shape.Descendants<V.TextPath>().FirstOrDefault()?.String?.Value,
            ["color"] = shape.FillColor?.Value?.TrimStart('#').ToUpperInvariant(),
            ["diagonal"] = style.Contains("rotation:315", StringComparison.Ordinal),
            ["headers"] = shapes.Count,
        };
    }

    /// <summary>Every watermark shape across all headers (one per header once added).</summary>
    private static List<V.Shape> FindWatermarkShapes(WordprocessingDocument doc) =>
        [.. (doc.MainDocumentPart?.HeaderParts ?? [])
            .Select(part => part.Header)
            .OfType<Header>()
            .SelectMany(header => header.Descendants<V.Shape>())
            .Where(shape => shape.Id?.Value?.StartsWith(WatermarkShapeIdPrefix, StringComparison.Ordinal) == true)];

    /// <summary>Resolves /watermark[1] or throws invalid_path.</summary>
    private static List<V.Shape> ResolveWatermark(WordprocessingDocument doc, DocPath path)
    {
        var segment = path.Segments[0];
        var shapes = FindWatermarkShapes(doc);

        if (path.Segments.Count != 1 || segment.Id is not null || (segment.Index ?? 1) != 1)
        {
            throw WatermarkNotFound($"'{path.ToCanonicalString()}' is not a watermark path; a document has at most /watermark[1].", shapes);
        }

        if (shapes.Count == 0)
        {
            throw WatermarkNotFound("This document has no watermark.", shapes);
        }

        return shapes;
    }

    private static AiofficeException WatermarkNotFound(string message, List<V.Shape> shapes) => new(
        ErrorCodes.InvalidPath,
        message,
        "Add one with {\"op\":\"add\",\"path\":\"/body\",\"type\":\"watermark\",\"props\":{\"text\":\"DRAFT\"}}.",
        candidates: shapes.Count > 0 ? ["/watermark[1]"] : []);

    // ----------------------------------------------------------------- build

    /// <summary>The default-type header part with one empty paragraph, referenced from the body sectPr.</summary>
    private static void CreateDefaultHeader(WordprocessingDocument doc, string file)
    {
        var main = doc.MainDocumentPart!;
        var body = GetBody(doc, file);

        var part = main.AddNewPart<HeaderPart>();
        part.Header = new Header(new Paragraph());
        var relId = main.GetIdOfPart(part);

        // Same round-trip-safe namespace handling as add --type header.
        const string RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var document = main.Document!;
        if (document.LookupNamespace("r") is null)
        {
            document.AddNamespaceDeclaration("r", RelationshipsNs);
        }

        var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectPr is null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }

        sectPr.InsertAt(new HeaderReference { Type = HeaderFooterValues.Default, Id = relId }, 0);
    }

    /// <summary>
    /// The w:pict with Word's classic text-path watermark: shapetype 136 plus a
    /// centered, half-transparent shape. Built from literal VML (namespaces
    /// declared inline) so the markup matches what Word itself writes.
    /// </summary>
    private static Picture BuildWatermarkPicture(string text, string colorHex, bool diagonal, int index)
    {
        const string ns = "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\"";
        var fill = "#" + colorHex.ToLowerInvariant();
        var rotation = diagonal ? ";rotation:315" : string.Empty;
        var spid = (2048 + index).ToString(CultureInfo.InvariantCulture);
        var style =
            "position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt" + rotation +
            ";z-index:-251654144;mso-position-horizontal:center;mso-position-horizontal-relative:margin;" +
            "mso-position-vertical:center;mso-position-vertical-relative:margin";

        var xml =
            $"<v:shapetype {ns} id=\"_x0000_t136\" coordsize=\"21600,21600\" o:spt=\"136\" adj=\"10800\" " +
            "path=\"m@7,l@8,m@5,21600l@6,21600e\">" +
            "<v:formulas>" +
            "<v:f eqn=\"sum #0 0 10800\"/><v:f eqn=\"prod #0 2 1\"/><v:f eqn=\"sum 21600 0 @1\"/>" +
            "<v:f eqn=\"sum 0 0 @2\"/><v:f eqn=\"sum 21600 0 @3\"/><v:f eqn=\"if @0 @3 0\"/>" +
            "<v:f eqn=\"if @0 21600 @1\"/><v:f eqn=\"if @0 0 @2\"/><v:f eqn=\"if @0 @4 21600\"/>" +
            "<v:f eqn=\"mid @5 @6\"/><v:f eqn=\"mid @8 @5\"/><v:f eqn=\"mid @7 @8\"/>" +
            "<v:f eqn=\"mid @6 @7\"/><v:f eqn=\"sum @6 0 @5\"/>" +
            "</v:formulas>" +
            "<v:path textpathok=\"t\" o:connecttype=\"custom\" o:connectlocs=\"@9,0;@10,10800;@11,21600;@12,10800\" " +
            "o:connectangles=\"270,180,90,0\"/>" +
            "<v:textpath on=\"t\" fitshape=\"t\"/>" +
            "</v:shapetype>" +
            $"<v:shape {ns} id=\"{WatermarkShapeIdPrefix}{index}\" o:spid=\"_x0000_s{spid}\" type=\"#_x0000_t136\" " +
            $"style=\"{style}\" o:allowincell=\"f\" fillcolor=\"{fill}\" stroked=\"f\">" +
            "<v:fill opacity=\".5\"/>" +
            $"<v:textpath style=\"font-family:&quot;Calibri&quot;;font-size:1pt\" string=\"{XmlAttribute(text)}\"/>" +
            "</v:shape>";

        return new Picture { InnerXml = xml };
    }

    private static string XmlAttribute(string value) => value
        .Replace("&", "&amp;", StringComparison.Ordinal)
        .Replace("<", "&lt;", StringComparison.Ordinal)
        .Replace(">", "&gt;", StringComparison.Ordinal)
        .Replace("\"", "&quot;", StringComparison.Ordinal);
}
