using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace AIOffice.Pptx;

/// <summary>
/// Reads and writes the conventional OOXML document core properties at
/// <c>docProps/core.xml</c> (the SDK's <see cref="CoreFilePropertiesPart"/>,
/// content-type <c>application/vnd.openxmlformats-package.core-properties+xml</c>,
/// related from the package root by the core-properties relationship).
///
/// The OpenXml SDK does not strongly-type the core properties, so this is a
/// hand-rolled reader/writer over the <c>cp:coreProperties</c> element. The map is:
/// <list type="bullet">
/// <item><c>dc:title</c> ↔ Title</item>
/// <item><c>dc:creator</c> ↔ Author</item>
/// <item><c>dc:subject</c> ↔ Subject</item>
/// <item><c>cp:keywords</c> ↔ Keywords</item>
/// <item><c>cp:category</c> ↔ Category</item>
/// <item><c>dc:description</c> ↔ Comments</item>
/// <item><c>cp:lastModifiedBy</c> ↔ LastModifiedBy</item>
/// <item><c>cp:revision</c> ↔ Revision</item>
/// <item><c>dcterms:created</c> ↔ Created (W3CDTF)</item>
/// <item><c>dcterms:modified</c> ↔ Modified (W3CDTF)</item>
/// </list>
///
/// We deliberately do NOT store via <see cref="System.IO.Packaging"/>
/// <c>PackageProperties</c>: that machinery writes a non-standard
/// <c>package/services/metadata/core-properties/{GUID}.psmdcp</c> part that many
/// tools (and a plain <c>unzip</c>) cannot find. Writing the conventional
/// <c>docProps/core.xml</c> part makes the properties visible to Office and to any
/// OOXML consumer. On read we migrate-on-read: a file that only carries the legacy
/// <c>PackageProperties</c> still resolves through the fallback in
/// <see cref="PptxProperties"/>, and the next write standardises it.
/// </summary>
internal static class PptxCoreProps
{
    private static readonly XNamespace Cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace Dcterms = "http://purl.org/dc/terms/";
    private static readonly XNamespace Dcmitype = "http://purl.org/dc/dcmitype/";
    private static readonly XNamespace Xsi = "http://www.w3.org/2001/XMLSchema-instance";

    /// <summary>The mutable view of the ten core properties this surface exposes.</summary>
    internal sealed class CoreModel
    {
        public string? Title { get; set; }

        public string? Subject { get; set; }

        public string? Author { get; set; }

        public string? Keywords { get; set; }

        public string? Category { get; set; }

        public string? Comments { get; set; }

        public string? LastModifiedBy { get; set; }

        public string? Revision { get; set; }

        public DateTime? Created { get; set; }

        public DateTime? Modified { get; set; }
    }

    // ------------------------------------------------------------------- read

    /// <summary>
    /// Reads <c>docProps/core.xml</c> into a model. Returns <c>null</c> when the
    /// part is absent (the caller then falls back to the legacy
    /// <c>PackageProperties</c> so old files still read).
    /// </summary>
    public static CoreModel? Read(PresentationDocument doc)
    {
        var part = doc.CoreFilePropertiesPart;
        if (part is null)
        {
            return null;
        }

        var root = LoadRoot(part);
        if (root is null)
        {
            return new CoreModel();
        }

        return new CoreModel
        {
            Title = Text(root, Dc + "title"),
            Subject = Text(root, Dc + "subject"),
            Author = Text(root, Dc + "creator"),
            Keywords = Text(root, Cp + "keywords"),
            Category = Text(root, Cp + "category"),
            Comments = Text(root, Dc + "description"),
            LastModifiedBy = Text(root, Cp + "lastModifiedBy"),
            Revision = Text(root, Cp + "revision"),
            Created = Date(root, Dcterms + "created"),
            Modified = Date(root, Dcterms + "modified"),
        };
    }

    // ------------------------------------------------------------------- write

    /// <summary>
    /// Writes the model back to <c>docProps/core.xml</c>, creating the
    /// <see cref="CoreFilePropertiesPart"/> on first use. Only non-null fields
    /// are emitted; an empty string clears its element. Always writes to the
    /// standard part — never to <c>PackageProperties</c>.
    /// </summary>
    public static void Write(PresentationDocument doc, CoreModel model)
    {
        var part = StandardCorePart(doc);
        var root = new XElement(
            Cp + "coreProperties",
            new XAttribute(XNamespace.Xmlns + "cp", Cp.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dc", Dc.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dcterms", Dcterms.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "dcmitype", Dcmitype.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "xsi", Xsi.NamespaceName));

        // The element order matches what Office emits; the schema is order-free,
        // but a stable order keeps writes deterministic across runs.
        AddText(root, Dc + "title", model.Title);
        AddText(root, Dc + "subject", model.Subject);
        AddText(root, Dc + "creator", model.Author);
        AddText(root, Cp + "keywords", model.Keywords);
        AddText(root, Cp + "category", model.Category);
        AddText(root, Dc + "description", model.Comments);
        AddText(root, Cp + "lastModifiedBy", model.LastModifiedBy);
        AddText(root, Cp + "revision", model.Revision);
        AddDate(root, Dcterms + "created", model.Created);
        AddDate(root, Dcterms + "modified", model.Modified);

        var document = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        document.Save(writer);
    }

    /// <summary>
    /// The core-properties part to write into, guaranteed to live at the
    /// conventional <c>docProps/core.xml</c> URI. When the only existing part is
    /// the legacy non-standard <c>.psmdcp</c> part (written by
    /// <c>System.IO.Packaging</c> on older builds), it is removed and a standard
    /// part is added so the write migrates the file onto conventional storage.
    /// </summary>
    private static CoreFilePropertiesPart StandardCorePart(PresentationDocument doc)
    {
        var existing = doc.CoreFilePropertiesPart;
        if (existing is not null && IsStandardUri(existing.Uri))
        {
            return existing;
        }

        if (existing is not null)
        {
            // Legacy .psmdcp part: drop it so the standard docProps/core.xml part
            // becomes the single source of truth (no duplicate metadata parts).
            doc.DeletePart(existing);
        }

        return doc.AddCoreFilePropertiesPart();
    }

    private static bool IsStandardUri(Uri uri) =>
        uri.OriginalString.EndsWith("/docProps/core.xml", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(uri.OriginalString, "/docProps/core.xml", StringComparison.OrdinalIgnoreCase);

    // ------------------------------------------------------------- xml helpers

    private static XElement? LoadRoot(CoreFilePropertiesPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        if (stream.Length == 0)
        {
            return null;
        }

        try
        {
            return XDocument.Load(stream).Root;
        }
        catch (System.Xml.XmlException)
        {
            return null;
        }
    }

    private static string? Text(XElement root, XName name)
    {
        var value = root.Element(name)?.Value;
        return string.IsNullOrEmpty(value) ? null : value;
    }

    private static DateTime? Date(XElement root, XName name)
    {
        var value = root.Element(name)?.Value;
        if (string.IsNullOrEmpty(value))
        {
            return null;
        }

        return DateTime.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
            out var parsed)
            ? parsed
            : null;
    }

    /// <summary>Appends a text element only when the value is non-null (empty clears, but is never written).</summary>
    private static void AddText(XElement root, XName name, string? value)
    {
        if (!string.IsNullOrEmpty(value))
        {
            root.Add(new XElement(name, value));
        }
    }

    /// <summary>Appends a dcterms date as an <c>xsi:type="dcterms:W3CDTF"</c> element in UTC.</summary>
    private static void AddDate(XElement root, XName name, DateTime? value)
    {
        if (value is not { } date)
        {
            return;
        }

        root.Add(new XElement(
            name,
            new XAttribute(Xsi + "type", "dcterms:W3CDTF"),
            date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)));
    }
}
