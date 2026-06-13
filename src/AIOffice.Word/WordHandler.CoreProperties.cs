using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace AIOffice.Word;

/// <summary>
/// Standard OOXML core document properties live in the package's
/// <see cref="CoreFilePropertiesPart"/> at <c>docProps/core.xml</c> (content type
/// <c>application/vnd.openxmlformats-package.core-properties+xml</c>, related from
/// the package root by the core-properties relationship). The OpenXml SDK does not
/// strongly-type that part, so this helper reads and writes its
/// <c>cp:coreProperties</c> XML directly (the <c>dc</c> / <c>cp</c> / <c>dcterms</c>
/// namespaces).
///
/// We deliberately do NOT route storage through <see cref="System.IO.Packaging"/>'s
/// <c>PackageProperties</c>: when no core part exists yet its setter creates a
/// non-conventional <c>package/services/metadata/core-properties/GUID.psmdcp</c>
/// part instead of <c>docProps/core.xml</c>, which Office and most tools cannot
/// see. Reads fall back to the legacy <c>PackageProperties</c> façade (migrate-on-
/// read) so older files still surface their title; writes always land in the
/// standard <c>docProps/core.xml</c> part.
/// </summary>
public sealed partial class WordHandler
{
    private static readonly XNamespace CpNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DcNs = "http://purl.org/dc/elements/1.1/";
    private static readonly XNamespace DcTermsNs = "http://purl.org/dc/terms/";
    private static readonly XNamespace XsiNs = "http://www.w3.org/2001/XMLSchema-instance";

    /// <summary>One core-properties element, identified by its namespace + local name.</summary>
    private enum CoreField
    {
        Title,
        Subject,
        Creator,
        Keywords,
        Category,
        Description,
        LastModifiedBy,
        Revision,
        Created,
        Modified,
    }

    /// <summary>
    /// A snapshot of the standard core properties, read from <c>docProps/core.xml</c>
    /// when present, else migrated from the legacy <c>PackageProperties</c> façade.
    /// </summary>
    private sealed class CoreProps
    {
        public string? Title { get; set; }

        public string? Subject { get; set; }

        public string? Creator { get; set; }

        public string? Keywords { get; set; }

        public string? Category { get; set; }

        public string? Description { get; set; }

        public string? LastModifiedBy { get; set; }

        public string? Revision { get; set; }

        public DateTime? Created { get; set; }

        public DateTime? Modified { get; set; }
    }

    /// <summary>The (namespace, local-name) the OOXML schema gives each core field.</summary>
    private static (XNamespace Ns, string Local) FieldName(CoreField field) => field switch
    {
        CoreField.Title => (DcNs, "title"),
        CoreField.Subject => (DcNs, "subject"),
        CoreField.Creator => (DcNs, "creator"),
        CoreField.Keywords => (CpNs, "keywords"),
        CoreField.Category => (CpNs, "category"),
        CoreField.Description => (DcNs, "description"),
        CoreField.LastModifiedBy => (CpNs, "lastModifiedBy"),
        CoreField.Revision => (CpNs, "revision"),
        CoreField.Created => (DcTermsNs, "created"),
        CoreField.Modified => (DcTermsNs, "modified"),
        _ => throw new ArgumentOutOfRangeException(nameof(field)),
    };

    // --------------------------------------------------------------- migration

    /// <summary>
    /// The legacy core-properties part lives at a non-standard
    /// <c>package/services/metadata/core-properties/GUID.psmdcp</c> URI; the
    /// conventional one is <c>docProps/core.xml</c>.
    /// </summary>
    private static bool IsLegacyCoreUri(Uri uri) =>
        uri.ToString().EndsWith(".psmdcp", StringComparison.OrdinalIgnoreCase);

    /// <summary>True when an edit-op path targets the virtual <c>/properties</c> node.</summary>
    private static bool IsPropertiesPath(string? path)
    {
        var trimmed = path?.Trim().TrimStart('/');
        return trimmed is not null &&
            (trimmed.Equals("properties", StringComparison.OrdinalIgnoreCase) ||
             trimmed.StartsWith("properties/", StringComparison.OrdinalIgnoreCase) ||
             trimmed.StartsWith("properties[", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// One-shot upgrade for files whose core properties only live in the legacy
    /// <c>.psmdcp</c> part: move them to the conventional <c>docProps/core.xml</c>
    /// while preserving every value. This runs as TWO package sessions (delete,
    /// then recreate) because the SDK only assigns the canonical <c>core.xml</c>
    /// name when the add happens in a session distinct from the one that deleted
    /// the old part. A no-op (returns the same bytes) when there is no legacy part,
    /// so the common, already-standard file pays nothing.
    /// </summary>
    private static byte[] MigrateLegacyCorePropertiesBytes(byte[] bytes, string file)
    {
        CoreProps? legacyValues = null;

        // Session A: detect + capture + delete the legacy psmdcp part.
        using var sessionA = new MemoryStream();
        sessionA.Write(bytes);
        sessionA.Position = 0;
        using (var doc = OpenPackage(sessionA, file, editable: true))
        {
            if (doc.CoreFilePropertiesPart is { Uri: { } uri } legacy && IsLegacyCoreUri(uri))
            {
                legacyValues = ReadCorePropsFromPart(legacy);
                doc.DeletePart(legacy);
            }
        }

        if (legacyValues is null)
        {
            return bytes; // already standard (or no core part at all)
        }

        var afterDelete = sessionA.ToArray();

        // Session B: add the standard part (now lands at docProps/core.xml) and
        // write the migrated values into it.
        using var sessionB = new MemoryStream();
        sessionB.Write(afterDelete);
        sessionB.Position = 0;
        using (var doc = OpenPackage(sessionB, file, editable: true))
        {
            var part = doc.AddCoreFilePropertiesPart();
            var root = NewCoreRoot();
            WriteAllCoreFields(root, legacyValues);
            SaveCoreRoot(part, root);
        }

        return sessionB.ToArray();
    }

    /// <summary>Reads a snapshot of every core field straight off a core-properties part's XML.</summary>
    private static CoreProps ReadCorePropsFromPart(CoreFilePropertiesPart part)
    {
        XElement root;
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            root = XDocument.Load(stream).Root ?? new XElement(CpNs + "coreProperties");
        }
        catch (XmlException)
        {
            root = new XElement(CpNs + "coreProperties");
        }

        return new CoreProps
        {
            Title = ReadField(root, CoreField.Title),
            Subject = ReadField(root, CoreField.Subject),
            Creator = ReadField(root, CoreField.Creator),
            Keywords = ReadField(root, CoreField.Keywords),
            Category = ReadField(root, CoreField.Category),
            Description = ReadField(root, CoreField.Description),
            LastModifiedBy = ReadField(root, CoreField.LastModifiedBy),
            Revision = ReadField(root, CoreField.Revision),
            Created = ReadDateField(root, CoreField.Created),
            Modified = ReadDateField(root, CoreField.Modified),
        };
    }

    /// <summary>Writes every non-empty field of a snapshot into a core-properties root.</summary>
    private static void WriteAllCoreFields(XElement root, CoreProps values)
    {
        ApplyCoreField(root, CoreField.Title, values.Title);
        ApplyCoreField(root, CoreField.Subject, values.Subject);
        ApplyCoreField(root, CoreField.Creator, values.Creator);
        ApplyCoreField(root, CoreField.Keywords, values.Keywords);
        ApplyCoreField(root, CoreField.Category, values.Category);
        ApplyCoreField(root, CoreField.Description, values.Description);
        ApplyCoreField(root, CoreField.LastModifiedBy, values.LastModifiedBy);
        ApplyCoreField(root, CoreField.Revision, values.Revision);
        if (values.Created is { } created)
        {
            ApplyCoreField(root, CoreField.Created, FormatCoreDate(created));
        }

        if (values.Modified is { } modified)
        {
            ApplyCoreField(root, CoreField.Modified, FormatCoreDate(modified));
        }
    }

    // -------------------------------------------------------------------- read

    /// <summary>
    /// Reads the standard core properties from <c>docProps/core.xml</c> (or the
    /// legacy <c>.psmdcp</c> part, whose XML the SDK still surfaces as the
    /// CoreFilePropertiesPart). When no core part exists at all, migrates the
    /// values off the <see cref="System.IO.Packaging.PackageProperties"/> façade so
    /// older files still surface their title and friends.
    /// </summary>
    private static CoreProps ReadCoreProps(WordprocessingDocument doc) =>
        doc.CoreFilePropertiesPart is { } part
            ? ReadCorePropsFromPart(part)
            : ReadLegacyCoreProps(doc);

    /// <summary>Migrate-on-read: project the legacy PackageProperties façade into our shape.</summary>
    private static CoreProps ReadLegacyCoreProps(WordprocessingDocument doc)
    {
        var legacy = doc.PackageProperties;
        return new CoreProps
        {
            Title = NullIfEmpty(legacy.Title),
            Subject = NullIfEmpty(legacy.Subject),
            Creator = NullIfEmpty(legacy.Creator),
            Keywords = NullIfEmpty(legacy.Keywords),
            Category = NullIfEmpty(legacy.Category),
            Description = NullIfEmpty(legacy.Description),
            LastModifiedBy = NullIfEmpty(legacy.LastModifiedBy),
            Revision = NullIfEmpty(legacy.Revision),
            Created = legacy.Created,
            Modified = legacy.Modified,
        };
    }

    /// <summary>The standardized core Title (docProps/core.xml), with the legacy fallback.</summary>
    private static string? ReadCoreTitle(WordprocessingDocument doc) => ReadCoreProps(doc).Title;

    private static string? ReadField(XElement root, CoreField field)
    {
        var (ns, local) = FieldName(field);
        var value = root.Element(ns + local)?.Value;
        return value is { Length: > 0 } ? value : null;
    }

    private static DateTime? ReadDateField(XElement root, CoreField field)
    {
        var raw = ReadField(root, field);
        if (raw is null)
        {
            return null;
        }

        return DateTime.TryParse(
            raw,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
            out var parsed)
            ? parsed
            : null;
    }

    // ------------------------------------------------------------------- write

    /// <summary>
    /// Sets one core field on the standard <c>docProps/core.xml</c> part, creating
    /// the part (and the W3CDTF-typed dcterms date attributes) as needed. An empty
    /// string clears the element. This always writes the conventional
    /// CoreFilePropertiesPart, never the legacy psmdcp part.
    /// </summary>
    private static void WriteCoreField(WordprocessingDocument doc, CoreField field, string? value)
    {
        var root = LoadOrCreateCoreRoot(doc, out var part);
        ApplyCoreField(root, field, value);
        SaveCoreRoot(part, root);
    }

    private static void ApplyCoreField(XElement root, CoreField field, string? value)
    {
        var (ns, local) = FieldName(field);
        var element = root.Element(ns + local);

        if (string.IsNullOrEmpty(value))
        {
            element?.Remove();
            return;
        }

        if (element is null)
        {
            element = new XElement(ns + local);
            root.Add(element);
        }

        element.Value = value;

        // dcterms dates carry the W3CDTF type attribute Office expects.
        if (field is CoreField.Created or CoreField.Modified)
        {
            element.SetAttributeValue(XsiNs + "type", "dcterms:W3CDTF");
        }
    }

    /// <summary>Loads the existing core-properties XML, or seeds an empty, correctly-namespaced root.</summary>
    private static XElement LoadOrCreateCoreRoot(WordprocessingDocument doc, out CoreFilePropertiesPart part)
    {
        // If a standard part already exists, parse it in place. We read the legacy
        // façade FIRST when no standard part exists, because once we add the empty
        // CoreFilePropertiesPart the PackageProperties façade would try (and fail)
        // to parse that empty part.
        if (doc.CoreFilePropertiesPart is { } existingPart)
        {
            part = existingPart;
            try
            {
                using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
                if (stream.Length > 0)
                {
                    var existing = XDocument.Load(stream).Root;
                    if (existing is not null)
                    {
                        EnsureCoreNamespaces(existing);
                        return existing;
                    }
                }
            }
            catch (XmlException)
            {
                // a malformed / empty part falls through to a fresh root (no legacy to migrate)
            }

            return NewCoreRoot();
        }

        // No standard part: snapshot the legacy façade before creating the part,
        // then seed the fresh root so old psmdcp-only files keep their values.
        var legacy = ReadLegacyCoreProps(doc);
        part = doc.AddCoreFilePropertiesPart();

        // Seed the fresh root from the legacy façade so old psmdcp-only files keep
        // their values when the first standard write happens.
        var root = NewCoreRoot();
        WriteAllCoreFields(root, legacy);
        return root;
    }

    private static XElement NewCoreRoot() => new(
        CpNs + "coreProperties",
        new XAttribute(XNamespace.Xmlns + "cp", CpNs.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "dc", DcNs.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "dcterms", DcTermsNs.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "xsi", XsiNs.NamespaceName));

    /// <summary>Guarantees the cp/dc/dcterms/xsi namespace declarations are present on the root.</summary>
    private static void EnsureCoreNamespaces(XElement root)
    {
        EnsureNamespace(root, "cp", CpNs);
        EnsureNamespace(root, "dc", DcNs);
        EnsureNamespace(root, "dcterms", DcTermsNs);
        EnsureNamespace(root, "xsi", XsiNs);
    }

    private static void EnsureNamespace(XElement root, string prefix, XNamespace ns)
    {
        if (root.GetPrefixOfNamespace(ns) is null)
        {
            root.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName);
        }
    }

    private static void SaveCoreRoot(CoreFilePropertiesPart part, XElement root)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            OmitXmlDeclaration = false,
            Indent = false,
        });
        new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root).Save(writer);
    }

    private static string FormatCoreDate(DateTime value) =>
        value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
}
