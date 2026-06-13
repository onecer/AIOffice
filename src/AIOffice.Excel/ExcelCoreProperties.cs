using System.IO.Packaging;

namespace AIOffice.Excel;

/// <summary>
/// 1.0 hardening (M9): keeps the workbook's CORE document properties in the
/// conventional OOXML part <c>docProps/core.xml</c> — the OpenXml SDK's
/// <see cref="DocumentFormat.OpenXml.Packaging.CoreFilePropertiesPart"/>,
/// content-type <c>application/vnd.openxmlformats-package.core-properties+xml</c>,
/// related from the package root by the core-properties relationship.
///
/// <para><b>Why this exists.</b> ClosedXML 0.105 (and the <c>System.IO.Packaging</c>
/// layer it routes <see cref="ClosedXML.Excel.XLWorkbookProperties"/> through)
/// writes the core properties into a NON-STANDARD package part named like
/// <c>package/services/metadata/core-properties/{GUID}.psmdcp</c>. The part is
/// related correctly (so the SDK and ClosedXML both still READ the title back),
/// but it does not live at <c>docProps/core.xml</c>, so unzipping the file shows
/// nothing there and tools that look for the conventional part — including
/// Microsoft Office's property surface and many third-party readers — miss the
/// title/author. This relocates that part to the standard URI, preserving its
/// bytes verbatim.</para>
///
/// <para><b>Bytes preserved verbatim, not re-serialized.</b> The relocate moves
/// ClosedXML's exact serialization to the new URI rather than re-authoring the
/// XML. That keeps the file deterministic under a later plain ClosedXML
/// open/save: ClosedXML reloads the part it already knows how to write and
/// re-emits it byte-identically, so <c>docProps/core.xml</c> stays out of the
/// round-trip law's changed-part set. (Re-authoring the XML by hand would make a
/// plain resave rewrite it — different encoding decl, namespace order, timestamp
/// precision — and break the round-trip law.)</para>
///
/// <para><b>Migrate-on-read is automatic.</b> Reads go through ClosedXML's
/// <see cref="ClosedXML.Excel.XLWorkbook.Properties"/>, which finds the core part
/// by FOLLOWING the package-root core-properties relationship, not by assuming a
/// URI. So an old file that still carries only the legacy <c>.psmdcp</c> part
/// reads back its title/author unchanged, and aioffice always WRITES the standard
/// <c>docProps/core.xml</c> on the next save — no read-path change is needed.</para>
///
/// <para>Idempotent: a no-op when the core part is already at
/// <c>docProps/core.xml</c> (or when there is no core part at all).</para>
/// </summary>
internal static class ExcelCoreProperties
{
    /// <summary>The conventional OOXML part path for core document properties.</summary>
    private const string StandardPartPath = "/docProps/core.xml";

    /// <summary>The package-root relationship type that points at the core-properties part.</summary>
    private const string CorePropertiesRelationshipType =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

    /// <summary>The content type of the core-properties part.</summary>
    private const string CorePropertiesContentType =
        "application/vnd.openxmlformats-package.core-properties+xml";

    private static readonly Uri StandardPartUri = new(StandardPartPath, UriKind.Relative);
    private static readonly Uri PackageRootUri = new("/", UriKind.Relative);

    /// <summary>
    /// Relocates the core-properties part to the standard <c>docProps/core.xml</c>
    /// URI, preserving its bytes verbatim, so unzip and Office see the title where
    /// they expect it. A no-op when the part is already standard (or absent).
    /// Run after a ClosedXML save (which authors the part at the legacy
    /// <c>.psmdcp</c> URI).
    /// </summary>
    public static void NormalizeAfterSave(string file)
    {
        using var package = Package.Open(file, FileMode.Open, FileAccess.ReadWrite);

        var relationships = package.GetRelationshipsByType(CorePropertiesRelationshipType).ToList();
        if (relationships.Count == 0)
        {
            return; // no core-properties part to relocate
        }

        var sourceUri = PackUriHelper.ResolvePartUri(PackageRootUri, relationships[0].TargetUri);
        if (sourceUri == StandardPartUri)
        {
            return; // already at the conventional path — nothing to do
        }

        if (!package.PartExists(sourceUri))
        {
            return; // dangling relationship; leave it for the reader's fallback
        }

        byte[] bytes;
        using (var source = package.GetPart(sourceUri).GetStream(FileMode.Open, FileAccess.Read))
        using (var buffer = new MemoryStream())
        {
            source.CopyTo(buffer);
            bytes = buffer.ToArray();
        }

        // Drop every core-properties relationship + the legacy part, then re-create
        // the part at the standard URI with the SAME content type and the SAME
        // bytes, and re-add the single relationship from the package root.
        foreach (var relationship in relationships)
        {
            package.DeleteRelationship(relationship.Id);
        }

        package.DeletePart(sourceUri);
        if (package.PartExists(StandardPartUri))
        {
            package.DeletePart(StandardPartUri);
        }

        var part = package.CreatePart(StandardPartUri, CorePropertiesContentType, CompressionOption.Normal);
        using (var destination = part.GetStream(FileMode.Create, FileAccess.Write))
        {
            destination.Write(bytes, 0, bytes.Length);
        }

        package.CreateRelationship(StandardPartUri, TargetMode.Internal, CorePropertiesRelationshipType);
    }
}
