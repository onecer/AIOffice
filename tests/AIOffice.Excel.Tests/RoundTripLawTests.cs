using System.IO.Compression;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The round-trip law: opening a file and saving it WITHOUT edits must leave
/// every zip part byte-identical — or this test documents exactly why a part
/// legitimately changes.
///
/// MEASURED REALITY for the ClosedXML engine (0.105.0): a no-edit open/save
/// rewrites <c>xl/worksheets/sheetN.xml</c> and ONLY that part, in exactly two
/// ways:
/// <list type="number">
/// <item>The namespace declarations on the root <c>&lt;x:worksheet&gt;</c>
/// element are re-emitted in a different attribute order (<c>mc:Ignorable</c>
/// moves after the xmlns declarations). Semantically a no-op.</item>
/// <item>Cached formula values (<c>&lt;v&gt;</c> inside cells that have an
/// <c>&lt;f&gt;</c>) are DROPPED, because ClosedXML marks loaded formulas as
/// needing recalculation and its default save omits unevaluated results. The
/// formula text itself is untouched, and Excel recomputes on open. aioffice's
/// own edit/template pipeline always saves with formula evaluation, so files
/// that go through aioffice get their cached values rewritten every time.</item>
/// </list>
/// Every other part — workbook, styles, sharedStrings, tables, calcChain,
/// theme, content types, rels, docProps — survives byte-for-byte.
/// </summary>
public sealed class RoundTripLawTests : ExcelTestBase
{
    /// <summary>The honesty baseline: parts ClosedXML rewrites on a no-edit save.</summary>
    private static readonly IReadOnlySet<string> DocumentedChangedParts =
        new HashSet<string>(StringComparer.Ordinal) { "xl/worksheets/sheet1.xml" };

    [Fact]
    public void NoEdit_resave_changes_exactly_the_documented_parts()
    {
        var original = BuildRichFixture();
        var resaved = Path.Combine(Dir, "resaved.xlsx");
        File.Copy(original, resaved);

        using (var workbook = new XLWorkbook(resaved))
        {
            workbook.Save(); // no edits at all
        }

        var before = ReadParts(original);
        var after = ReadParts(resaved);

        // No part may appear or disappear.
        Assert.Equal(before.Keys.Order(StringComparer.Ordinal), after.Keys.Order(StringComparer.Ordinal));

        var changed = before.Keys
            .Where(name => !before[name].AsSpan().SequenceEqual(after[name]))
            .ToHashSet(StringComparer.Ordinal);
        Assert.Equal(DocumentedChangedParts, changed);

        // After factoring out the two documented diffs (attribute order,
        // dropped cached formula values) the parts must be identical.
        foreach (var name in changed)
        {
            var beforeXml = Normalize(LoadXml(before[name]).Root!);
            var afterXml = Normalize(LoadXml(after[name]).Root!);
            Assert.True(
                XNode.DeepEquals(beforeXml, afterXml),
                $"{name} changed beyond the two documented diffs (attribute order, dropped cached <v>)");
        }

        AssertValidatorClean(resaved);
    }

    /// <summary>A workbook exercising values, formulas, styles, merges and a table.</summary>
    private string BuildRichFixture()
    {
        var file = CreateWorkbook("fixture.xlsx");
        var envelope = EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 1)),
            SetOp("/Sheet1/A2", ("value", 2)),
            SetOp("/Sheet1/A3", ("value", "=SUM(A1:A2)")),
            SetOp("/Sheet1/B1", ("value", "hello"), ("bold", true)),
            SetOp("/Sheet1/B2", ("value", "2024-05-01")),
            SetOp("/Sheet1/C1", ("value", 3.14159), ("numberFormat", "0.00")),
            SetOp("/Sheet1/D1:E2", ("merge", true)),
            SetOp("/Sheet1/A10:C12", ("values", new JsonArray(
                new JsonArray("Name", "Qty", "Price"),
                new JsonArray("ant", 5, 1.5),
                new JsonArray("bee", 7, 2.5)))),
            AddOp("/Sheet1/A10:C12", "table", ("name", "Stock")));
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    private static Dictionary<string, byte[]> ReadParts(string file)
    {
        using var zip = ZipFile.OpenRead(file);
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries)
        {
            using var stream = entry.Open();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            parts[entry.FullName] = buffer.ToArray();
        }

        return parts;
    }

    /// <summary>Loads part bytes as XML (handles the UTF-8 BOM the parts carry).</summary>
    private static XDocument LoadXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        return XDocument.Load(stream);
    }

    /// <summary>
    /// Rebuilds an element tree with attributes sorted (factors out diff 1) and
    /// cached formula results removed — a <c>&lt;v&gt;</c> next to an
    /// <c>&lt;f&gt;</c> sibling (factors out diff 2).
    /// </summary>
    private static XElement Normalize(XElement element) => new(
        element.Name,
        element.Attributes()
            .OrderBy(a => a.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(a => a.Name.LocalName, StringComparer.Ordinal),
        element.Nodes()
            .Where(n => n is not XElement { Name.LocalName: "v", Parent: not null } v ||
                        !v.Parent!.Elements().Any(s => s.Name.LocalName == "f"))
            .Select(n => n is XElement child ? Normalize(child) : n));
}
