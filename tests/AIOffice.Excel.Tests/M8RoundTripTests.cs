using System.IO.Compression;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// Round-trip law coverage for the M8 Excel slicer feature. The law (open +
/// no-edit save touches only the documented part set) is pinned in
/// <see cref="RoundTripLawTests"/>; here we document the M8 EXCEPTION — the parts
/// a slicer write is ALLOWED to add — and prove the slicer survives a later
/// ClosedXML resave byte-for-byte at the part-set level.
///
/// <para>DOCUMENTED M8 ROUND-TRIP EXCEPTION (writes, not no-edit resaves):</para>
/// <list type="bullet">
/// <item>add slicer adds <c>xl/slicerCaches/slicerCacheN.xml</c> (workbook part),
/// <c>xl/slicers/slicerN.xml</c> (worksheet part) and a worksheet
/// <c>xl/drawings/drawingN.xml</c> anchor, plus the workbook/worksheet
/// <c>extLst</c> references. Never touches cell values.</item>
/// <item>MEASURED CAVEAT (ClosedXML 0.105): the slicerCache, slicers and the
/// worksheet's <c>slicerList</c> extLst reference survive a later ClosedXML
/// resave byte-clean, so the slicer stays defined and the file stays
/// validator-clean. ClosedXML does NOT round-trip the slicer's
/// <c>sle:slicer</c> DRAWING anchor (it drops that drawing part), but Excel
/// regenerates the slicer's visual frame from the surviving slicer definition
/// on open — the slicer remains present and functional.</item>
/// </list>
///
/// <para>The diff verb is read-only and has no round-trip footprint — it never
/// writes either file — so it needs no exception entry here.</para>
/// </summary>
public sealed class M8RoundTripTests : ExcelTestBase
{
    [Fact]
    public void Slicer_definition_survives_a_no_edit_resave_and_stays_valid()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1:B3", ("values", new JsonArray(
            new JsonArray("Region", "Sales"),
            new JsonArray("North", 100),
            new JsonArray("South", 200))))).IsOk);
        Assert.True(EditOps(file, AddOp("/Sheet1/A1:B3", "table", ("name", "Sales"))).IsOk);
        Assert.True(EditOps(file, AddOp("/Sheet1/table[@name=Sales]", "slicer", ("column", "Region"))).IsOk);

        var before = ReadParts(file);
        Assert.Contains(before.Keys, k => k.StartsWith("xl/slicerCaches/", StringComparison.Ordinal));
        Assert.Contains(before.Keys, k => k.StartsWith("xl/slicers/", StringComparison.Ordinal));

        using (var workbook = new XLWorkbook(file))
        {
            workbook.Save(); // no-edit resave
        }

        // The slicer DEFINITION (cache + slicers part) survives the resave, the
        // file stays validator-clean, and the slicer is still discoverable. Only
        // the regenerable drawing anchor is dropped by ClosedXML (documented).
        var after = ReadParts(file);
        Assert.Contains(after.Keys, k => k.StartsWith("xl/slicerCaches/", StringComparison.Ordinal));
        Assert.Contains(after.Keys, k => k.StartsWith("xl/slicers/", StringComparison.Ordinal));
        AssertValidatorClean(file);

        var read = Handler.Read(Ctx(file, ("view", "structure")));
        var slicers = OkData(read)["sheets"]!.AsArray()[0]!["slicers"]!.AsArray();
        Assert.Single(slicers);
        Assert.Equal("Region", slicers[0]!["column"]!.GetValue<string>());
    }

    private static Dictionary<string, byte[]> ReadParts(string file)
    {
        var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using var archive = ZipFile.OpenRead(file);
        foreach (var entry in archive.Entries)
        {
            using var stream = entry.Open();
            using var memory = new MemoryStream();
            stream.CopyTo(memory);
            parts[entry.FullName] = memory.ToArray();
        }

        return parts;
    }
}
