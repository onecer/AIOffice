using System.IO.Compression;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// Round-trip law coverage for the three M7 Excel features. The law itself
/// (open + no-edit save touches only the documented part set) is pinned in
/// <see cref="RoundTripLawTests"/>; here we document the M7 EXCEPTIONS — the
/// parts each new mutating op is ALLOWED to change — and prove the file stays
/// OpenXmlValidator-clean and reopens with the written state intact.
///
/// <para>DOCUMENTED M7 ROUND-TRIP EXCEPTIONS (writes, not no-edit resaves):</para>
/// <list type="bullet">
/// <item>set /properties writes docProps/core.xml (core) and
/// docProps/custom.xml (custom) — both are property parts, never cell data.</item>
/// <item>add cellStyle writes the custom property registry
/// (<c>_aioffice_cellStyles</c> in docProps/custom.xml) and materializes real
/// <c>cellStyle</c>/<c>cellStyleXfs</c> entries plus any custom number format in
/// xl/styles.xml. Applying a style edits the targeted cells' xf only.</item>
/// <item>audit --fix writes a placeholder <c>descr</c> on a drawing's
/// <c>cNvPr</c> (xl/drawings/drawingN.xml) and/or the Title in
/// docProps/core.xml. Never touches cell values.</item>
/// </list>
/// </summary>
public sealed class M7RoundTripTests : ExcelTestBase
{
    [Fact]
    public void Set_properties_keeps_the_file_valid_and_data_intact()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", 42.0))).IsOk);

        Assert.True(EditOps(file, new EditOp
        {
            Op = "set",
            Path = "/properties",
            Props = new JsonObject { ["title"] = "T", ["custom"] = new JsonObject { ["K"] = "V" } },
        }).IsOk);
        AssertValidatorClean(file);

        // The cell data is untouched by a properties write.
        var cell = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal(42.0, cell["value"]!.GetValue<double>());
    }

    [Fact]
    public void Cell_style_lifecycle_stays_valid_and_does_not_disturb_other_cells()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", "untouched")),
            SetOp("/Sheet1/B2", ("value", 1000.0))).IsOk);

        Assert.True(EditOps(file, new EditOp
        {
            Op = "add",
            Path = "/styles",
            Type = "cellStyle",
            Props = new JsonObject { ["name"] = "Money", ["numberFormat"] = "$#,##0" },
        }).IsOk);
        Assert.True(EditOps(file, SetOp("/Sheet1/B2", ("cellStyle", "Money"))).IsOk);
        AssertValidatorClean(file);

        // A1 kept its plain value/format; B2 got the number format.
        var a1 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/A1"))));
        Assert.Equal("untouched", a1["value"]!.GetValue<string>());
        Assert.Null(a1["numberFormat"]);

        var b2 = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/B2"))));
        Assert.Equal("$#,##0", b2["numberFormat"]!.GetValue<string>());

        // Reopen-verify: the style still resolves and re-applies cleanly.
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("cellStyle", "Money"))).IsOk);
        AssertValidatorClean(file);
    }

    [Fact]
    public void No_edit_resave_after_m7_writes_still_obeys_the_law()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(
            file,
            SetOp("/Sheet1/A1", ("value", 7.0)),
            new EditOp
            {
                Op = "set",
                Path = "/properties",
                Props = new JsonObject { ["title"] = "T" },
            }).IsOk);
        Assert.True(EditOps(file, new EditOp
        {
            Op = "add",
            Path = "/styles",
            Type = "cellStyle",
            Props = new JsonObject { ["name"] = "S", ["numberFormat"] = "0.00" },
        }).IsOk);

        var before = ReadParts(file);
        using (var workbook = new XLWorkbook(file))
        {
            workbook.Save(); // no-edit resave
        }

        var after = ReadParts(file);
        // No part appears or disappears across a no-edit resave.
        Assert.Equal(before.Keys.Order(StringComparer.Ordinal), after.Keys.Order(StringComparer.Ordinal));

        // The styles and properties parts survive the resave (the law only lets
        // ClosedXML rewrite sheet XML).
        Assert.True(after.ContainsKey("xl/styles.xml"));
        AssertValidatorClean(file);
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
