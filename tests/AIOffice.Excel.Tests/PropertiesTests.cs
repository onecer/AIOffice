using System.IO.Compression;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M7 document-properties slice: read core + custom properties, set them
/// through <c>{op:set, path:/properties}</c>, and the payoff — both blocks
/// reopen with the values aioffice wrote, custom properties keep their JSON
/// types, and a null custom value deletes the key.
/// </summary>
public sealed class PropertiesTests : ExcelTestBase
{
    private static EditOp SetProperties(JsonObject props) =>
        new() { Op = "set", Path = "/properties", Props = props };

    [Fact]
    public void Set_core_properties_then_read_returns_them()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetProperties(new JsonObject
        {
            ["title"] = "Q3 Revenue",
            ["author"] = "Onecer",
            ["subject"] = "Finance",
            ["company"] = "AIOffice",
        }));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var applied = Json(envelope)["data"]!["ops"]![0]!["applied"]!.AsArray()
            .Select(n => n!.GetValue<string>()).ToList();
        Assert.Contains("title", applied);
        AssertValidatorClean(file);

        var core = OkData(Handler.Read(Ctx(file, ("view", "properties"))))["core"]!;
        Assert.Equal("Q3 Revenue", core["title"]!.GetValue<string>());
        Assert.Equal("Onecer", core["author"]!.GetValue<string>());
        Assert.Equal("Finance", core["subject"]!.GetValue<string>());
        Assert.Equal("AIOffice", core["company"]!.GetValue<string>());

        // Raw oracle: the core properties part carries the title.
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Equal("Q3 Revenue", document.PackageProperties.Title);
    }

    /// <summary>
    /// The 1.0 hardening payoff: after a <c>set /properties</c>, the SAVED package
    /// carries the core properties in the CONVENTIONAL part <c>docProps/core.xml</c>
    /// (the SDK's CoreFilePropertiesPart) — not the legacy
    /// <c>package/services/metadata/core-properties/{GUID}.psmdcp</c> part — so
    /// unzip and Microsoft Office see the title where they expect it.
    /// </summary>
    [Fact]
    public void Set_title_lands_in_docProps_core_xml_not_the_legacy_psmdcp_part()
    {
        var file = CreateWorkbook();

        Assert.True(EditOps(file, SetProperties(new JsonObject
        {
            ["title"] = "Q3 Revenue",
            ["author"] = "Onecer",
        })).IsOk);

        // Raw zip oracle: the conventional part exists and carries dc:title; the
        // non-standard .psmdcp part is gone.
        using (var zip = ZipFile.OpenRead(file))
        {
            var core = zip.GetEntry("docProps/core.xml");
            Assert.NotNull(core);
            Assert.DoesNotContain(zip.Entries, e => e.FullName.EndsWith(".psmdcp", StringComparison.Ordinal));

            using var reader = new StreamReader(core!.Open());
            var xml = reader.ReadToEnd();
            Assert.Contains("<dc:title>Q3 Revenue</dc:title>", xml, StringComparison.Ordinal);
        }

        // SDK oracle: CoreFilePropertiesPart is at the standard URI and reads the title.
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.NotNull(document.CoreFilePropertiesPart);
        Assert.Equal("/docProps/core.xml", document.CoreFilePropertiesPart!.Uri.ToString());
        Assert.Equal("Q3 Revenue", document.PackageProperties.Title);
    }

    /// <summary>
    /// Even a freshly CREATED workbook (no <c>set /properties</c> yet) keeps its
    /// core part at the standard <c>docProps/core.xml</c> path — the legacy
    /// <c>.psmdcp</c> part never ships.
    /// </summary>
    [Fact]
    public void Created_workbook_uses_the_standard_core_part_path()
    {
        var file = CreateWorkbook();

        using var zip = ZipFile.OpenRead(file);
        Assert.NotNull(zip.GetEntry("docProps/core.xml"));
        Assert.DoesNotContain(zip.Entries, e => e.FullName.EndsWith(".psmdcp", StringComparison.Ordinal));
    }

    [Fact]
    public void Custom_properties_keep_their_types_across_reopen()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetProperties(new JsonObject
        {
            ["custom"] = new JsonObject
            {
                ["Region"] = "EU",
                ["Reviewed"] = true,
                ["Score"] = 9.5,
            },
        }));

        Assert.True(envelope.IsOk, envelope.ToJson());
        AssertValidatorClean(file);

        var custom = OkData(Handler.Get(Ctx(file, ("path", "/properties"))))["custom"]!;
        Assert.Equal("EU", custom["Region"]!.GetValue<string>());
        Assert.True(custom["Reviewed"]!.GetValue<bool>());
        Assert.Equal(9.5, custom["Score"]!.GetValue<double>());
    }

    [Fact]
    public void Null_custom_value_deletes_the_property()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, SetProperties(new JsonObject
        {
            ["custom"] = new JsonObject { ["Temp"] = "x", ["Keep"] = "y" },
        })).IsOk);

        Assert.True(EditOps(file, SetProperties(new JsonObject
        {
            ["custom"] = new JsonObject { ["Temp"] = null },
        })).IsOk);
        AssertValidatorClean(file);

        var custom = OkData(Handler.Read(Ctx(file, ("view", "properties"))))["custom"]!.AsObject();
        Assert.False(custom.ContainsKey("Temp"));
        Assert.True(custom.ContainsKey("Keep"));
    }

    [Fact]
    public void Custom_iso_date_string_round_trips_as_a_date()
    {
        var file = CreateWorkbook();

        Assert.True(EditOps(file, SetProperties(new JsonObject
        {
            ["custom"] = new JsonObject { ["Reviewed"] = "2026-06-13" },
        })).IsOk);

        var custom = OkData(Handler.Read(Ctx(file, ("view", "properties"))))["custom"]!;
        Assert.StartsWith("2026-06-13", custom["Reviewed"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void The_cellStyle_registry_property_is_hidden_from_the_custom_block()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, new EditOp
        {
            Op = "add",
            Path = "/styles",
            Type = "cellStyle",
            Props = new JsonObject { ["name"] = "S", ["bold"] = true },
        }).IsOk);

        var custom = OkData(Handler.Read(Ctx(file, ("view", "properties"))))["custom"]!.AsObject();
        Assert.False(custom.ContainsKey("_aioffice_cellStyles"));
    }

    [Fact]
    public void Unknown_core_property_is_invalid_args_with_candidates()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetProperties(new JsonObject { ["titel"] = "x" }));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("title", envelope.Error.Candidates!);
    }

    [Fact]
    public void Reserved_registry_name_cannot_be_set_as_a_custom_property()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, SetProperties(new JsonObject
        {
            ["custom"] = new JsonObject { ["_aioffice_cellStyles"] = "hack" },
        }));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("reserved", envelope.Error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Remove_on_properties_is_rejected()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, RemoveOp("/properties"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("set", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    /// <summary>
    /// Migrate-on-read: an OLD file that ClosedXML saved with only the legacy
    /// <c>.psmdcp</c> core part still reads its title back (the reader follows the
    /// package-root core-properties relationship, not a URI), and the next
    /// aioffice write moves the properties to the standard
    /// <c>docProps/core.xml</c> while preserving the existing title.
    /// </summary>
    [Fact]
    public void Legacy_psmdcp_file_reads_its_title_then_a_write_relocates_the_part()
    {
        // Build a legacy fixture exactly as ClosedXML's default save produces it:
        // core props land in package/services/metadata/core-properties/{GUID}.psmdcp.
        var file = NewFile("legacy.xlsx");
        using (var workbook = new ClosedXML.Excel.XLWorkbook())
        {
            workbook.AddWorksheet("Sheet1").Cell("A1").Value = 1;
            workbook.Properties.Title = "Legacy Title";
            workbook.Properties.Author = "OldAuthor";
            workbook.SaveAs(file);
        }

        // Precondition: the fixture really carries the legacy part, not the standard one.
        using (var zip = ZipFile.OpenRead(file))
        {
            Assert.Contains(zip.Entries, e => e.FullName.EndsWith(".psmdcp", StringComparison.Ordinal));
            Assert.Null(zip.GetEntry("docProps/core.xml"));
        }

        // Read still works (fallback via the relationship).
        var core = OkData(Handler.Read(Ctx(file, ("view", "properties"))))["core"]!;
        Assert.Equal("Legacy Title", core["title"]!.GetValue<string>());
        Assert.Equal("OldAuthor", core["author"]!.GetValue<string>());

        // Any aioffice write relocates the part to the standard path, title intact.
        Assert.True(EditOps(file, SetOp("/Sheet1/A2", ("value", 2))).IsOk);
        AssertValidatorClean(file);

        using (var zip = ZipFile.OpenRead(file))
        {
            Assert.NotNull(zip.GetEntry("docProps/core.xml"));
            Assert.DoesNotContain(zip.Entries, e => e.FullName.EndsWith(".psmdcp", StringComparison.Ordinal));
        }

        var migrated = OkData(Handler.Read(Ctx(file, ("view", "properties"))))["core"]!;
        Assert.Equal("Legacy Title", migrated["title"]!.GetValue<string>());
        Assert.Equal("OldAuthor", migrated["author"]!.GetValue<string>());
    }
}
