using System.IO.Compression;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class DocumentPropertiesTests : WordTestBase
{
    private static readonly XNamespace DcNs = "http://purl.org/dc/elements/1.1/";

    [Fact]
    public void Set_title_writes_the_standard_core_part_not_the_legacy_psmdcp_part()
    {
        var file = CreateDoc(title: "Draft");
        Edit(file, """[{"op":"set","path":"/properties","props":{"title":"Q3 Report","author":"Ada Lovelace"}}]""");

        // 1) The SDK sees the standard CoreFilePropertiesPart at docProps/core.xml.
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.NotNull(doc.CoreFilePropertiesPart);
            Assert.Equal("/docProps/core.xml", doc.CoreFilePropertiesPart!.Uri.ToString());

            using var stream = doc.CoreFilePropertiesPart.GetStream();
            var root = XDocument.Load(stream).Root!;
            Assert.Equal("Q3 Report", root.Element(DcNs + "title")!.Value);
            Assert.Equal("Ada Lovelace", root.Element(DcNs + "creator")!.Value);
        }

        // 2) Unzipping the package shows docProps/core.xml and NOT a legacy .psmdcp part.
        using var zip = ZipFile.OpenRead(file);
        var entries = zip.Entries.Select(e => e.FullName).ToList();
        Assert.Contains("docProps/core.xml", entries);
        Assert.DoesNotContain(entries, e => e.EndsWith(".psmdcp", StringComparison.OrdinalIgnoreCase));

        var coreEntry = zip.GetEntry("docProps/core.xml")!;
        using var reader = new StreamReader(coreEntry.Open());
        var xml = reader.ReadToEnd();
        Assert.Contains("<dc:title>Q3 Report</dc:title>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void Set_core_properties_then_get_round_trips_through_reopen()
    {
        var file = CreateDoc(title: "Draft");

        Edit(file, """
            [{"op":"set","path":"/properties","props":{
                "title":"Q3 Report","subject":"Quarterly","author":"Ada Lovelace",
                "keywords":"finance;q3","category":"Reports","comments":"Internal only"
            }}]
            """);

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/properties" })));
        Assert.Equal("properties", got["type"]!.GetValue<string>());
        var core = got["properties"]!["core"]!;
        Assert.Equal("Q3 Report", core["title"]!.GetValue<string>());
        Assert.Equal("Quarterly", core["subject"]!.GetValue<string>());
        Assert.Equal("Ada Lovelace", core["author"]!.GetValue<string>());
        Assert.Equal("finance;q3", core["keywords"]!.GetValue<string>());
        Assert.Equal("Reports", core["category"]!.GetValue<string>());
        Assert.Equal("Internal only", core["comments"]!.GetValue<string>());

        // The values truly landed on the package, not just in our shape.
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Equal("Q3 Report", doc.PackageProperties.Title);
        Assert.Equal("Ada Lovelace", doc.PackageProperties.Creator);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Legacy_psmdcp_title_is_read_and_migrated_to_the_standard_part_on_write()
    {
        // Build a file whose ONLY core metadata is the non-standard legacy psmdcp part
        // (what System.IO.Packaging's PackageProperties setter creates when no core part exists).
        var file = CreateDoc(title: "Body Heading");
        using (var doc = WordprocessingDocument.Open(file, isEditable: true))
        {
            Assert.Null(doc.CoreFilePropertiesPart); // create() writes no core part
            doc.PackageProperties.Title = "Legacy Title";
            doc.PackageProperties.Creator = "Old Author";
        }

        // The legacy part is the non-standard one; docProps/core.xml does not exist yet.
        using (var zip = ZipFile.OpenRead(file))
        {
            Assert.Contains(zip.Entries, e => e.FullName.EndsWith(".psmdcp", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(zip.Entries, e => e.FullName == "docProps/core.xml");
        }

        // Migrate-on-read: the handler still surfaces the legacy title.
        var before = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/properties" })))["properties"]!["core"]!;
        Assert.Equal("Legacy Title", before["title"]!.GetValue<string>());
        Assert.Equal("Old Author", before["author"]!.GetValue<string>());

        // A write migrates everything into the standard docProps/core.xml part, keeping legacy values.
        Edit(file, """[{"op":"set","path":"/properties","props":{"subject":"Migrated"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.NotNull(doc.CoreFilePropertiesPart);
            using var stream = doc.CoreFilePropertiesPart!.GetStream();
            var root = XDocument.Load(stream).Root!;
            Assert.Equal("Legacy Title", root.Element(DcNs + "title")!.Value); // preserved through migration
            Assert.Equal("Old Author", root.Element(DcNs + "creator")!.Value);
            Assert.Equal("Migrated", root.Element(DcNs + "subject")!.Value);
        }

        using (var zip = ZipFile.OpenRead(file))
        {
            Assert.Contains(zip.Entries, e => e.FullName == "docProps/core.xml");
            // The non-standard legacy part is gone — values moved, not duplicated.
            Assert.DoesNotContain(zip.Entries, e => e.FullName.EndsWith(".psmdcp", StringComparison.OrdinalIgnoreCase));
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Core_dates_write_w3cdtf_to_the_standard_part_and_round_trip()
    {
        var file = CreateDoc(title: "Dated");
        Edit(file, """[{"op":"set","path":"/properties","props":{"created":"2026-06-13T10:00:00Z","modified":"2026-06-14"}}]""");

        // The envelope reads the ISO dates back.
        var core = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/properties" })))["properties"]!["core"]!;
        Assert.Equal("2026-06-13T10:00:00Z", core["created"]!.GetValue<string>());
        Assert.StartsWith("2026-06-14", core["modified"]!.GetValue<string>(), StringComparison.Ordinal);

        // The dates land in docProps/core.xml as W3CDTF-typed dcterms elements; the package validates clean.
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        using var stream = doc.CoreFilePropertiesPart!.GetStream();
        var xml = new StreamReader(stream).ReadToEnd();
        Assert.Contains("dcterms:W3CDTF", xml, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Custom_properties_keep_their_json_types_on_reopen()
    {
        var file = CreateDoc(title: "Typed");

        Edit(file, """
            [{"op":"set","path":"/properties","props":{"custom":{
                "Project":"Acme","Reviewed":true,"Budget":1000,"DueDate":"2026-06-13"
            }}}]
            """);

        var custom = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "properties" })))["properties"]!["custom"]!;
        Assert.Equal("Acme", custom["Project"]!.GetValue<string>());
        Assert.True(custom["Reviewed"]!.GetValue<bool>());
        Assert.Equal(1000, custom["Budget"]!.GetValue<double>());
        Assert.Contains("2026-06-13", custom["DueDate"]!.GetValue<string>(), StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void View_properties_reports_core_and_custom_sections()
    {
        var file = CreateDoc(title: "Sections");
        Edit(file, """[{"op":"set","path":"/properties","props":{"title":"With Custom","custom":{"Team":"Platform"}}}]""");

        var props = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "properties" })))["properties"]!;
        Assert.Equal("With Custom", props["core"]!["title"]!.GetValue<string>());
        Assert.Equal("Platform", props["custom"]!["Team"]!.GetValue<string>());
    }

    [Fact]
    public void Setting_a_custom_property_twice_updates_in_place()
    {
        var file = CreateDoc(title: "Once");
        Edit(file, """[{"op":"set","path":"/properties","props":{"custom":{"Status":"Draft"}}}]""");
        Edit(file, """[{"op":"set","path":"/properties","props":{"custom":{"Status":"Final"}}}]""");

        var custom = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/properties" })))["properties"]!["custom"]!;
        Assert.Equal("Final", custom["Status"]!.GetValue<string>());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var defined = doc.CustomFilePropertiesPart!.Properties!
            .Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>()
            .Count(p => p.Name?.Value == "Status");
        Assert.Equal(1, defined); // not duplicated
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_properties_with_no_recognized_keys_is_invalid_args()
    {
        var file = CreateDoc(title: "Empty");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/properties","props":{"custom":{}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Unknown_core_property_is_unsupported_feature_with_candidates()
    {
        var file = CreateDoc(title: "Unknown");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/properties","props":{"manager":"Bob"}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("title", ex.Candidates!);
    }

    [Fact]
    public void Author_alias_maps_to_creator()
    {
        var file = CreateDoc(title: "Aliased");
        Edit(file, """[{"op":"set","path":"/properties","props":{"author":"Grace Hopper"}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Equal("Grace Hopper", doc.PackageProperties.Creator);
    }
}
