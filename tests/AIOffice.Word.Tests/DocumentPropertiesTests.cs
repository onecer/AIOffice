using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class DocumentPropertiesTests : WordTestBase
{
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
