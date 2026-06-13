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
}
