using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class StyleTests : WordTestBase
{
    private JsonArray Styles(string file) =>
        Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "styles" })))["styles"]!.AsArray();

    private JsonNode GetStyle(string file, string id) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = $"/style[@id={id}]" })))["properties"]!;

    [Fact]
    public void Styles_view_lists_defaults_with_builtin_flags()
    {
        var file = CreateDoc(title: "Styled doc");

        var styles = Styles(file);

        var normal = styles.Single(s => s!["id"]!.GetValue<string>() == "Normal")!;
        Assert.True(normal["builtin"]!.GetValue<bool>());
        Assert.Equal("paragraph", normal["kind"]!.GetValue<string>());

        var heading1 = styles.Single(s => s!["id"]!.GetValue<string>() == "Heading1")!;
        Assert.True(heading1["builtin"]!.GetValue<bool>());
        Assert.True(heading1["inUse"]!.GetValue<bool>()); // the title paragraph uses it
        Assert.Equal("Normal", heading1["basedOn"]!.GetValue<string>());
    }

    [Fact]
    public void Add_custom_style_then_apply_it_to_a_paragraph()
    {
        var file = CreateDoc(title: "Callouts");

        Edit(file, """
            [{"op":"add","path":"/styles","type":"style","props":
              {"id":"Callout","kind":"paragraph","basedOn":"Normal","bold":true,"color":"1F4E79",
               "fontSize":12,"alignment":"center","spacingBefore":6,"spacingAfter":6}}]
            """);
        Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"text":"Note well.","style":"Callout"}}]""");

        var style = GetStyle(file, "Callout");
        Assert.Equal("Callout", style["id"]!.GetValue<string>());
        Assert.True(style["bold"]!.GetValue<bool>());
        Assert.Equal("1F4E79", style["color"]!.GetValue<string>());
        Assert.Equal(12, style["fontSize"]!.GetValue<double>());
        Assert.Equal("center", style["alignment"]!.GetValue<string>());
        Assert.Equal(6, style["spacingBefore"]!.GetValue<double>());
        Assert.Equal(6, style["spacingAfter"]!.GetValue<double>());
        Assert.False(style["builtin"]!.GetValue<bool>());
        Assert.True(style["inUse"]!.GetValue<bool>());

        var paragraph = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[2]" })))["properties"]!;
        Assert.Equal("Callout", paragraph["style"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_modifies_an_existing_style_definition()
    {
        var file = CreateDoc(title: "Tunable");
        Edit(file, """[{"op":"add","path":"/styles","type":"style","props":{"id":"Callout","bold":true}}]""");

        Edit(file, """[{"op":"set","path":"/style[@id=Callout]","props":{"color":"FF0000","fontSize":14}}]""");

        var style = GetStyle(file, "Callout");
        Assert.Equal("FF0000", style["color"]!.GetValue<string>());
        Assert.Equal(14, style["fontSize"]!.GetValue<double>());
        Assert.True(style["bold"]!.GetValue<bool>()); // untouched props persist
        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_builtin_style_is_refused_with_workaround()
    {
        var file = CreateDoc(title: "Protected");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"remove","path":"/style[@id=Normal]"}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("set", ex.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Remove_custom_style_deletes_the_definition()
    {
        var file = CreateDoc(title: "Cleanup");
        Edit(file, """[{"op":"add","path":"/styles","type":"style","props":{"id":"Temp"}}]""");
        Assert.Contains(Styles(file), s => s!["id"]!.GetValue<string>() == "Temp");

        Edit(file, """[{"op":"remove","path":"/style[@id=Temp]"}]""");

        Assert.DoesNotContain(Styles(file), s => s!["id"]!.GetValue<string>() == "Temp");
        AssertValidatesClean(file);
    }

    [Fact]
    public void Duplicate_style_id_is_invalid_args()
    {
        var file = CreateDoc(title: "Dupes");
        Edit(file, """[{"op":"add","path":"/styles","type":"style","props":{"id":"Callout"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/styles","type":"style","props":{"id":"Callout"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("/style[@id=Callout]", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Character_style_rejects_paragraph_level_props()
    {
        var file = CreateDoc(title: "Chars");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """
                [{"op":"add","path":"/styles","type":"style","props":
                  {"id":"Emph","kind":"character","alignment":"center"}}]
                """));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("paragraph", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Character_style_applies_to_runs()
    {
        var file = CreateDoc(title: "Inline style");

        Edit(file, """[{"op":"add","path":"/styles","type":"style","props":{"id":"Emph","kind":"character","italic":true,"color":"880000"}}]""");
        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"style":"Emph"}}]""");

        var run = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]/run[1]" })))["properties"]!;
        Assert.Equal("Emph", run["style"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Unknown_basedOn_is_invalid_args_with_candidates()
    {
        var file = CreateDoc(title: "Based");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/styles","type":"style","props":{"id":"X","basedOn":"NoSuchStyle"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.Contains("Normal", ex.Candidates!);
    }

    [Fact]
    public void Missing_style_id_is_invalid_path_with_candidates()
    {
        var file = CreateDoc(title: "Missing");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/style[@id=Ghost]","props":{"bold":true}}]"""));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/style[@id=Normal]", ex.Candidates!);
    }

    [Fact]
    public void Style_kind_table_is_unsupported()
    {
        var file = CreateDoc(title: "Tables");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/styles","type":"style","props":{"id":"Grid","kind":"table"}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    [Fact]
    public void Removing_an_in_use_style_reports_the_fallback()
    {
        var file = CreateDoc(title: "Used");
        Edit(file, """
            [
              {"op":"add","path":"/styles","type":"style","props":{"id":"Loose"}},
              {"op":"set","path":"/body/p[2]","props":{"text":"styled","style":"Loose"}}
            ]
            """);

        var envelope = Edit(file, """[{"op":"remove","path":"/style[@id=Loose]"}]""");

        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Contains("Normal", summary["note"]!.GetValue<string>(), StringComparison.Ordinal);
        AssertValidatesClean(file);
    }
}
