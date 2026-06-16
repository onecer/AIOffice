using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// 1.8 paragraph-level visual primitives: shading (w:shd), border box (w:pBdr),
/// spacingBefore/After (w:spacing) and indentLeft/Right (w:ind). Every prop must
/// round-trip set -> get and leave the file validating clean. Also covers the
/// run-level + paragraph-fanout `font` prop.
/// </summary>
public sealed class ParagraphVisualTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })))["properties"]!;

    [Fact]
    public void Shading_round_trips_and_clears()
    {
        var file = CreateDoc(title: "Shaded");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"shading":"FEF9C3"}}]""");
        Assert.Equal("FEF9C3", Get(file, "/body/p[1]")["shading"]!.GetValue<string>());
        AssertValidatesClean(file);

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"shading":"none"}}]""");
        Assert.Null(Get(file, "/body/p[1]")["shading"]?.GetValue<string?>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Border_round_trips_style_color_width_sides()
    {
        var file = CreateDoc(title: "Bordered");

        Edit(file, """
            [{"op":"set","path":"/body/p[1]","props":
              {"border":{"style":"double","color":"2563EB","widthPt":1.5,"sides":"all"}}}]
            """);

        var border = Get(file, "/body/p[1]")["border"]!;
        Assert.Equal("double", border["style"]!.GetValue<string>());
        Assert.Equal("2563EB", border["color"]!.GetValue<string>());
        Assert.Equal(1.5, border["widthPt"]!.GetValue<double>());
        Assert.Equal("all", border["sides"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Border_one_side_then_none_clears()
    {
        var file = CreateDoc(title: "Underline rule");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"border":{"style":"single","sides":"bottom"}}}]""");
        var border = Get(file, "/body/p[1]")["border"]!;
        Assert.Equal("single", border["style"]!.GetValue<string>());
        Assert.Equal("bottom", border["sides"]!.GetValue<string>());
        AssertValidatesClean(file);

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"border":"none"}}]""");
        Assert.Null(Get(file, "/body/p[1]")["border"]);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Spacing_before_after_round_trip_in_points()
    {
        var file = CreateDoc(title: "Spaced");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"spacingBefore":12,"spacingAfter":6}}]""");

        var props = Get(file, "/body/p[1]");
        Assert.Equal(12, props["spacingBefore"]!.GetValue<double>());
        Assert.Equal(6, props["spacingAfter"]!.GetValue<double>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Indent_left_right_round_trip_in_centimeters()
    {
        var file = CreateDoc(title: "Indented");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"indentLeft":1.5,"indentRight":0.5}}]""");

        var props = Get(file, "/body/p[1]");
        Assert.Equal(1.5, props["indentLeft"]!.GetValue<double>());
        Assert.Equal(0.5, props["indentRight"]!.GetValue<double>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void All_visual_props_compose_in_one_op_and_validate()
    {
        var file = CreateDoc(title: "Callout block");

        Edit(file, """
            [{"op":"set","path":"/body/p[1]","props":{
                "text":"Important",
                "shading":"EFF6FF",
                "border":{"style":"single","color":"2563EB","widthPt":1,"sides":"all"},
                "spacingBefore":8,"spacingAfter":8,
                "indentLeft":1,"indentRight":1}}]
            """);

        var props = Get(file, "/body/p[1]");
        Assert.Equal("Important", props["text"]!.GetValue<string>());
        Assert.Equal("EFF6FF", props["shading"]!.GetValue<string>());
        Assert.Equal("single", props["border"]!["style"]!.GetValue<string>());
        Assert.Equal(8, props["spacingBefore"]!.GetValue<double>());
        Assert.Equal(1, props["indentLeft"]!.GetValue<double>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Visual_props_work_on_a_header_paragraph()
    {
        var file = CreateDoc(title: "With header");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"Confidential"}}]""");

        Edit(file, """
            [{"op":"set","path":"/header[1]/p[1]","props":
              {"shading":"FDE68A","border":{"style":"single","sides":"bottom"}}}]
            """);

        var props = Get(file, "/header[1]/p[1]");
        Assert.Equal("FDE68A", props["shading"]!.GetValue<string>());
        Assert.Equal("bottom", props["border"]!["sides"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Border_with_unknown_style_is_invalid_args()
    {
        var file = CreateDoc(title: "Bad border");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"border":{"style":"squiggle"}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.Contains("single", ex.Candidates!);
    }

    [Fact]
    public void Shading_on_a_run_is_unsupported_with_a_workaround()
    {
        var file = CreateDoc(title: "Run shading");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"shading":"FFFFFF"}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    [Fact]
    public void Run_font_round_trips()
    {
        var file = CreateDoc(title: "Fonted run");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"font":"Georgia"}}]""");

        Assert.Equal("Georgia", Get(file, "/body/p[1]/run[1]")["font"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Paragraph_font_fans_out_to_every_run()
    {
        var file = CreateDoc(title: "Fonted paragraph");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"font":"Verdana"}}]""");

        // get on the paragraph echoes the first run's font, and the run itself carries it.
        Assert.Equal("Verdana", Get(file, "/body/p[1]")["font"]!.GetValue<string>());
        Assert.Equal("Verdana", Get(file, "/body/p[1]/run[1]")["font"]!.GetValue<string>());
        AssertValidatesClean(file);
    }
}
