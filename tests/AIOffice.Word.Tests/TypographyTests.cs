using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.Json.Nodes;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// 1.10 docx typography primitives. Character props (highlight, strike, doubleStrike,
/// smallCaps, allCaps, super/subscript, characterSpacing) on runs (and fanned out from a
/// paragraph), and paragraph props (lineSpacing, keepNext/keepLines/pageBreakBefore/
/// widowControl, outlineLevel, tabStops). Every prop round-trips set -> get, and the unit
/// conversions are asserted against the saved XML (the main risk). validate stays clean.
/// </summary>
public sealed class TypographyTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })))["properties"]!;

    /// <summary>The first run's w:rPr in the saved body (typography lives here).</summary>
    private static RunProperties FirstRunPr(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First().RunProperties!;
    }

    /// <summary>The first paragraph's w:pPr in the saved body.</summary>
    private static ParagraphProperties FirstParaPr(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>().First().ParagraphProperties!;
    }

    // ----------------------------------------------------------- run character props

    [Fact]
    public void Highlight_named_color_round_trips_and_writes_named_val()
    {
        var file = CreateDoc(title: "Highlighted");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"clause","highlight":"yellow"}}]""");

        Assert.Equal("yellow", Get(file, "/body/p[1]/run[1]")["highlight"]!.GetValue<string>());
        Assert.Equal("yellow", FirstRunPr(file).Highlight!.Val!.ToString());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Highlight_is_case_insensitive()
    {
        var file = CreateDoc(title: "Highlighted ci");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"x","highlight":"DarkGreen"}}]""");

        Assert.Equal("darkGreen", Get(file, "/body/p[1]/run[1]")["highlight"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Highlight_hex_is_rejected_with_the_named_candidate_list()
    {
        var file = CreateDoc(title: "Bad highlight");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"highlight":"FFFF00"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.Contains("yellow", ex.Candidates!);
        Assert.Contains("none", ex.Candidates!);
    }

    [Fact]
    public void Highlight_none_round_trips()
    {
        var file = CreateDoc(title: "Highlight none");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"x","highlight":"none"}}]""");

        Assert.Equal("none", Get(file, "/body/p[1]/run[1]")["highlight"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Strike_and_doubleStrike_round_trip()
    {
        var file = CreateDoc(title: "Struck");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"old","strike":true}}]""");
        Assert.True(Get(file, "/body/p[1]/run[1]")["strike"]!.GetValue<bool>());
        Assert.True(FirstRunPr(file).Strike!.Val!.Value);
        AssertValidatesClean(file);

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"doubleStrike":true}}]""");
        Assert.True(Get(file, "/body/p[1]/run[1]")["doubleStrike"]!.GetValue<bool>());
        Assert.NotNull(FirstRunPr(file).DoubleStrike);
        AssertValidatesClean(file);
    }

    [Fact]
    public void SmallCaps_and_allCaps_round_trip()
    {
        var file = CreateDoc(title: "Caps");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"Title","smallCaps":true,"allCaps":false}}]""");

        var run = Get(file, "/body/p[1]/run[1]");
        Assert.True(run["smallCaps"]!.GetValue<bool>());
        Assert.False(run["allCaps"]!.GetValue<bool>());
        Assert.NotNull(FirstRunPr(file).SmallCaps);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Superscript_then_subscript_keeps_exactly_one_vertAlign()
    {
        var file = CreateDoc(title: "Footnote mark");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"x","superscript":true}}]""");
        Assert.True(Get(file, "/body/p[1]/run[1]")["superscript"]!.GetValue<bool>());
        Assert.Equal(VerticalPositionValues.Superscript, FirstRunPr(file).VerticalTextAlignment!.Val!.Value);
        AssertValidatesClean(file);

        // Flipping to subscript replaces the same single w:vertAlign element.
        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"subscript":true}}]""");
        var rPr = FirstRunPr(file);
        Assert.Single(rPr.Elements<VerticalTextAlignment>());
        Assert.Equal(VerticalPositionValues.Subscript, rPr.VerticalTextAlignment!.Val!.Value);
        var run = Get(file, "/body/p[1]/run[1]");
        Assert.False(run["superscript"]!.GetValue<bool>());
        Assert.True(run["subscript"]!.GetValue<bool>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Subscript_false_removes_the_vertAlign()
    {
        var file = CreateDoc(title: "Baseline");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"x","subscript":true}}]""");
        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"subscript":false}}]""");

        Assert.Empty(FirstRunPr(file).Elements<VerticalTextAlignment>());
        Assert.False(Get(file, "/body/p[1]/run[1]")["subscript"]!.GetValue<bool>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void CharacterSpacing_one_point_writes_val_20_and_round_trips()
    {
        var file = CreateDoc(title: "Spaced chars");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"wide","characterSpacing":1}}]""");

        Assert.Equal(1, Get(file, "/body/p[1]/run[1]")["characterSpacing"]!.GetValue<double>());
        Assert.Equal(20, FirstRunPr(file).Spacing!.Val!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void CharacterSpacing_negative_condenses()
    {
        var file = CreateDoc(title: "Condensed");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"tight","characterSpacing":-0.5}}]""");

        Assert.Equal(-0.5, Get(file, "/body/p[1]/run[1]")["characterSpacing"]!.GetValue<double>());
        Assert.Equal(-10, FirstRunPr(file).Spacing!.Val!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Character_props_fan_out_when_set_on_a_paragraph()
    {
        var file = CreateDoc(title: "Highlighted para");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"clause","highlight":"cyan","strike":true}}]""");

        // get on the paragraph echoes the first run; the run itself carries them.
        Assert.Equal("cyan", Get(file, "/body/p[1]")["highlight"]!.GetValue<string>());
        Assert.True(Get(file, "/body/p[1]")["strike"]!.GetValue<bool>());
        Assert.Equal("cyan", Get(file, "/body/p[1]/run[1]")["highlight"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    // ----------------------------------------------------------- emphasis mark (w:em)

    [Theory]
    [InlineData("dot")]
    [InlineData("comma")]
    [InlineData("circle")]
    [InlineData("underDot")]
    [InlineData("none")]
    public void EmphasisMark_round_trips_and_writes_the_val(string mark)
    {
        var file = CreateDoc(title: "Emphasized");

        Edit(file, "[{\"op\":\"set\",\"path\":\"/body/p[1]/run[1]\",\"props\":{\"text\":\"重点\",\"emphasisMark\":\"" + mark + "\"}}]");

        Assert.Equal(mark, Get(file, "/body/p[1]/run[1]")["emphasisMark"]!.GetValue<string>());
        Assert.Equal(mark, FirstRunPr(file).Emphasis!.Val!.ToString());
        AssertValidatesClean(file);
    }

    [Fact]
    public void EmphasisMark_is_case_insensitive_and_normalizes_to_the_canonical_token()
    {
        var file = CreateDoc(title: "Emphasized ci");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"x","emphasisMark":"UNDERDOT"}}]""");

        Assert.Equal("underDot", Get(file, "/body/p[1]/run[1]")["emphasisMark"]!.GetValue<string>());
        Assert.Equal("underDot", FirstRunPr(file).Emphasis!.Val!.ToString());
        AssertValidatesClean(file);
    }

    [Fact]
    public void EmphasisMark_fans_out_when_set_on_a_paragraph()
    {
        var file = CreateDoc(title: "Emphasized para");

        // Set on the paragraph: fans out to every (direct) run's w:em, like characterSpacing.
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"重点","emphasisMark":"circle"}}]""");

        // get on the paragraph echoes the first run; the run itself carries it.
        Assert.Equal("circle", Get(file, "/body/p[1]")["emphasisMark"]!.GetValue<string>());
        Assert.Equal("circle", Get(file, "/body/p[1]/run[1]")["emphasisMark"]!.GetValue<string>());
        Assert.Equal("circle", FirstRunPr(file).Emphasis!.Val!.ToString());
        AssertValidatesClean(file);
    }

    [Fact]
    public void EmphasisMark_absent_reads_back_null()
    {
        var file = CreateDoc(title: "Plain");

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"text":"plain"}}]""");

        Assert.Null(Get(file, "/body/p[1]/run[1]")["emphasisMark"]?.GetValue<string?>());
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<Emphasis>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void EmphasisMark_unknown_token_is_invalid_args_with_the_candidate_list()
    {
        var file = CreateDoc(title: "Bad emphasis");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"emphasisMark":"star"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.Equal(new[] { "none", "dot", "comma", "circle", "underDot" }, ex.Candidates!);
    }

    // ----------------------------------------------------------- paragraph line spacing

    [Fact]
    public void LineSpacing_multiple_writes_auto_line_and_round_trips()
    {
        var file = CreateDoc(title: "Double spaced");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"lineSpacing":1.5}}]""");

        Assert.Equal(1.5, Get(file, "/body/p[1]")["lineSpacing"]!.GetValue<double>());
        var spacing = FirstParaPr(file).SpacingBetweenLines!;
        Assert.Equal("360", spacing.Line!.Value);
        Assert.Equal(LineSpacingRuleValues.Auto, spacing.LineRule!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void LineSpacing_double_writes_480()
    {
        var file = CreateDoc(title: "Report");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"lineSpacing":2}}]""");

        Assert.Equal(2, Get(file, "/body/p[1]")["lineSpacing"]!.GetValue<double>());
        Assert.Equal("480", FirstParaPr(file).SpacingBetweenLines!.Line!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void LineSpacing_atLeast_object_writes_atLeast_rule()
    {
        var file = CreateDoc(title: "At least");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"lineSpacing":{"atLeast":12}}}]""");

        var ls = Get(file, "/body/p[1]")["lineSpacing"]!;
        Assert.Equal(12, ls["atLeast"]!.GetValue<double>());
        var spacing = FirstParaPr(file).SpacingBetweenLines!;
        Assert.Equal("240", spacing.Line!.Value); // 12pt * 20
        Assert.Equal(LineSpacingRuleValues.AtLeast, spacing.LineRule!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void LineSpacing_exactly_object_writes_exact_rule()
    {
        var file = CreateDoc(title: "Exactly");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"lineSpacing":{"exactly":14}}}]""");

        var ls = Get(file, "/body/p[1]")["lineSpacing"]!;
        Assert.Equal(14, ls["exactly"]!.GetValue<double>());
        var spacing = FirstParaPr(file).SpacingBetweenLines!;
        Assert.Equal("280", spacing.Line!.Value); // 14pt * 20
        Assert.Equal(LineSpacingRuleValues.Exact, spacing.LineRule!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void LineSpacing_coexists_with_spacingBefore_on_one_spacing_element()
    {
        var file = CreateDoc(title: "Spaced report");

        // Both ride the same w:spacing — neither must drop the other.
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"spacingBefore":12,"lineSpacing":1.5}}]""");

        var props = Get(file, "/body/p[1]");
        Assert.Equal(12, props["spacingBefore"]!.GetValue<double>());
        Assert.Equal(1.5, props["lineSpacing"]!.GetValue<double>());

        var spacing = FirstParaPr(file).SpacingBetweenLines!;
        Assert.Equal("240", spacing.Before!.Value);   // spacingBefore preserved
        Assert.Equal("360", spacing.Line!.Value);      // lineSpacing 1.5
        Assert.Equal(LineSpacingRuleValues.Auto, spacing.LineRule!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void LineSpacing_garbage_object_is_invalid_args()
    {
        var file = CreateDoc(title: "Bad spacing");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"lineSpacing":{"atLeast":12,"exactly":14}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    // ----------------------------------------------------------- paragraph toggles

    [Fact]
    public void Keep_and_break_toggles_round_trip_and_clear()
    {
        var file = CreateDoc(title: "Keep together");

        Edit(file, """
            [{"op":"set","path":"/body/p[1]","props":
              {"keepNext":true,"keepLines":true,"pageBreakBefore":true,"widowControl":true}}]
            """);

        var props = Get(file, "/body/p[1]");
        Assert.True(props["keepNext"]!.GetValue<bool>());
        Assert.True(props["keepLines"]!.GetValue<bool>());
        Assert.True(props["pageBreakBefore"]!.GetValue<bool>());
        Assert.True(props["widowControl"]!.GetValue<bool>());

        var pPr = FirstParaPr(file);
        Assert.NotNull(pPr.KeepNext);
        Assert.NotNull(pPr.KeepLines);
        Assert.NotNull(pPr.PageBreakBefore);
        Assert.NotNull(pPr.WidowControl);
        AssertValidatesClean(file);

        // false removes each element (the get then reports null).
        Edit(file, """
            [{"op":"set","path":"/body/p[1]","props":
              {"keepNext":false,"keepLines":false,"pageBreakBefore":false,"widowControl":false}}]
            """);

        var cleared = FirstParaPr(file);
        Assert.Null(cleared.KeepNext);
        Assert.Null(cleared.KeepLines);
        Assert.Null(cleared.PageBreakBefore);
        Assert.Null(cleared.WidowControl);
        Assert.Null(Get(file, "/body/p[1]")["keepNext"]?.GetValue<bool?>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void OutlineLevel_round_trips_zero_based()
    {
        var file = CreateDoc(title: "Outline");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"outlineLevel":0}}]""");

        Assert.Equal(0, Get(file, "/body/p[1]")["outlineLevel"]!.GetValue<int>());
        Assert.Equal(0, FirstParaPr(file).OutlineLevel!.Val!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void OutlineLevel_out_of_range_is_invalid_args()
    {
        var file = CreateDoc(title: "Bad outline");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"outlineLevel":10}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    // ----------------------------------------------------------- tab stops

    [Fact]
    public void TabStops_5cm_decimal_dot_writes_pos_2835_and_round_trips()
    {
        var file = CreateDoc(title: "Tabbed");

        Edit(file, """
            [{"op":"set","path":"/body/p[1]","props":
              {"tabStops":[{"pos":5,"align":"decimal","leader":"dot"}]}}]
            """);

        var stops = Get(file, "/body/p[1]")["tabStops"]!.AsArray();
        Assert.Single(stops);
        Assert.Equal(5, stops[0]!["pos"]!.GetValue<double>());
        Assert.Equal("decimal", stops[0]!["align"]!.GetValue<string>());
        Assert.Equal("dot", stops[0]!["leader"]!.GetValue<string>());

        var tab = FirstParaPr(file).Tabs!.Elements<TabStop>().Single();
        Assert.Equal(2835, tab.Position!.Value); // 5cm * 567
        Assert.Equal(TabStopValues.Decimal, tab.Val!.Value);
        Assert.Equal(TabStopLeaderCharValues.Dot, tab.Leader!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void TabStops_defaults_left_align_and_no_leader()
    {
        var file = CreateDoc(title: "Simple tab");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"tabStops":[{"pos":2.5}]}}]""");

        var stops = Get(file, "/body/p[1]")["tabStops"]!.AsArray();
        Assert.Equal(2.5, stops[0]!["pos"]!.GetValue<double>());
        Assert.Equal("left", stops[0]!["align"]!.GetValue<string>());
        Assert.Equal("none", stops[0]!["leader"]!.GetValue<string>());

        var tab = FirstParaPr(file).Tabs!.Elements<TabStop>().Single();
        Assert.Null(tab.Leader); // omitted when none
        AssertValidatesClean(file);
    }

    [Fact]
    public void TabStops_multiple_round_trip_in_order()
    {
        var file = CreateDoc(title: "Multi tab");

        Edit(file, """
            [{"op":"set","path":"/body/p[1]","props":
              {"tabStops":[{"pos":3,"align":"left"},{"pos":8,"align":"right","leader":"underscore"}]}}]
            """);

        var stops = Get(file, "/body/p[1]")["tabStops"]!.AsArray();
        Assert.Equal(2, stops.Count);
        Assert.Equal(3, stops[0]!["pos"]!.GetValue<double>());
        Assert.Equal(8, stops[1]!["pos"]!.GetValue<double>());
        Assert.Equal("right", stops[1]!["align"]!.GetValue<string>());
        Assert.Equal("underscore", stops[1]!["leader"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void TabStops_empty_array_clears_the_set()
    {
        var file = CreateDoc(title: "Cleared tabs");

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"tabStops":[{"pos":5}]}}]""");
        Assert.NotNull(Get(file, "/body/p[1]")["tabStops"]);

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"tabStops":[]}}]""");
        Assert.Null(Get(file, "/body/p[1]")["tabStops"]);
        Assert.Null(FirstParaPr(file).Tabs);
        AssertValidatesClean(file);
    }

    [Fact]
    public void TabStops_bad_align_is_invalid_args()
    {
        var file = CreateDoc(title: "Bad tab");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"tabStops":[{"pos":5,"align":"sideways"}]}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.Contains("decimal", ex.Candidates!);
    }

    // ----------------------------------------------------------- composition

    [Fact]
    public void All_typography_props_compose_in_one_op_and_validate()
    {
        var file = CreateDoc(title: "Designed clause");

        Edit(file, """
            [{"op":"set","path":"/body/p[1]","props":{
                "text":"Important note",
                "highlight":"yellow","strike":true,"smallCaps":true,"superscript":true,"characterSpacing":0.5,
                "lineSpacing":2,"spacingBefore":6,"keepNext":true,"pageBreakBefore":true,
                "outlineLevel":1,
                "tabStops":[{"pos":4,"align":"decimal","leader":"dot"}]}}]
            """);

        var props = Get(file, "/body/p[1]");
        Assert.Equal("Important note", props["text"]!.GetValue<string>());
        Assert.Equal("yellow", props["highlight"]!.GetValue<string>());
        Assert.True(props["strike"]!.GetValue<bool>());
        Assert.True(props["superscript"]!.GetValue<bool>());
        Assert.Equal(0.5, props["characterSpacing"]!.GetValue<double>());
        Assert.Equal(2, props["lineSpacing"]!.GetValue<double>());
        Assert.Equal(6, props["spacingBefore"]!.GetValue<double>());
        Assert.True(props["keepNext"]!.GetValue<bool>());
        Assert.Equal(1, props["outlineLevel"]!.GetValue<int>());
        Assert.Equal(4, props["tabStops"]!.AsArray()[0]!["pos"]!.GetValue<double>());
        AssertValidatesClean(file);
    }
}
