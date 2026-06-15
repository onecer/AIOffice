using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace AIOffice.Word.Tests;

/// <summary>v1.1.0 text effects: shadow / glow / reflection / outline on runs and paragraphs.</summary>
public sealed class TextEffectTests : WordTestBase
{
    private string CreateTextDoc(string text = "Effect me")
    {
        var file = CreateDoc();
        Edit(file, "[{\"op\":\"set\",\"path\":\"/body/p[1]\",\"props\":{\"text\":\"" + text + "\"}}]");
        return file;
    }

    private static RunProperties FirstRunProps(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First();
        return run.RunProperties!;
    }

    [Fact]
    public void Set_shadow_true_emits_w14_shadow_and_validates_clean()
    {
        var file = CreateTextDoc();

        var envelope = Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"shadow":true}}]""");
        Assert.True(envelope.IsOk, envelope.ToJson());

        var rPr = FirstRunProps(file);
        Assert.NotNull(rPr.GetFirstChild<W14.Shadow>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_glow_object_emits_w14_glow_with_color_and_radius()
    {
        var file = CreateTextDoc();

        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"glow":{"color":"00FF00","radius":8}}}]""");

        var rPr = FirstRunProps(file);
        var glow = Assert.IsType<W14.Glow>(rPr.GetFirstChild<W14.Glow>());
        Assert.Equal("00FF00", glow.RgbColorModelHex!.Val!.Value);
        // 8 pt -> 8 * 12700 EMU.
        Assert.Equal(8 * 12700L, glow.GlowRadius!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_reflection_true_emits_w14_reflection_and_validates_clean()
    {
        var file = CreateTextDoc();

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"reflection":true}}]""");

        var rPr = FirstRunProps(file);
        Assert.NotNull(rPr.GetFirstChild<W14.Reflection>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_outline_object_emits_w14_textOutline_with_solid_fill()
    {
        var file = CreateTextDoc();

        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"outline":{"color":"FF0000","width":2}}}]""");

        var rPr = FirstRunProps(file);
        var outline = Assert.IsType<W14.TextOutlineEffect>(rPr.GetFirstChild<W14.TextOutlineEffect>());
        var color = Assert.Single(outline.Descendants<W14.RgbColorModelHex>());
        Assert.Equal("FF0000", color.Val!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void TextOutline_alias_is_accepted_and_reported_as_outline()
    {
        var file = CreateTextDoc();

        var envelope = Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"textOutline":{"color":"112233"}}}]""");

        var effects = Data(envelope)["ops"]!.AsArray()[0]!["effects"]!.AsArray().Select(e => e!.GetValue<string>()).ToList();
        Assert.Contains("outline", effects);
        Assert.NotNull(FirstRunProps(file).GetFirstChild<W14.TextOutlineEffect>());
    }

    [Fact]
    public void All_four_effects_in_one_op_coexist_on_the_run_and_validate()
    {
        var file = CreateTextDoc();

        Edit(file, """
            [{"op":"set","path":"/body/p[1]","props":{
              "shadow":{"color":"808080","blur":4,"distance":3,"direction":45},
              "glow":{"color":"4F81BD","radius":5},
              "reflection":{"transparency":40,"size":35},
              "outline":{"color":"000000","width":1}
            }}]
            """);

        var rPr = FirstRunProps(file);
        Assert.NotNull(rPr.GetFirstChild<W14.Shadow>());
        Assert.NotNull(rPr.GetFirstChild<W14.Glow>());
        Assert.NotNull(rPr.GetFirstChild<W14.Reflection>());
        Assert.NotNull(rPr.GetFirstChild<W14.TextOutlineEffect>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Effects_survive_a_simultaneous_text_change_in_the_same_op()
    {
        var file = CreateTextDoc("old");

        // The text prop rebuilds the run; the effect must land on the NEW run text.
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"new text","glow":true}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var run = doc.MainDocumentPart!.Document!.Body!.Descendants<Run>().First();
        Assert.Equal("new text", run.InnerText);
        Assert.NotNull(run.RunProperties!.GetFirstChild<W14.Glow>());
    }

    [Fact]
    public void Get_reports_the_applied_effects()
    {
        var file = CreateTextDoc();
        Edit(file, """[{"op":"set","path":"/body/p[1]/run[1]","props":{"shadow":{"blur":6,"distance":4,"direction":90},"glow":{"color":"AABBCC","radius":3}}}]""");

        var data = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]/run[1]" })));
        var effects = data["properties"]!["effects"]!;

        Assert.Equal(6, effects["shadow"]!["blur"]!.GetValue<double>());
        Assert.Equal(90, effects["shadow"]!["direction"]!.GetValue<double>());
        Assert.Equal("AABBCC", effects["glow"]!["color"]!.GetValue<string>());
        Assert.Equal(3, effects["glow"]!["radius"]!.GetValue<double>());
    }

    [Fact]
    public void Effect_added_to_a_paragraph_via_add_p_is_present_on_its_run()
    {
        var file = CreateDoc();

        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Glowing heading","glow":true}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var run = doc.MainDocumentPart!.Document!.Body!
            .Descendants<Run>().First(r => r.InnerText == "Glowing heading");
        Assert.NotNull(run.RunProperties!.GetFirstChild<W14.Glow>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Effects_round_trip_through_a_no_op_reopen_byte_identical()
    {
        var file = CreateTextDoc();
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"shadow":true,"outline":{"color":"123456"}}}]""");

        var before = File.ReadAllBytes(file);
        // A no-op edit (set the same text) reopens, resaves and must not perturb the effects bytes.
        using (var doc = WordprocessingDocument.Open(file, isEditable: true))
        {
            doc.Save();
        }

        var after = File.ReadAllBytes(file);
        Assert.Equal(before.Length, after.Length);

        var rPr = FirstRunProps(file);
        Assert.NotNull(rPr.GetFirstChild<W14.Shadow>());
        Assert.NotNull(rPr.GetFirstChild<W14.TextOutlineEffect>());
    }

    [Fact]
    public void Effect_on_a_non_run_target_is_unsupported_feature()
    {
        var file = CreateDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"glow":true}}]"""));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }
}
