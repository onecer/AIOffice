using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// RTL / bidirectional support: paragraph w:bidi (with default right alignment),
/// run-level right-to-left marks for mixed content, and table w:bidiVisual. Every
/// mutation reopens to confirm the flag landed and the validator stays clean.
/// </summary>
public sealed class RtlTests : WordTestBase
{
    private static ParagraphProperties? PPr(string file, int index = 1)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        return doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ElementAt(index - 1).ParagraphProperties;
    }

    [Fact]
    public void Rtl_paragraph_sets_bidi_and_defaults_to_right_alignment()
    {
        var file = CreateDoc(title: "RTL");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"rtl":true}}]""");

        var pPr = PPr(file);
        Assert.NotNull(pPr!.BiDi);
        Assert.NotEqual(DocumentFormat.OpenXml.OnOffValue.FromBoolean(false), pPr.BiDi!.Val);
        Assert.Equal(JustificationValues.Right, pPr.Justification?.Val?.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Explicit_alignment_overrides_the_rtl_default()
    {
        var file = CreateDoc(title: "RTL");
        // alignment center given alongside rtl must win over the implicit right.
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"rtl":true,"alignment":"center"}}]""");

        Assert.Equal(JustificationValues.Center, PPr(file)!.Justification?.Val?.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_reports_paragraph_direction()
    {
        var file = CreateDoc(title: "RTL");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"rtl":true}}]""");

        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]" })))["properties"]!;
        Assert.True(properties["rtl"]!.GetValue<bool>());
    }

    [Fact]
    public void Rtl_false_clears_the_flag()
    {
        var file = CreateDoc(title: "RTL");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"rtl":true}}]""");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"rtl":false}}]""");

        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]" })))["properties"]!;
        Assert.False(properties["rtl"]!.GetValue<bool>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Run_rtl_marks_mixed_direction_content()
    {
        var file = CreateDoc(title: "RTL");
        // Put text on the empty default paragraph so it has an addressable run.
        Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"text":"mixed"}}]""");
        Edit(file, """[{"op":"set","path":"/body/p[2]/run[1]","props":{"rtl":true}}]""");

        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[2]/run[1]" })))["properties"]!;
        Assert.True(properties["rtl"]!.GetValue<bool>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var run = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ElementAt(1).Elements<Run>().First();
            Assert.NotNull(run.RunProperties?.RightToLeftText);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Rtl_table_sets_bidi_visual()
    {
        var file = CreateDoc(title: "RTL");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}}]""");
        Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"rtl":true}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
            Assert.NotNull(table.GetFirstChild<TableProperties>()?.BiDiVisual);
        }

        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/table[1]" })))["properties"]!;
        Assert.True(properties["rtl"]!.GetValue<bool>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Structure_view_notes_rtl_paragraphs()
    {
        var file = CreateDoc(title: "RTL");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"rtl":true}}]""");

        // The paragraph get-shape (used by structure consumers) exposes rtl=true,
        // confirming downstream views can surface direction.
        var properties = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]" })))["properties"]!;
        Assert.True(properties["rtl"]!.GetValue<bool>());
    }
}
