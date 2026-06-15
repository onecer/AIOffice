using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// (1.7) Drop caps: set props.dropCap on a body paragraph lifts its first letter
/// into a framed paragraph (w:framePr w:dropCap). These verify the structure
/// reopens, the validator stays clean, get reads it back, and it can be cleared.
/// </summary>
public sealed class DropCapTests : WordTestBase
{
    /// <summary>Creates a doc whose body paragraph p[2] holds the sentence to drop-cap.</summary>
    private string CreateWithParagraph(string text = "Once upon a time there was a document.")
    {
        var file = CreateDoc(title: "DropCap");
        var ops = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "set",
                ["path"] = "/body/p[2]",
                ["props"] = new JsonObject { ["text"] = text },
            },
        };
        Edit(file, ops.ToJsonString());
        return file;
    }

    /// <summary>Sets dropCap props on a path via a structured op (avoids brittle JSON brace escaping).</summary>
    private void SetDropCap(string file, string path, JsonObject props)
    {
        var ops = new JsonArray
        {
            new JsonObject { ["op"] = "set", ["path"] = path, ["props"] = props },
        };
        Edit(file, ops.ToJsonString());
    }

    private static Paragraph FramedDropCapParagraph(WordprocessingDocument doc) =>
        doc.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>()
            .Single(p => p.ParagraphProperties?.FrameProperties?.DropCap is not null);

    [Fact]
    public void Drop_cap_lifts_the_first_letter_into_a_framed_paragraph()
    {
        var file = CreateWithParagraph();

        var envelope = Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"dropCap":"drop","dropCapLines":3}}]""");
        var summary = Data(envelope)["ops"]!.AsArray()[0]!["dropCap"]!;
        Assert.Equal("drop", summary["location"]!.GetValue<string>());
        Assert.Equal(3, summary["lines"]!.GetValue<int>());
        Assert.Equal("O", summary["letter"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var framed = FramedDropCapParagraph(doc);
            Assert.Equal(DropCapLocationValues.Drop, framed.ParagraphProperties!.FrameProperties!.DropCap!.Value);
            Assert.Equal(3, framed.ParagraphProperties.FrameProperties.Lines!.Value);
            Assert.Equal("O", framed.InnerText);

            // The body paragraph lost its leading "O" — the rest of the sentence stays.
            var body = doc.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>()
                .First(p => p.InnerText.StartsWith("nce upon", StringComparison.Ordinal));
            Assert.StartsWith("nce upon a time", body.InnerText, StringComparison.Ordinal);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Margin_drop_cap_anchors_to_the_margin()
    {
        var file = CreateWithParagraph();
        Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"dropCap":"margin"}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var framePr = FramedDropCapParagraph(doc).ParagraphProperties!.FrameProperties!;
        Assert.Equal(DropCapLocationValues.Margin, framePr.DropCap!.Value);
        Assert.Equal(3, framePr.Lines!.Value); // default lines
        AssertValidatesClean(file);
    }

    [Fact]
    public void Drop_cap_font_lands_on_the_lifted_letter()
    {
        var file = CreateWithParagraph();
        Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"dropCap":"drop","dropCapFont":"Georgia"}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var fonts = FramedDropCapParagraph(doc).Descendants<RunFonts>().Single();
        Assert.Equal("Georgia", fonts.Ascii!.Value);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_reports_the_drop_cap_on_the_body_paragraph()
    {
        var file = CreateWithParagraph();
        Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"dropCap":"drop","dropCapLines":4,"dropCapFont":"Georgia"}}]""");

        // The body paragraph is now p[3] (the framed letter took p[2]); query finds it by text.
        var bodyPath = BodyParagraphPathStartingWith(file, "nce upon");
        var props = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = bodyPath })))["properties"]!;
        var dropCap = props["dropCap"]!;
        Assert.Equal("drop", dropCap["location"]!.GetValue<string>());
        Assert.Equal(4, dropCap["lines"]!.GetValue<int>());
        Assert.Equal("O", dropCap["letter"]!.GetValue<string>());
        Assert.Equal("Georgia", dropCap["font"]!.GetValue<string>());
    }

    [Fact]
    public void Drop_cap_none_rejoins_the_letter()
    {
        var file = CreateWithParagraph();
        Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"dropCap":"drop"}}]""");

        // Clear it: the framed paragraph disappears and the letter rejoins the body.
        var bodyPath = BodyParagraphPathStartingWith(file, "nce upon");
        SetDropCap(file, bodyPath, new JsonObject { ["dropCap"] = "none" });

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<FrameProperties>());
        Assert.Contains(
            doc.MainDocumentPart.Document.Body!.Descendants<Paragraph>(),
            p => p.InnerText.StartsWith("Once upon a time", StringComparison.Ordinal));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Reapplying_a_drop_cap_does_not_stack_a_second_one()
    {
        var file = CreateWithParagraph();
        Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"dropCap":"drop"}}]""");
        var bodyPath = BodyParagraphPathStartingWith(file, "nce upon");
        SetDropCap(file, bodyPath, new JsonObject { ["dropCap"] = "drop", ["dropCapLines"] = 5 });

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var framed = doc.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>()
            .Where(p => p.ParagraphProperties?.FrameProperties?.DropCap is not null)
            .ToList();
        Assert.Single(framed); // exactly one framed drop-cap paragraph, not two
        Assert.Equal(5, framed[0].ParagraphProperties!.FrameProperties!.Lines!.Value);
        Assert.Equal("O", framed[0].InnerText); // still a single letter, no duplicate
        AssertValidatesClean(file);
    }

    [Fact]
    public void Invalid_drop_cap_location_is_invalid_args()
    {
        var file = CreateWithParagraph();
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/p[2]","props":{"dropCap":"sideways"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("margin", ex.Candidates!);
    }

    [Fact]
    public void Drop_cap_on_a_non_paragraph_is_unsupported()
    {
        var file = CreateDoc(title: "Tbl");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":1}}]""");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/body/table[1]","props":{"dropCap":"drop"}}]"""));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    [Fact]
    public void Drop_cap_can_ride_alongside_a_text_change_in_one_op()
    {
        var file = CreateDoc(title: "Combined");
        Edit(file, """
            [{"op":"set","path":"/body/p[1]","props":{"text":"Beautiful prose here.","dropCap":"drop"}}]
            """);

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var framed = FramedDropCapParagraph(doc);
        Assert.Equal("B", framed.InnerText); // first letter of the just-set text
        AssertValidatesClean(file);
    }

    /// <summary>Finds the body paragraph path whose text starts with the given prefix (via query).</summary>
    private string BodyParagraphPathStartingWith(string file, string prefix)
    {
        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "text" })));
        var lines = data["lines"]!.AsArray();
        var line = lines.First(l => l!["text"]!.GetValue<string>().StartsWith(prefix, StringComparison.Ordinal));
        return line!["path"]!.GetValue<string>();
    }
}
