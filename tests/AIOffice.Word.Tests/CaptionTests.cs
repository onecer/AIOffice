using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class CaptionTests : WordTestBase
{
    private static JsonNode FirstOp(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["data"]!["ops"]!.AsArray()[0]!;

    private static List<string> WarningCodes(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"] is JsonArray warnings
            ? [.. warnings.Select(w => w!["code"]!.GetValue<string>())]
            : [];

    /// <summary>A doc with a body paragraph standing in for a figure.</summary>
    private string FigureDoc(string name = "doc.docx")
    {
        var file = CreateDoc(name, title: "Captioned");
        Edit(file, """[{"op":"set","path":"/body/p[1]","props":{"text":"[figure placeholder]"}}]""");
        return file;
    }

    // --------------------------------------------------- caption SEQ + bookmark

    [Fact]
    public void Add_caption_emits_seq_field_with_bookmark_and_cached_number()
    {
        var file = FigureDoc();
        var envelope = Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"Quarterly trend","position":"after"}}]""");

        var op = FirstOp(envelope);
        Assert.Equal("/caption[@label=Figure][1]", op["path"]!.GetValue<string>());
        Assert.Equal(1, op["number"]!.GetValue<int>());
        Assert.Contains("caption_numbers_cached", WarningCodes(envelope));

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var caption = Assert.Single(
            body.Elements<Paragraph>(),
            p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "Caption");

        // The SEQ complex field instruction is present and labelled.
        var seq = caption.Descendants<FieldCode>().Select(c => c.Text).FirstOrDefault(t => t.Contains("SEQ"));
        Assert.NotNull(seq);
        Assert.Contains("SEQ Figure", seq);

        // The caption is wrapped in a _Ref bookmark.
        var bookmark = Assert.Single(caption.Descendants<BookmarkStart>());
        Assert.StartsWith("_Ref", bookmark.Name!.Value);

        // The label and the cached number both render before Word refreshes fields.
        Assert.Contains("Figure", caption.InnerText);
        Assert.Contains("Quarterly trend", caption.InnerText);

        AssertValidatesClean(file);
    }

    [Fact]
    public void Successive_same_label_captions_number_sequentially()
    {
        var file = FigureDoc();
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"One"}}]""");
        var second = Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"Two","position":"after"}}]""");

        Assert.Equal(2, FirstOp(second)["number"]!.GetValue<int>());

        // Different label gets its own counter starting at 1.
        var table = Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Table","text":"Grid","position":"after"}}]""");
        Assert.Equal(1, FirstOp(table)["number"]!.GetValue<int>());
        Assert.Equal("/caption[@label=Table][1]", FirstOp(table)["path"]!.GetValue<string>());
    }

    // ------------------------------------------------------------- get caption

    [Fact]
    public void Get_caption_returns_label_number_and_text()
    {
        var file = FigureDoc();
        Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"Quarterly trend","position":"after"}}]""");

        var data = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/caption[@label=Figure][1]" })));
        Assert.Equal("caption", data["type"]!.GetValue<string>());
        var props = data["properties"]!;
        Assert.Equal("Figure", props["label"]!.GetValue<string>());
        Assert.Equal(1, props["number"]!.GetValue<int>());
        Assert.Equal("Quarterly trend", props["text"]!.GetValue<string>());
        Assert.StartsWith("_Ref", props["bookmark"]!.GetValue<string>());
    }

    [Fact]
    public void Get_missing_caption_is_invalid_path_with_candidates()
    {
        var file = FigureDoc();
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"Only one"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/caption[@label=Figure][9]" })));
        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/caption[@label=Figure][1]", ex.Candidates!);
    }

    // -------------------------------------------------- structure lists captions

    [Fact]
    public void Structure_view_lists_captions()
    {
        var file = FigureDoc();
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"Trend"}}]""");

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));
        var captions = data["captions"]!.AsArray();
        var caption = Assert.Single(captions);
        Assert.Equal("/caption[@label=Figure][1]", caption!["path"]!.GetValue<string>());
        Assert.Equal("Figure", caption["label"]!.GetValue<string>());
        Assert.Equal("Trend", caption["text"]!.GetValue<string>());
    }

    // ------------------------------------------------------ crossRef REF field

    [Fact]
    public void Add_crossRef_emits_ref_field_with_cached_text_and_resolves_target()
    {
        var file = FigureDoc();
        Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"Trend","position":"after"}}]""");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"As shown in "}}]""");

        var envelope = Edit(file,
            """[{"op":"add","path":"/body/p[3]","type":"crossRef","props":{"to":"/caption[@label=Figure][1]","show":"labelAndNumber"}}]""");

        var op = FirstOp(envelope);
        Assert.Equal("/caption[@label=Figure][1]", op["to"]!.GetValue<string>());
        Assert.Equal("Figure 1", op["cached"]!.GetValue<string>());
        Assert.StartsWith("_Ref", op["target"]!.GetValue<string>());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var refField = doc.MainDocumentPart!.Document!.Descendants<FieldCode>()
            .Select(c => c.Text)
            .FirstOrDefault(t => t.Contains("REF") && !t.Contains("SEQ"));
        Assert.NotNull(refField);
        Assert.Contains(op["target"]!.GetValue<string>(), refField);

        AssertValidatesClean(file);
    }

    [Fact]
    public void CrossRef_numberOnly_caches_just_the_number()
    {
        var file = FigureDoc();
        Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"Trend","position":"after"}}]""");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"See "}}]""");

        var envelope = Edit(file,
            """[{"op":"add","path":"/body/p[3]","type":"crossRef","props":{"to":"Figure 1","show":"numberOnly"}}]""");
        Assert.Equal("1", FirstOp(envelope)["cached"]!.GetValue<string>());
    }

    [Fact]
    public void Get_crossRef_reports_target_and_show()
    {
        var file = FigureDoc();
        Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"Trend","position":"after"}}]""");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Ref: "}}]""");
        Edit(file,
            """[{"op":"add","path":"/body/p[3]","type":"crossRef","props":{"to":"/caption[@label=Figure][1]","show":"labelAndNumber"}}]""");

        var data = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/crossRef[1]" })));
        Assert.Equal("crossRef", data["type"]!.GetValue<string>());
        var props = data["properties"]!;
        Assert.Equal("/caption[@label=Figure][1]", props["to"]!.GetValue<string>());
        Assert.Equal("labelAndNumber", props["show"]!.GetValue<string>());
        Assert.Equal("Figure 1", props["cached"]!.GetValue<string>());
    }

    [Fact]
    public void CrossRef_to_a_missing_caption_is_invalid_path()
    {
        var file = FigureDoc();
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"See "}}]""");

        var ex = Assert.Throws<AiofficeException>(() => Edit(file,
            """[{"op":"add","path":"/body/p[2]","type":"crossRef","props":{"to":"/caption[@label=Figure][1]"}}]"""));
        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
    }

    // --------------------------------------------------------------- guards

    [Fact]
    public void Add_caption_with_unknown_label_is_invalid_args()
    {
        var file = FigureDoc();
        var ex = Assert.Throws<AiofficeException>(() => Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Diagram","text":"x"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("Figure", ex.Candidates!);
    }

    [Fact]
    public void Tracked_caption_is_unsupported_feature()
    {
        var file = FigureDoc();
        var ex = Assert.Throws<AiofficeException>(() => Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"x"}}]""",
            new JsonObject { ["track"] = true }));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    // ------------------------------------------------------------ round-trip

    [Fact]
    public void Caption_and_crossRef_survive_reopen_byte_for_field_structure()
    {
        var file = FigureDoc();
        Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"caption","props":{"label":"Figure","text":"Trend","position":"after"}}]""");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"As in "}}]""");
        Edit(file,
            """[{"op":"add","path":"/body/p[3]","type":"crossRef","props":{"to":"/caption[@label=Figure][1]"}}]""");

        // Re-read the caption + crossRef from a fresh open: numbers and targets intact.
        var caption = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/caption[@label=Figure][1]" })))["properties"]!;
        Assert.Equal(1, caption["number"]!.GetValue<int>());

        var crossRef = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/crossRef[1]" })))["properties"]!;
        Assert.Equal("/caption[@label=Figure][1]", crossRef["to"]!.GetValue<string>());

        AssertValidatesClean(file);
    }
}
