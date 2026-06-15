using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class IndexTests : WordTestBase
{
    private static JsonNode FirstOp(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["data"]!["ops"]!.AsArray()[0]!;

    private static List<string> WarningCodes(Envelope envelope) =>
        JsonNode.Parse(envelope.ToJson())!["meta"]!["warnings"] is JsonArray warnings
            ? [.. warnings.Select(w => w!["code"]!.GetValue<string>())]
            : [];

    /// <summary>A three-paragraph doc to mark index entries on.</summary>
    private string ProseDoc()
    {
        var file = CreateDoc(title: "Indexed");
        Edit(file, """
            [
              {"op":"set","path":"/body/p[1]","props":{"text":"Machine learning powers modern AI systems."}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Neural networks and agents collaborate."}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Backpropagation trains the network."}}
            ]
            """);
        return file;
    }

    [Fact]
    public void Add_index_entry_marks_an_xe_field_on_the_paragraph()
    {
        var file = ProseDoc();

        var envelope = Edit(file,
            """[{"op":"add","path":"/body/p[1]","type":"indexEntry","props":{"text":"AI"}}]""");

        var op = FirstOp(envelope);
        Assert.Equal("/body/p[1]", op["path"]!.GetValue<string>());
        Assert.Equal("AI", op["text"]!.GetValue<string>());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var p = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        var xe = p.Descendants<FieldCode>().Select(c => c.Text).FirstOrDefault(t => t.Contains("XE"));
        Assert.NotNull(xe);
        Assert.Contains("XE \"AI\"", xe, StringComparison.Ordinal);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Sub_entry_encodes_as_main_colon_sub()
    {
        var file = ProseDoc();

        Edit(file, """[{"op":"add","path":"/body/p[2]","type":"indexEntry","props":{"text":"AI","subEntry":"agents"}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var xe = doc.MainDocumentPart!.Document!.Descendants<FieldCode>().Select(c => c.Text)
            .First(t => t.Contains("XE"));
        Assert.Contains("XE \"AI:agents\"", xe, StringComparison.Ordinal);
    }

    [Fact]
    public void Find_places_the_marker_before_the_matched_text()
    {
        var file = ProseDoc();

        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"indexEntry","props":{"text":"learning","find":"learning"}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var p = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
        // The XE begin field char appears before the run holding "learning".
        var children = p.ChildElements.ToList();
        var beginIndex = children.FindIndex(c =>
            c is Run r && r.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Begin);
        var learningIndex = children.FindIndex(c => c.InnerText.Contains("learning", StringComparison.Ordinal));
        Assert.True(beginIndex >= 0 && beginIndex < learningIndex);
    }

    [Fact]
    public void Add_index_builds_field_plus_alphabetized_entries_with_warning()
    {
        var file = ProseDoc();
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"indexEntry","props":{"text":"Neural networks"}},
              {"op":"add","path":"/body/p[1]","type":"indexEntry","props":{"text":"AI"}},
              {"op":"add","path":"/body/p[2]","type":"indexEntry","props":{"text":"AI","subEntry":"agents"}},
              {"op":"add","path":"/body/p[3]","type":"indexEntry","props":{"text":"Backpropagation"}}
            ]
            """);

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"index","props":{"columns":2}}]""");

        var op = FirstOp(envelope);
        Assert.Equal("/index[1]", op["path"]!.GetValue<string>());
        Assert.Equal(2, op["columns"]!.GetValue<int>());
        // Four distinct entries: AI, AI:agents, Backpropagation, Neural networks.
        Assert.Equal(4, op["entries"]!.GetValue<int>());
        Assert.Contains("index_cached", WarningCodes(envelope));

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var sdt = Assert.Single(body.Elements<SdtBlock>(), s =>
            s.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value
                == "Index");

        // The INDEX field with the column switch.
        var instruction = Assert.Single(sdt.Descendants<FieldCode>(), c => c.Text.Contains("INDEX"));
        Assert.Contains("INDEX \\c \"2\"", instruction.Text, StringComparison.Ordinal);

        // Alphabetized pre-rendered entries; page numbers cached as "?". The
        // entry display text is the "term, ?" run (the first paragraph also hosts
        // the field machinery, so read each entry's last text run).
        var entries = sdt.Descendants<Paragraph>()
            .Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value is "Index1" or "Index2")
            .Select(p => p.Descendants<Text>().Select(t => t.Text).LastOrDefault(t => t.EndsWith(", ?", StringComparison.Ordinal)))
            .OfType<string>()
            .ToList();
        Assert.Equal(4, entries.Count);
        // Alphabetized: AI, agents (sub of AI), Backpropagation, Neural networks.
        Assert.Equal(new[] { "AI, ?", "agents, ?", "Backpropagation, ?", "Neural networks, ?" }, entries);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_index_reports_columns_and_counts()
    {
        var file = ProseDoc();
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"indexEntry","props":{"text":"AI"}},
              {"op":"add","path":"/body/p[2]","type":"indexEntry","props":{"text":"agents"}},
              {"op":"add","path":"/body","type":"index","props":{"columns":3}}
            ]
            """);

        var data = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/index[1]" })));

        Assert.Equal("/index[1]", data["path"]!.GetValue<string>());
        Assert.Equal("index", data["type"]!.GetValue<string>());
        Assert.Equal(3, data["properties"]!["columns"]!.GetValue<int>());
        Assert.Equal(2, data["properties"]!["entryCount"]!.GetValue<int>());
        Assert.Equal(2, data["properties"]!["markedEntries"]!.GetValue<int>());
    }

    [Fact]
    public void Structure_view_lists_index_with_entry_count()
    {
        var file = ProseDoc();
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"indexEntry","props":{"text":"AI"}},
              {"op":"add","path":"/body","type":"index"}
            ]
            """);

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));
        var indexes = data["indexes"]!.AsArray();

        Assert.Single(indexes);
        Assert.Equal("/index[1]", indexes[0]!["path"]!.GetValue<string>());
        Assert.Equal(1, indexes[0]!["properties"]!["entryCount"]!.GetValue<int>());
        Assert.Equal(1, indexes[0]!["properties"]!["markedEntries"]!.GetValue<int>());
    }

    [Fact]
    public void Rerunning_index_refreshes_in_place()
    {
        var file = ProseDoc();
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"indexEntry","props":{"text":"AI"}},
              {"op":"add","path":"/body","type":"index"}
            ]
            """);

        Edit(file, """[{"op":"add","path":"/body/p[2]","type":"indexEntry","props":{"text":"networks"}}]""");
        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"index"}]""");

        Assert.True(FirstOp(envelope)["refreshed"]!.GetValue<bool>());
        Assert.Equal(2, FirstOp(envelope)["entries"]!.GetValue<int>());

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Single(doc.MainDocumentPart!.Document!.Body!.Elements<SdtBlock>(), s =>
            s.SdtProperties?.GetFirstChild<SdtContentDocPartObject>()?.GetFirstChild<DocPartGallery>()?.Val?.Value
                == "Index");
        AssertValidatesClean(file);
    }

    [Fact]
    public void Missing_text_is_invalid_args()
    {
        var file = ProseDoc();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"indexEntry","props":{}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }
}
