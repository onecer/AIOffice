using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class CommentTests : WordTestBase
{
    private JsonArray Comments(string file) =>
        Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "comments" })))["comments"]!.AsArray();

    [Fact]
    public void Add_comment_on_paragraph_reads_back_with_anchor()
    {
        var file = CreateDoc(title: "Quarterly report");

        var envelope = Edit(
            file,
            """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"Please verify the numbers."}}]""");

        Assert.Equal("/comment[@id=1]", Data(envelope)["ops"]!.AsArray()[0]!["path"]!.GetValue<string>());

        var comment = Assert.Single(Comments(file))!;
        Assert.Equal(1, comment["id"]!.GetValue<int>());
        Assert.Equal("AIOffice", comment["author"]!.GetValue<string>());
        Assert.Equal("Please verify the numbers.", comment["text"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", comment["anchorPath"]!.GetValue<string>());
        Assert.Equal("Quarterly report", comment["anchorText"]!.GetValue<string>());
        Assert.False(string.IsNullOrEmpty(comment["date"]!.GetValue<string>()));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_comment_on_run_anchors_just_that_run()
    {
        var file = CreateDoc(title: "Alpha");

        Edit(file, """[{"op":"add","path":"/body/p[1]/run[1]","type":"comment","props":{"text":"run note","author":"Reviewer"}}]""");

        var comment = Assert.Single(Comments(file))!;
        Assert.Equal("Reviewer", comment["author"]!.GetValue<string>());
        Assert.Equal("Alpha", comment["anchorText"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", comment["anchorPath"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Comment_ids_increment_and_get_resolves_them()
    {
        var file = CreateDoc(title: "Two notes");

        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"first"}},
              {"op":"add","path":"/body/p[2]","type":"comment","props":{"text":"second"}}
            ]
            """);

        Assert.Equal(2, Comments(file).Count);
        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/comment[@id=2]" })));
        Assert.Equal("comment", got["type"]!.GetValue<string>());
        Assert.Equal("second", got["properties"]!["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_comment_clears_part_entry_and_anchor_markers()
    {
        var file = CreateDoc(title: "Disposable");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"temp"}}]""");

        Edit(file, """[{"op":"remove","path":"/comment[@id=1]"}]""");

        Assert.Empty(Comments(file));
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<CommentRangeStart>());
            Assert.Empty(body.Descendants<CommentRangeEnd>());
            Assert.Empty(body.Descendants<CommentReference>());
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Comment_text_is_required()
    {
        var file = CreateDoc(title: "Empty");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"author":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("text", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Comment_on_non_content_path_is_invalid_args()
    {
        var file = CreateDoc(title: "Anchors");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":1}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/table[1]/tr[1]/tc[1]","type":"comment","props":{"text":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("/p[1]", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Comment_replies_are_unsupported_until_m3()
    {
        var file = CreateDoc(title: "Thread");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"root"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"reply","replyTo":1}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("M3", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_comment_id_is_invalid_path_with_candidates()
    {
        var file = CreateDoc(title: "Lost");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"only"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/comment[@id=42]" })));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/comment[@id=1]", ex.Candidates!);
    }

    [Fact]
    public void Comment_survives_reopen_and_other_edits()
    {
        var file = CreateDoc(title: "Durable");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"sticky"}}]""");

        Edit(file, """[{"op":"add","path":"/body","props":{"text":"More content"}}]""");

        var comment = Assert.Single(Comments(file))!;
        Assert.Equal("sticky", comment["text"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", comment["anchorPath"]!.GetValue<string>());
        AssertValidatesClean(file);
    }
}
