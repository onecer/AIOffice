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
    public void Reply_props_on_a_comment_add_point_to_the_reply_op()
    {
        var file = CreateDoc(title: "Thread");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"root"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"reply","replyTo":1}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("\"type\":\"reply\"", ex.Suggestion, StringComparison.Ordinal);
        Assert.Contains("/comment[@id=1]", ex.Suggestion, StringComparison.Ordinal);
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

    // ----------------------------------------------------------------- replies

    [Fact]
    public void Reply_threads_under_its_parent_in_the_comments_view()
    {
        var file = CreateDoc(title: "Discussion");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"root","author":"Alice"}}]""");

        var envelope = Edit(file, """[{"op":"add","path":"/comment[@id=1]","type":"reply","props":{"text":"agreed","author":"Bob"}}]""");
        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("/comment[@id=2]", summary["path"]!.GetValue<string>());
        Assert.Equal("/comment[@id=1]", summary["parent"]!.GetValue<string>());

        var comments = Comments(file);
        var root = Assert.Single(comments)!; // the reply nests instead of listing top-level
        Assert.Equal("root", root["text"]!.GetValue<string>());
        var reply = Assert.Single(root["replies"]!.AsArray())!;
        Assert.Equal("agreed", reply["text"]!.GetValue<string>());
        Assert.Equal("Bob", reply["author"]!.GetValue<string>());
        Assert.Equal(1, reply["parentId"]!.GetValue<int>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Reply_wires_w15_parent_paraId_and_survives_reopen()
    {
        var file = CreateDoc(title: "Wired");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"root"}}]""");
        Edit(file, """[{"op":"add","path":"/comment[@id=1]","type":"reply","props":{"text":"child"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var comments = doc.MainDocumentPart!.WordprocessingCommentsPart!.Comments!;
            var byId = comments.Elements<Comment>().ToDictionary(c => c.Id!.Value!);
            var rootParaId = byId["1"].Elements<Paragraph>().Last().ParagraphId!.Value;
            var replyParaId = byId["2"].Elements<Paragraph>().Last().ParagraphId!.Value;
            Assert.NotEqual(rootParaId, replyParaId);

            var entry = Assert.Single(
                doc.MainDocumentPart.WordprocessingCommentsExPart!.CommentsEx!
                    .Elements<DocumentFormat.OpenXml.Office2013.Word.CommentEx>());
            Assert.Equal(replyParaId, entry.ParaId!.Value);
            Assert.Equal(rootParaId, entry.ParaIdParent!.Value);

            // The reply anchors on the same content as its parent.
            var body = doc.MainDocumentPart.Document!.Body!;
            Assert.Equal(2, body.Descendants<CommentRangeStart>().Count());
            Assert.Equal(2, body.Descendants<CommentReference>().Count());
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_reply_reports_parent_and_removing_the_parent_cascades()
    {
        var file = CreateDoc(title: "Cascade");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"root"}}]""");
        Edit(file, """[{"op":"add","path":"/comment[@id=1]","type":"reply","props":{"text":"child"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/comment[@id=2]" })));
        Assert.Equal(1, got["properties"]!["parentId"]!.GetValue<int>());

        Edit(file, """[{"op":"remove","path":"/comment[@id=1]"}]""");

        Assert.Empty(Comments(file));
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<CommentRangeStart>());
            Assert.Empty(body.Descendants<CommentReference>());
            Assert.Empty(
                doc.MainDocumentPart.WordprocessingCommentsExPart?.CommentsEx?
                    .Elements<DocumentFormat.OpenXml.Office2013.Word.CommentEx>() ?? []);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Reply_to_missing_comment_is_invalid_path()
    {
        var file = CreateDoc(title: "Orphan");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/comment[@id=7]","type":"reply","props":{"text":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
    }

    [Fact]
    public void Reply_text_is_required()
    {
        var file = CreateDoc(title: "Quiet");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"comment","props":{"text":"root"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/comment[@id=1]","type":"reply","props":{"author":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("text", ex.Message, StringComparison.Ordinal);
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
