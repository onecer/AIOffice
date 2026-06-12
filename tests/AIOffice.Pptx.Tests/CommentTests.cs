using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>M4 slide comments: classic commentAuthors + per-slide comments parts.</summary>
public sealed class CommentTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void CreateDeck(int slides = 2)
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        for (var i = 2; i <= slides; i++)
        {
            TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
                [TestEnv.Op("add", $"/slide[{i}]", type: "slide")]));
        }
    }

    private JsonObject AddComment(string path, string text, string? author = null, params (string Key, JsonNode? Value)[] extra)
    {
        var props = TestEnv.Props([("text", JsonValue.Create(text)), .. extra]);
        if (author is not null)
        {
            props["author"] = author;
        }

        return TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", path, type: "comment", props: props)]));
    }

    [Fact]
    public void AddComment_CreatesAuthorsAndSlideCommentsParts()
    {
        CreateDeck();
        var data = AddComment("/slide[2]", "Tighten this slide", author: "Dana Reviewer",
            ("x", JsonValue.Create("2cm")), ("y", JsonValue.Create("3cm")));
        Assert.Equal("/slide[2]/comment[@id=1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var presentation = doc.PresentationPart!;
            var author = presentation.CommentAuthorsPart!.CommentAuthorList!
                .Elements<P.CommentAuthor>().Single();
            Assert.Equal("Dana Reviewer", author.Name!.Value);
            Assert.Equal("DR", author.Initials!.Value);
            Assert.Equal(1u, author.LastIndex!.Value);

            var slide2 = presentation.SlideParts.Single(p => p.SlideCommentsPart is not null);
            var comment = slide2.SlideCommentsPart!.CommentList!.Elements<P.Comment>().Single();
            Assert.Equal(author.Id!.Value, comment.AuthorId!.Value);
            Assert.Equal(1u, comment.Index!.Value);
            Assert.Equal("Tighten this slide", comment.GetFirstChild<P.Text>()!.Text);
            Assert.Equal(720_000L, comment.GetFirstChild<P.Position>()!.X!.Value); // 2cm
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Authors_AreDeduplicated_CaseInsensitively()
    {
        CreateDeck();
        AddComment("/slide[1]", "First", author: "Dana");
        AddComment("/slide[2]", "Second", author: "dana");
        AddComment("/slide[2]", "Third", author: "Riley");

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var authors = doc.PresentationPart!.CommentAuthorsPart!.CommentAuthorList!
                .Elements<P.CommentAuthor>().ToList();
            Assert.Equal(2, authors.Count);
            Assert.Equal(["Dana", "Riley"], authors.Select(a => a.Name!.Value));
            Assert.Equal(2u, authors[0].LastIndex!.Value); // Dana owns comments 1 and 2
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Get_CommentPath_ReportsAuthorTextAndPosition()
    {
        CreateDeck();
        AddComment("/slide[2]", "Check the numbers", author: "Dana",
            ("x", JsonValue.Create("1cm")), ("y", JsonValue.Create("2cm")));

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[2]/comment[@id=1]"))));
        Assert.Equal("/slide[2]/comment[@id=1]", detail["path"]!.GetValue<string>());
        Assert.Equal(2, detail["slide"]!.GetValue<int>());
        Assert.Equal(1u, detail["id"]!.GetValue<uint>());
        Assert.Equal("Dana", detail["author"]!.GetValue<string>());
        Assert.Equal("Check the numbers", detail["text"]!.GetValue<string>());
        Assert.Equal(1, detail["x"]!.GetValue<double>());
        Assert.Equal(2, detail["y"]!.GetValue<double>());
        Assert.NotNull(detail["date"]);
    }

    [Fact]
    public void ReadViewComments_ListsAcrossSlides()
    {
        CreateDeck(slides: 3);
        AddComment("/slide[1]", "On the title");
        AddComment("/slide[3]", "On the close", author: "Dana");

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "comments"))));
        Assert.Equal("comments", data["view"]!.GetValue<string>());
        Assert.Equal(2, data["count"]!.GetValue<int>());

        var comments = data["comments"]!.AsArray();
        Assert.Equal("/slide[1]/comment[@id=1]", comments[0]!["path"]!.GetValue<string>());
        Assert.Equal("AIOffice", comments[0]!["author"]!.GetValue<string>()); // default author
        Assert.Equal(3, comments[1]!["slide"]!.GetValue<int>());
        Assert.Equal("On the close", comments[1]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void CommentIds_StayUniqueAcrossSlides()
    {
        CreateDeck();
        AddComment("/slide[1]", "one");
        AddComment("/slide[2]", "two");
        var third = AddComment("/slide[1]", "three");

        Assert.Equal("/slide[1]/comment[@id=3]", third["results"]![0]!["target"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void Remove_Comment_DeletesIt_AndEmptyPartIsDropped()
    {
        CreateDeck();
        AddComment("/slide[2]", "temp note");

        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("remove", "/slide[2]/comment[@id=1]")]));
        Assert.Equal("/slide[2]/comment[@id=1]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.All(doc.PresentationPart!.SlideParts, p => Assert.Null(p.SlideCommentsPart));
        }

        TestEnv.AssertFail(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[2]/comment[@id=1]"))),
            ErrorCodes.InvalidPath);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void AddComment_OnCommentPath_PointsAtTypeReply()
    {
        CreateDeck();
        AddComment("/slide[2]", "root comment");

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "add", "/slide[2]/comment[@id=1]", type: "comment",
            props: TestEnv.Props(("text", "a reply")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("reply", error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void AddComment_WithoutText_IsInvalidArgs()
    {
        CreateDeck();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", "/slide[1]", type: "comment", props: TestEnv.Props(("author", "Dana")))]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }
}
