using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;
using P15 = DocumentFormat.OpenXml.Office2013.PowerPoint;

namespace AIOffice.Pptx.Tests;

/// <summary>M5 comment replies: p15:threadingInfo/parentCm on the classic comment parts.</summary>
public sealed class CommentReplyTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void CreateWithComment()
    {
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx")));
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "add", "/slide[1]", type: "comment",
            props: TestEnv.Props(("text", "Tighten this"), ("author", "Dana"), ("x", "2cm"), ("y", "3cm")))]));
    }

    private JsonObject AddReply(string path, string text, string? author = null)
    {
        var props = TestEnv.Props(("text", text));
        if (author is not null)
        {
            props["author"] = author;
        }

        return TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("add", path, type: "reply", props: props)]));
    }

    [Fact]
    public void AddReply_ThreadsViaParentCm_AtTheParentsAnchor()
    {
        CreateWithComment();
        var data = AddReply("/slide[1]/comment[@id=1]", "Agreed", author: "Riley");
        Assert.Equal("/slide[1]/comment[@id=2]", data["results"]![0]!["target"]!.GetValue<string>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var comments = slidePart.SlideCommentsPart!.CommentList!.Elements<P.Comment>().ToList();
            Assert.Equal(2, comments.Count);

            var root = comments[0];
            var reply = comments[1];
            Assert.Empty(root.Descendants<P15.ParentCommentIdentifier>());

            var parentCm = reply.Descendants<P15.ParentCommentIdentifier>().Single();
            Assert.Equal(root.AuthorId!.Value, parentCm.AuthorId!.Value);
            Assert.Equal(1u, parentCm.Index!.Value);
            Assert.Equal(2u, reply.Index!.Value); // globally unique idx continues

            // The reply sits at its parent's anchor.
            Assert.Equal(root.GetFirstChild<P.Position>()!.X!.Value, reply.GetFirstChild<P.Position>()!.X!.Value);

            // Riley became a second author.
            Assert.Equal(2, doc.PresentationPart!.CommentAuthorsPart!.CommentAuthorList!
                .Elements<P.CommentAuthor>().Count());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ReadViewComments_NestsRepliesUnderTheirRoot()
    {
        CreateWithComment();
        AddReply("/slide[1]/comment[@id=1]", "First reply", author: "Riley");
        AddReply("/slide[1]/comment[@id=1]", "Second reply");

        var data = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "comments"))));
        Assert.Equal(1, data["count"]!.GetValue<int>()); // one top-level thread

        var root = data["comments"]![0]!;
        Assert.Equal("/slide[1]/comment[@id=1]", root["path"]!.GetValue<string>());
        var replies = root["replies"]!.AsArray();
        Assert.Equal(2, replies.Count);
        Assert.Equal("First reply", replies[0]!["text"]!.GetValue<string>());
        Assert.Equal("Riley", replies[0]!["author"]!.GetValue<string>());
        Assert.Equal("AIOffice", replies[1]!["author"]!.GetValue<string>()); // default author
        Assert.Equal("/slide[1]/comment[@id=1]", replies[0]!["parentPath"]!.GetValue<string>());
    }

    [Fact]
    public void Get_RootAndReply_ReportTheThreadBothWays()
    {
        CreateWithComment();
        AddReply("/slide[1]/comment[@id=1]", "Agreed");

        var root = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/comment[@id=1]"))));
        Assert.Null(root["parentId"]);
        Assert.Equal("Agreed", root["replies"]![0]!["text"]!.GetValue<string>());

        var reply = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/comment[@id=2]"))));
        Assert.Equal(1u, reply["parentId"]!.GetValue<uint>());
        Assert.Equal("/slide[1]/comment[@id=1]", reply["parentPath"]!.GetValue<string>());
        Assert.Null(reply["replies"]);
    }

    [Fact]
    public void RemoveReply_LeavesTheRootInPlace()
    {
        CreateWithComment();
        AddReply("/slide[1]/comment[@id=1]", "Drop me");

        var data = TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"),
            [TestEnv.Op("remove", "/slide[1]/comment[@id=2]")]));
        Assert.Equal("/slide[1]/comment[@id=2]", data["results"]![0]!["target"]!.GetValue<string>());

        var root = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/comment[@id=1]"))));
        Assert.Null(root["replies"]);
        TestEnv.AssertFail(
            _handler.Get(_ws.Ctx("deck.pptx", ("path", "/slide[1]/comment[@id=2]"))),
            ErrorCodes.InvalidPath);
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void RemoveRoot_TakesItsRepliesWithIt()
    {
        CreateWithComment();
        AddReply("/slide[1]/comment[@id=1]", "Reply 1");
        AddReply("/slide[1]/comment[@id=1]", "Reply 2");

        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op("remove", "/slide[1]/comment[@id=1]")]));

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            Assert.All(doc.PresentationPart!.SlideParts, p => Assert.Null(p.SlideCommentsPart));
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ReplyToAReply_IsInvalidArgs_PointingAtTheRoot()
    {
        CreateWithComment();
        AddReply("/slide[1]/comment[@id=1]", "First level");

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
            "add", "/slide[1]/comment[@id=2]", type: "reply", props: TestEnv.Props(("text", "Too deep")))]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("/slide[1]/comment[@id=1]", error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Reply_OnSlidePath_IsInvalidArgs()
    {
        CreateWithComment();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
                "add", "/slide[1]", type: "reply", props: TestEnv.Props(("text", "lost")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void Reply_WithoutText_IsInvalidArgs()
    {
        CreateWithComment();
        TestEnv.AssertFail(
            _handler.Edit(_ws.Ctx("deck.pptx"), [TestEnv.Op(
                "add", "/slide[1]/comment[@id=1]", type: "reply", props: TestEnv.Props(("author", "Dana")))]),
            ErrorCodes.InvalidArgs);
    }

    [Fact]
    public void ReplyIds_AreGloballyUnique_AcrossSlides()
    {
        CreateWithComment();
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[2]", type: "slide"),
            TestEnv.Op("add", "/slide[2]", type: "comment", props: TestEnv.Props(("text", "On two"))),
        ]));

        var data = AddReply("/slide[1]/comment[@id=1]", "Reply after slide-2 comment");
        Assert.Equal("/slide[1]/comment[@id=3]", data["results"]![0]!["target"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }
}
