using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class BookmarkTests : WordTestBase
{
    [Fact]
    public void Add_bookmark_wraps_the_paragraph_with_matching_ids()
    {
        var file = CreateDoc(title: "Findings");

        var envelope = Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Results"}}]""");
        Assert.Equal("/bookmark[@name=Results]", Data(envelope)["ops"]!.AsArray()[0]!["path"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
            var start = Assert.Single(paragraph.Elements<BookmarkStart>());
            var end = Assert.Single(paragraph.Elements<BookmarkEnd>());
            Assert.Equal("Results", start.Name!.Value);
            Assert.Equal(start.Id!.Value, end.Id!.Value);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Bookmark_ids_are_unique_across_the_document()
    {
        var file = CreateDoc(title: "Two");
        Edit(file, """[{"op":"add","path":"/body","type":"p","props":{"text":"Other"}}]""");

        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"First"}},
              {"op":"add","path":"/body/p[3]","type":"bookmark","props":{"name":"Second"}}
            ]
            """);

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var ids = doc.MainDocumentPart!.Document!.Body!.Descendants<BookmarkStart>()
            .Select(b => b.Id!.Value)
            .ToList();
        Assert.Equal(2, ids.Count);
        Assert.Equal(ids.Distinct().Count(), ids.Count);
    }

    [Fact]
    public void Get_by_name_reports_anchor_path_and_snippet()
    {
        var file = CreateDoc(title: "The quick brown fox");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Fox"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/bookmark[@name=Fox]" })));

        Assert.Equal("bookmark", got["type"]!.GetValue<string>());
        Assert.Equal("Fox", got["properties"]!["name"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", got["properties"]!["anchorPath"]!.GetValue<string>());
        Assert.Equal("The quick brown fox", got["properties"]!["snippet"]!.GetValue<string>());
    }

    [Fact]
    public void Structure_view_lists_bookmarks()
    {
        var file = CreateDoc(title: "Mapped");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Top"}}]""");

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));

        var bookmark = Assert.Single(data["bookmarks"]!.AsArray())!;
        Assert.Equal("Top", bookmark["name"]!.GetValue<string>());
        Assert.Equal("/body/p[1]", bookmark["anchorPath"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_bookmark_clears_both_markers_and_keeps_content()
    {
        var file = CreateDoc(title: "Sticky text");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Temp"}}]""");

        Edit(file, """[{"op":"remove","path":"/bookmark[@name=Temp]"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            Assert.Empty(body.Descendants<BookmarkStart>());
            Assert.Empty(body.Descendants<BookmarkEnd>());
        }

        Assert.Equal("Sticky text", BodyTexts(file)[0]);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Duplicate_name_is_invalid_args()
    {
        var file = CreateDoc(title: "Once");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Solo"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[2]","type":"bookmark","props":{"name":"Solo"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("already exists", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Invalid_name_is_invalid_args()
    {
        var file = CreateDoc(title: "Naming");

        foreach (var bad in (string[])["1stPlace", "has space"])
        {
            var ops = """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"NAME"}}]"""
                .Replace("NAME", bad, StringComparison.Ordinal);
            var ex = Assert.Throws<AiofficeException>(() => Edit(file, ops));
            Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        }
    }

    [Fact]
    public void Unknown_bookmark_is_invalid_path_with_candidates()
    {
        var file = CreateDoc(title: "Lost");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Here"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/bookmark[@name=There]" })));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/bookmark[@name=Here]", ex.Candidates!);
    }

    [Fact]
    public void Bookmark_on_non_paragraph_is_invalid_args()
    {
        var file = CreateDoc(title: "Cells");
        Edit(file, """[{"op":"add","path":"/body","type":"table","props":{"rows":1,"columns":1}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/table[1]/tr[1]/tc[1]","type":"bookmark","props":{"name":"Cell"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("/p[1]", ex.Suggestion, StringComparison.Ordinal);
    }
}
