using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>v1.5.0 line numbering: w:lnNumType on the section sectPr, get + clear.</summary>
public sealed class LineNumberTests : WordTestBase
{
    private JsonNode GetSection(string file, string path = "/section[1]") =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    [Fact]
    public void Set_line_numbers_writes_lnNumType()
    {
        var file = CreateDoc(title: "Lines");

        Edit(file, """
        [{"op":"set","path":"/section[1]","props":{"lineNumbers":{"start":1,"increment":1,"restart":"continuous"}}}]
        """);

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var ln = doc.MainDocumentPart!.Document!.Body!
                .Elements<SectionProperties>().Single()
                .GetFirstChild<LineNumberType>()!;
            Assert.Equal((short)1, ln.Start!.Value);
            Assert.Equal((short)1, ln.CountBy!.Value);
            Assert.Equal(LineNumberRestartValues.Continuous, ln.Restart!.Value);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Get_reports_line_numbers_reopen_verifiable()
    {
        var file = CreateDoc(title: "Lines");
        Edit(file, """
        [{"op":"set","path":"/section[1]","props":{"lineNumbers":{"start":5,"increment":2,"restart":"newPage","distance":"0.5cm"}}}]
        """);

        var ln = GetSection(file)["properties"]!["lineNumbers"]!;
        Assert.Equal(5, ln["start"]!.GetValue<int>());
        Assert.Equal(2, ln["increment"]!.GetValue<int>());
        Assert.Equal("newPage", ln["restart"]!.GetValue<string>());
        Assert.Equal(0.5, ln["distanceCm"]!.GetValue<double>());
    }

    [Fact]
    public void Restart_new_section_round_trips()
    {
        var file = CreateDoc(title: "Lines");
        Edit(file, """
        [{"op":"set","path":"/section[1]","props":{"lineNumbers":{"start":1,"increment":1,"restart":"newSection"}}}]
        """);

        Assert.Equal("newSection", GetSection(file)["properties"]!["lineNumbers"]!["restart"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Defaults_apply_when_only_restart_is_given()
    {
        var file = CreateDoc(title: "Lines");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"lineNumbers":{"restart":"continuous"}}}]""");

        var ln = GetSection(file)["properties"]!["lineNumbers"]!;
        Assert.Equal(1, ln["start"]!.GetValue<int>());
        Assert.Equal(1, ln["increment"]!.GetValue<int>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Clear_with_none_removes_lnNumType()
    {
        var file = CreateDoc(title: "Lines");
        Edit(file, """[{"op":"set","path":"/section[1]","props":{"lineNumbers":{"start":1,"increment":1}}}]""");
        Assert.NotNull(GetSection(file)["properties"]!["lineNumbers"]);

        Edit(file, """[{"op":"set","path":"/section[1]","props":{"lineNumbers":"none"}}]""");

        Assert.Null(GetSection(file)["properties"]!["lineNumbers"]);
        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Null(doc.MainDocumentPart!.Document!.Body!
                .Elements<SectionProperties>().Single()
                .GetFirstChild<LineNumberType>());
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Section_without_line_numbers_reports_null()
    {
        var file = CreateDoc(title: "Lines");
        Assert.Null(GetSection(file)["properties"]!["lineNumbers"]);
    }

    [Fact]
    public void Invalid_restart_is_invalid_args()
    {
        var file = CreateDoc(title: "Lines");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/section[1]","props":{"lineNumbers":{"restart":"never"}}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Bad_increment_is_invalid_args()
    {
        var file = CreateDoc(title: "Lines");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/section[1]","props":{"lineNumbers":{"increment":0}}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }
}
