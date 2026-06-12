using System.Text.Json.Nodes;
using AIOffice.Core;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class CreateValidateTests : WordTestBase
{
    [Fact]
    public void Create_produces_a_validator_clean_docx()
    {
        var file = CreateDoc();

        AssertValidatesClean(file);
    }

    [Fact]
    public void Create_with_title_adds_a_heading1_paragraph()
    {
        var file = CreateDoc(title: "Quarterly Report");

        var get = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/body/p[1]" })));

        Assert.Equal("Quarterly Report", get["properties"]!["text"]!.GetValue<string>());
        Assert.Equal("Heading1", get["properties"]!["style"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Create_over_an_existing_file_is_invalid_args()
    {
        CreateDoc();

        var ex = Assert.Throws<AiofficeException>(() => CreateDoc());

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.False(string.IsNullOrWhiteSpace(ex.Suggestion));
    }

    [Fact]
    public void Create_returns_rev_and_file_in_meta()
    {
        var envelope = Handler.Create(Ctx("meta.docx"));

        Assert.True(envelope.IsOk);
        Assert.Equal(Rev.OfFile(Path.Combine(Dir, "meta.docx")), envelope.Meta.Rev);
        Assert.NotNull(envelope.Meta.File);
    }

    [Fact]
    public void Validate_reports_zero_issues_for_a_fresh_document()
    {
        var file = CreateDoc(title: "Valid");

        var data = Data(Handler.Validate(Ctx(file)));

        Assert.True(data["valid"]!.GetValue<bool>());
        Assert.Equal(0, data["count"]!.GetValue<int>());
    }

    [Fact]
    public void Missing_file_is_file_not_found()
    {
        var ex = Assert.Throws<AiofficeException>(() => Handler.Validate(Ctx("ghost.docx")));

        Assert.Equal(ErrorCodes.FileNotFound, ex.Code);
    }
}
