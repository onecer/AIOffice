using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class ContentControlTests : WordTestBase
{
    private string DocWithControl(string opsJson)
    {
        var file = CreateDoc(title: "Template");
        Edit(file, opsJson);
        return file;
    }

    [Fact]
    public void Add_text_control_reopens_with_sdt_and_appears_in_fields_view()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"clientName","title":"Client Name"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sdt = Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>());
            Assert.Equal("clientName", sdt.SdtProperties!.GetFirstChild<Tag>()!.Val!.Value);
            Assert.Equal("Client Name", sdt.SdtProperties!.GetFirstChild<SdtAlias>()!.Val!.Value);
        }

        var fields = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "fields" })));
        var field = Assert.Single(fields["fields"]!.AsArray())!;
        Assert.Equal("text", field["kind"]!.GetValue<string>());
        Assert.Equal("clientName", field["tag"]!.GetValue<string>());
        Assert.Equal("Client Name", field["title"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_text_control_value_reopens_with_new_text()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"clientName"}}]""");

        Edit(file, """[{"op":"set","path":"/sdt[@tag=clientName]","props":{"text":"Acme"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=clientName]" })));
        Assert.Equal("contentControl", got["type"]!.GetValue<string>());
        Assert.Equal("Acme", got["properties"]!["value"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Dropdown_control_carries_items_and_rejects_off_list_values()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"dropdown","tag":"region","items":["North","South","East"]}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=region]" })));
        var items = got["properties"]!["items"]!.AsArray().Select(i => i!.GetValue<string>()).ToList();
        Assert.Equal(["North", "South", "East"], items);

        Edit(file, """[{"op":"set","path":"/sdt[@tag=region]","props":{"text":"South"}}]""");
        var afterSet = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=region]" })));
        Assert.Equal("South", afterSet["properties"]!["value"]!.GetValue<string>());

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/sdt[@tag=region]","props":{"text":"West"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("North", ex.Candidates!);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Checkbox_control_toggles_checked_state_on_reopen()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"checkbox","tag":"agreed","checked":false}}]""");

        var before = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=agreed]" })));
        Assert.False(before["properties"]!["value"]!.GetValue<bool>());

        Edit(file, """[{"op":"set","path":"/sdt[@tag=agreed]","props":{"checked":true}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var checkbox = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>()
                .Single().SdtProperties!.GetFirstChild<W14.SdtContentCheckBox>()!;
            Assert.Equal(W14.OnOffValues.One, checkbox.Checked!.Val!.Value);
        }

        var after = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=agreed]" })));
        Assert.True(after["properties"]!["value"]!.GetValue<bool>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Date_control_stores_an_iso_value()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"date","tag":"signed"}}]""");

        Edit(file, """[{"op":"set","path":"/sdt[@tag=signed]","props":{"text":"2026-06-13"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=signed]" })));
        Assert.Equal("date", got["properties"]!["kind"]!.GetValue<string>());
        Assert.Equal("2026-06-13", got["properties"]!["value"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Positional_addressing_resolves_a_control()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"only"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[1]" })));
        Assert.Equal("only", got["properties"]!["tag"]!.GetValue<string>());
    }

    [Fact]
    public void Remove_control_keeps_its_filled_content_by_default()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"keep"}}]""");
        Edit(file, """[{"op":"set","path":"/sdt[@tag=keep]","props":{"text":"Survives"}}]""");

        Edit(file, """[{"op":"remove","path":"/sdt[@tag=keep]"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>());
        }

        Assert.Contains("Survives", BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_control_with_keepContent_false_drops_everything()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"gone","text":"Throwaway"}}]""");

        Edit(file, """[{"op":"remove","path":"/sdt[@tag=gone]","props":{"keepContent":false}}]""");

        Assert.DoesNotContain("Throwaway", BodyTexts(file));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Duplicate_tag_is_invalid_args()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"dup"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"dup"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("already exists", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_tag_is_invalid_path_with_candidates()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"here"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=there]" })));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/sdt[@tag=here]", ex.Candidates!);
    }

    [Fact]
    public void Dropdown_without_items_is_invalid_args()
    {
        var file = CreateDoc(title: "NoItems");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"dropdown","tag":"empty"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("items", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_kind_is_invalid_args_with_candidates()
    {
        var file = CreateDoc(title: "BadKind");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"slider","tag":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("text", ex.Candidates!);
    }
}
