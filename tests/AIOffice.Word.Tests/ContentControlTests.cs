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

    // ---------------------------------------------------------- dataBinding (v1.19)

    [Fact]
    public void Control_without_dataBinding_writes_no_dataBinding_and_omits_the_key()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"plain","title":"Plain"}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sdtPr = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>().Single().SdtProperties!;
            Assert.Null(sdtPr.GetFirstChild<DataBinding>());
            // sdtPr children unchanged from today: id, tag, alias, then the text kind child.
            var kinds = sdtPr.ChildElements.Select(c => c.GetType().Name).ToList();
            Assert.Equal(["SdtId", "Tag", "SdtAlias", "SdtContentText"], kinds);
        }

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=plain]" })));
        Assert.False(got["properties"]!.AsObject().ContainsKey("dataBinding"));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Minimal_dataBinding_writes_xpath_after_alias_and_round_trips()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"client","title":"Client","dataBinding":{"xpath":"/root/client"}}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sdtPr = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>().Single().SdtProperties!;
            var binding = Assert.IsType<DataBinding>(sdtPr.GetFirstChild<DataBinding>());
            Assert.Equal("/root/client", binding.XPath!.Value);
            // w:storeItemID is schema-required, so it is always present (empty when omitted).
            Assert.Equal(string.Empty, binding.StoreItemId!.Value);
            Assert.Null(binding.PrefixMappings);

            // Order: w:dataBinding sits AFTER w:alias and BEFORE the kind child.
            var kinds = sdtPr.ChildElements.Select(c => c.GetType().Name).ToList();
            Assert.Equal(kinds.IndexOf("SdtAlias") + 1, kinds.IndexOf("DataBinding"));
            Assert.True(kinds.IndexOf("DataBinding") < kinds.IndexOf("SdtContentText"));
        }

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=client]" })));
        Assert.Equal("/root/client", got["properties"]!["dataBinding"]!["xpath"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Full_dataBinding_on_dropdown_writes_all_three_attributes_and_round_trips()
    {
        var file = DocWithControl(
            """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"dropdown","tag":"region","items":["North","South"],"dataBinding":{"xpath":"/root/region","storeItemId":"{12345678-1234-1234-1234-1234567890AB}","prefixMappings":"xmlns:ns0='urn:demo'"}}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sdtPr = doc.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>().Single().SdtProperties!;
            var binding = sdtPr.GetFirstChild<DataBinding>()!;
            Assert.Equal("/root/region", binding.XPath!.Value);
            Assert.Equal("{12345678-1234-1234-1234-1234567890AB}", binding.StoreItemId!.Value);
            Assert.Equal("xmlns:ns0='urn:demo'", binding.PrefixMappings!.Value);

            // After w:alias-or-id, before the dropdown kind child.
            var kinds = sdtPr.ChildElements.Select(c => c.GetType().Name).ToList();
            Assert.True(kinds.IndexOf("DataBinding") < kinds.IndexOf("SdtContentDropDownList"));
        }

        var projected = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=region]" })))["properties"]!["dataBinding"]!;
        Assert.Equal("/root/region", projected["xpath"]!.GetValue<string>());
        Assert.Equal("{12345678-1234-1234-1234-1234567890AB}", projected["storeItemId"]!.GetValue<string>());
        Assert.Equal("xmlns:ns0='urn:demo'", projected["prefixMappings"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Theory]
    [InlineData("""{"kind":"text","tag":"b","dataBinding":{"xpath":"/r/b"}}""")]
    [InlineData("""{"kind":"dropdown","tag":"b","items":["A","B"],"dataBinding":{"xpath":"/r/b"}}""")]
    [InlineData("""{"kind":"date","tag":"b","dataBinding":{"xpath":"/r/b"}}""")]
    [InlineData("""{"kind":"checkbox","tag":"b","dataBinding":{"xpath":"/r/b"}}""")]
    public void DataBinding_works_on_every_kind_and_set_remove_still_function(string addProps)
    {
        var file = DocWithControl($$"""[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{{addProps}}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=b]" })));
        Assert.Equal("/r/b", got["properties"]!["dataBinding"]!["xpath"]!.GetValue<string>());

        // set still works with a binding present.
        var kind = got["properties"]!["kind"]!.GetValue<string>();
        var setProps = kind == "checkbox"
            ? """{"checked":true}"""
            : kind == "date" ? """{"text":"2026-06-27"}"""
            : kind == "dropdown" ? """{"text":"B"}"""
            : """{"text":"Filled"}""";
        Edit(file, $$"""[{"op":"set","path":"/sdt[@tag=b]","props":{{setProps}}}]""");

        // binding survives the set, and the control is still removable.
        var afterSet = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/sdt[@tag=b]" })));
        Assert.Equal("/r/b", afterSet["properties"]!["dataBinding"]!["xpath"]!.GetValue<string>());
        AssertValidatesClean(file);

        Edit(file, """[{"op":"remove","path":"/sdt[@tag=b]"}]""");
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<SdtBlock>());
    }

    [Fact]
    public void DataBinding_with_empty_xpath_is_invalid_args_naming_xpath()
    {
        var file = CreateDoc(title: "BadBinding");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"x","dataBinding":{"xpath":""}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("xpath", ex.Candidates!);
    }

    [Fact]
    public void DataBinding_with_missing_xpath_is_invalid_args_naming_xpath()
    {
        var file = CreateDoc(title: "NoXpath");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"x","dataBinding":{"storeItemId":"{GUID}"}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("xpath", ex.Candidates!);
    }
}
