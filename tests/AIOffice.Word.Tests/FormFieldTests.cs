using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class FormFieldTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    private JsonNode Fields(string file) =>
        Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "fields" })));

    [Fact]
    public void Add_text_form_field_reopens_with_ffdata_and_lists_in_fields_view()
    {
        var file = CreateDoc(title: "Form");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"text","name":"clientName","default":"Acme","maxLength":50}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var ffData = Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<FormFieldData>());
            Assert.Equal("clientName", ffData.GetFirstChild<FormFieldName>()!.Val!.Value);
            var textInput = ffData.GetFirstChild<TextInput>()!;
            Assert.Equal("Acme", textInput.GetFirstChild<DefaultTextBoxFormFieldString>()!.Val!.Value);
            Assert.Equal((short)50, textInput.GetFirstChild<MaxLength>()!.Val!.Value);
        }

        var fields = Fields(file);
        var field = Assert.Single(fields["fields"]!.AsArray())!;
        Assert.Equal("formField", field["kind"]!.GetValue<string>());
        Assert.Equal("text", field["fieldKind"]!.GetValue<string>());
        Assert.Equal("clientName", field["name"]!.GetValue<string>());
        Assert.Equal("Acme", field["value"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_text_form_field_value_reopens_with_new_result()
    {
        var file = CreateDoc(title: "FormSet");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"text","name":"city"}}]""");

        Edit(file, """[{"op":"set","path":"/formField[@name=city]","props":{"value":"Berlin"}}]""");

        var got = Get(file, "/formField[@name=city]");
        Assert.Equal("formField", got["type"]!.GetValue<string>());
        Assert.Equal("Berlin", got["properties"]!["value"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_checkbox_form_field_toggles_checked_on_set()
    {
        var file = CreateDoc(title: "FormCheck");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"checkbox","name":"agreed","checked":false}}]""");

        var before = Get(file, "/formField[@name=agreed]");
        Assert.False(before["properties"]!["value"]!.GetValue<bool>());
        Assert.Equal("checkbox", before["properties"]!["fieldKind"]!.GetValue<string>());

        Edit(file, """[{"op":"set","path":"/formField[@name=agreed]","props":{"checked":true}}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var cb = doc.MainDocumentPart!.Document!.Body!.Descendants<FormFieldData>().Single().GetFirstChild<CheckBox>()!;
            Assert.True(cb.GetFirstChild<Checked>()!.Val!.Value);
        }

        var after = Get(file, "/formField[@name=agreed]");
        Assert.True(after["properties"]!["value"]!.GetValue<bool>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_dropdown_form_field_carries_items_and_rejects_off_list_values()
    {
        var file = CreateDoc(title: "FormDrop");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"dropdown","name":"region","items":["North","South","East"]}}]""");

        var got = Get(file, "/formField[@name=region]");
        var items = got["properties"]!["items"]!.AsArray().Select(i => i!.GetValue<string>()).ToList();
        Assert.Equal(["North", "South", "East"], items);

        Edit(file, """[{"op":"set","path":"/formField[@name=region]","props":{"value":"South"}}]""");
        var afterSet = Get(file, "/formField[@name=region]");
        Assert.Equal("South", afterSet["properties"]!["value"]!.GetValue<string>());

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/formField[@name=region]","props":{"value":"West"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("North", ex.Candidates!);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_form_field_drops_its_run_chain()
    {
        var file = CreateDoc(title: "FormRemove");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"text","name":"gone","default":"x"}}]""");

        Edit(file, """[{"op":"remove","path":"/formField[@name=gone]"}]""");

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            Assert.Empty(doc.MainDocumentPart!.Document!.Body!.Descendants<FormFieldData>());
        }

        var ex = Assert.Throws<AiofficeException>(() => Get(file, "/formField[@name=gone]"));
        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        AssertValidatesClean(file);
    }

    [Fact]
    public void Fields_view_distinguishes_form_fields_from_content_controls()
    {
        var file = CreateDoc(title: "Mixed");
        Edit(file, """
            [
              {"op":"add","path":"/body/p[1]","type":"contentControl","props":{"kind":"text","tag":"cc1"}},
              {"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"text","name":"ff1"}}
            ]
            """);

        var fields = Fields(file)["fields"]!.AsArray();
        Assert.Contains(fields, f => f!["kind"]!.GetValue<string>() == "text" && f["tag"]?.GetValue<string>() == "cc1");
        Assert.Contains(fields, f => f!["kind"]!.GetValue<string>() == "formField" && f["name"]?.GetValue<string>() == "ff1");
    }

    [Fact]
    public void Duplicate_name_is_invalid_args()
    {
        var file = CreateDoc(title: "Dup");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"text","name":"dup"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"text","name":"dup"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("already exists", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unknown_kind_is_invalid_args_with_candidates()
    {
        var file = CreateDoc(title: "BadKind");
        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"slider","name":"x"}}]"""));
        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("checkbox", ex.Candidates!);
    }

    [Fact]
    public void Unknown_name_is_invalid_path_with_candidates()
    {
        var file = CreateDoc(title: "BadName");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"text","name":"here"}}]""");

        var ex = Assert.Throws<AiofficeException>(() => Get(file, "/formField[@name=there]"));
        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/formField[@name=here]", ex.Candidates!);
    }

    [Fact]
    public void Structure_view_lists_form_fields()
    {
        var file = CreateDoc(title: "StructForm");
        Edit(file, """[{"op":"add","path":"/body/p[1]","type":"formField","props":{"kind":"text","name":"ff"}}]""");

        var structure = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));
        var field = Assert.Single(structure["formFields"]!.AsArray())!;
        Assert.Equal("/formField[@name=ff]", field["path"]!.GetValue<string>());
    }
}
