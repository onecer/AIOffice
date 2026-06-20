using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// 1.13 document-level protection on the document root: <c>set /</c> writes
/// enforcement-flag <c>w:documentProtection</c> / <c>w:writeProtection</c> into
/// settings.xml (this is NOT password encryption — CONTRACT §10), and <c>get /</c>
/// round-trips the reported shape. Mirrors the xlsx workbook-root protection
/// surface. Every mutation must still validate clean for Word.
/// </summary>
public sealed class DocumentProtectionTests : WordTestBase
{
    private JsonNode Get(string file) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/" })));

    /// <summary>The w:documentProtection element of the document, or null when absent.</summary>
    private static DocumentProtection? DocProtection(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        return doc.MainDocumentPart?.DocumentSettingsPart?.Settings?.GetFirstChild<DocumentProtection>();
    }

    private static WriteProtection? WriteProt(string file)
    {
        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        return doc.MainDocumentPart?.DocumentSettingsPart?.Settings?.GetFirstChild<WriteProtection>();
    }

    // ------------------------------------------------------------ set: edit modes

    [Theory]
    [InlineData("readOnly")]
    [InlineData("comments")]
    [InlineData("trackedChanges")]
    [InlineData("forms")]
    public void Set_each_edit_mode_writes_documentProtection_with_enforcement(string mode)
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(
            file,
            """[{"op":"set","path":"/","props":{"protection":{"edit":"MODE"}}}]""".Replace("MODE", mode, StringComparison.Ordinal));
        Assert.True(envelope.IsOk, envelope.ToJson());

        var prot = DocProtection(file);
        Assert.NotNull(prot);
        Assert.Equal(mode, prot!.Edit!.InnerText);
        Assert.True(OnOffValue.ToBoolean(prot.Enforcement!));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_protection_default_enforce_is_true()
    {
        var file = CreateDoc(title: "Doc");

        // No explicit enforce -> defaults to enforced (the whole point of the op).
        Edit(file, """[{"op":"set","path":"/","props":{"protection":{"edit":"readOnly"}}}]""");

        var prot = DocProtection(file);
        Assert.NotNull(prot);
        Assert.True(OnOffValue.ToBoolean(prot!.Enforcement!));
    }

    // ----------------------------------------------------------------- get round-trip

    [Fact]
    public void Get_root_round_trips_protection_and_enforced_flag()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"set","path":"/","props":{"protection":{"edit":"comments"}}}]""");

        var props = Get(file);
        Assert.Equal("document", props["type"]!.GetValue<string>());
        var protection = props["properties"]!["protection"]!;
        Assert.Equal("comments", protection["edit"]!.GetValue<string>());
        Assert.True(protection["enforced"]!.GetValue<bool>());
    }

    [Fact]
    public void Get_root_on_unprotected_doc_reports_none_and_false()
    {
        var file = CreateDoc(title: "Doc");

        var props = Get(file)["properties"]!;
        Assert.Equal("none", props["protection"]!["edit"]!.GetValue<string>());
        Assert.False(props["protection"]!["enforced"]!.GetValue<bool>());
        Assert.False(props["readOnlyRecommended"]!.GetValue<bool>());
    }

    // ---------------------------------------------------------------------- removal

    [Fact]
    public void Edit_none_removes_documentProtection()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"set","path":"/","props":{"protection":{"edit":"readOnly"}}}]""");
        Assert.NotNull(DocProtection(file));

        Edit(file, """[{"op":"set","path":"/","props":{"protection":{"edit":"none"}}}]""");

        Assert.Null(DocProtection(file));
        Assert.Equal("none", Get(file)["properties"]!["protection"]!["edit"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Enforce_false_lifts_protection_even_with_a_mode()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"set","path":"/","props":{"protection":{"edit":"forms"}}}]""");
        Assert.NotNull(DocProtection(file));

        Edit(file, """[{"op":"set","path":"/","props":{"protection":{"edit":"forms","enforce":false}}}]""");

        Assert.Null(DocProtection(file));
        AssertValidatesClean(file);
    }

    // -------------------------------------------------------------- readOnlyRecommended

    [Fact]
    public void Set_readOnlyRecommended_writes_writeProtection_recommended()
    {
        var file = CreateDoc(title: "Doc");

        Edit(file, """[{"op":"set","path":"/","props":{"readOnlyRecommended":true}}]""");

        var write = WriteProt(file);
        Assert.NotNull(write);
        Assert.True(OnOffValue.ToBoolean(write!.Recommended!));
        Assert.True(Get(file)["properties"]!["readOnlyRecommended"]!.GetValue<bool>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_readOnlyRecommended_false_removes_writeProtection()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"set","path":"/","props":{"readOnlyRecommended":true}}]""");
        Assert.NotNull(WriteProt(file));

        Edit(file, """[{"op":"set","path":"/","props":{"readOnlyRecommended":false}}]""");

        Assert.Null(WriteProt(file));
        Assert.False(Get(file)["properties"]!["readOnlyRecommended"]!.GetValue<bool>());
    }

    [Fact]
    public void Protection_and_readOnlyRecommended_coexist_in_one_op()
    {
        var file = CreateDoc(title: "Doc");

        Edit(
            file,
            """[{"op":"set","path":"/","props":{"protection":{"edit":"trackedChanges"},"readOnlyRecommended":true}}]""");

        Assert.Equal("trackedChanges", DocProtection(file)!.Edit!.InnerText);
        Assert.NotNull(WriteProt(file));

        var props = Get(file)["properties"]!;
        Assert.Equal("trackedChanges", props["protection"]!["edit"]!.GetValue<string>());
        Assert.True(props["readOnlyRecommended"]!.GetValue<bool>());
        AssertValidatesClean(file);
    }

    // ------------------------------------------------------- password ignored (CONTRACT §10)

    [Fact]
    public void Password_is_ignored_protection_is_still_flag_enforcement()
    {
        var file = CreateDoc(title: "Doc");

        // A password buys nothing here: enforcement-flag protection, not encryption.
        // It must not error and must not turn into a strong-encryption attribute.
        Edit(
            file,
            """[{"op":"set","path":"/","props":{"protection":{"edit":"readOnly","password":"hunter2"}}}]""");

        var prot = DocProtection(file);
        Assert.NotNull(prot);
        Assert.Equal("readOnly", prot!.Edit!.InnerText);
        Assert.Null(prot.CryptographicProviderType); // no encryption attributes authored
        Assert.Null(prot.Hash);
        AssertValidatesClean(file);
    }

    // -------------------------------------------------------------------- guardrails

    [Fact]
    public void Unknown_root_prop_is_invalid_args_with_candidates()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/","props":{"locked":true}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("protection", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Protection_object_without_edit_is_invalid_args()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/","props":{"protection":{"enforce":true}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Unknown_edit_mode_is_invalid_args_with_candidates()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/","props":{"protection":{"edit":"lockAll"}}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.NotNull(ex.Candidates);
        Assert.Contains("readOnly", ex.Candidates!);
    }

    [Fact]
    public void Propless_root_set_still_refuses_with_the_old_pointer()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/","props":{}}]"""));

        // A propless set / has nothing to protect, so it keeps the refusal that
        // points at the real edit targets (now also mentioning protection).
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    [Fact]
    public void Non_set_root_op_still_refuses()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"remove","path":"/"}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
    }

    // --------------------------------------------------------------------- idempotency

    [Fact]
    public void Re_setting_a_mode_overwrites_rather_than_duplicating()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"set","path":"/","props":{"protection":{"edit":"readOnly"}}}]""");
        Edit(file, """[{"op":"set","path":"/","props":{"protection":{"edit":"comments"}}}]""");

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        var protections = doc.MainDocumentPart!.DocumentSettingsPart!.Settings!
            .Elements<DocumentProtection>().ToList();
        Assert.Single(protections);
        Assert.Equal("comments", protections[0].Edit!.InnerText);
    }
}
