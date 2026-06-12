using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>M5: first-page and even-page header/footer variants (the M2 debt).</summary>
public sealed class HeaderFooterVariantTests : WordTestBase
{
    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    private string CreateWithAllHeaderVariants()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """
            [
              {"op":"add","path":"/header[1]","type":"header","props":{"text":"All pages"}},
              {"op":"add","path":"/header[firstPage]","type":"header","props":{"text":"First page only"}},
              {"op":"add","path":"/header[even]","type":"header","props":{"text":"Even pages"}}
            ]
            """);
        return file;
    }

    // -------------------------------------------------------------- creation

    [Fact]
    public void Add_first_page_header_wires_title_page_flag()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/header[firstPage]","type":"header","props":{"text":"Cover"}}]
            """);

        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("firstPage", summary["variant"]!.GetValue<string>());
        Assert.Equal("/header[1]", summary["path"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var sectPr = doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().Single();
            Assert.NotNull(sectPr.GetFirstChild<TitlePage>());
            var reference = sectPr.Elements<HeaderReference>().Single();
            Assert.Equal(HeaderFooterValues.First, reference.Type!.Value);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_even_footer_wires_even_and_odd_headers_setting()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/footer[even]","type":"footer","props":{"text":"verso"}}]
            """);

        Assert.Equal("even", Data(envelope)["ops"]!.AsArray()[0]!["variant"]!.GetValue<string>());

        using (var doc = WordprocessingDocument.Open(file, isEditable: false))
        {
            var settings = doc.MainDocumentPart!.DocumentSettingsPart!.Settings!;
            Assert.NotNull(settings.GetFirstChild<EvenAndOddHeaders>());
            var reference = doc.MainDocumentPart.Document!.Body!
                .Elements<SectionProperties>().Single()
                .Elements<FooterReference>().Single();
            Assert.Equal(HeaderFooterValues.Even, reference.Type!.Value);
        }

        AssertValidatesClean(file);
    }

    [Fact]
    public void All_three_variants_coexist_and_resolve_by_name()
    {
        var file = CreateWithAllHeaderVariants();

        Assert.Equal("All pages", Get(file, "/header[default]")["properties"]!["text"]!.GetValue<string>());
        Assert.Equal("First page only", Get(file, "/header[firstPage]")["properties"]!["text"]!.GetValue<string>());
        Assert.Equal("Even pages", Get(file, "/header[even]")["properties"]!["text"]!.GetValue<string>());

        // Named roots report their variant through get.
        Assert.Equal("firstPage", Get(file, "/header[firstPage]")["properties"]!["variant"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Named_variant_paths_address_paragraphs_inside()
    {
        var file = CreateWithAllHeaderVariants();

        Edit(file, """[{"op":"set","path":"/header[firstPage]/p[1]","props":{"text":"Cover page","bold":true}}]""");

        var props = Get(file, "/header[firstPage]/p[1]")["properties"]!;
        Assert.Equal("Cover page", props["text"]!.GetValue<string>());
        Assert.True(props["bold"]!.GetValue<bool>());

        // The default header was not touched.
        Assert.Equal("All pages", Get(file, "/header[default]")["properties"]!["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Structure_view_reports_variants()
    {
        var file = CreateWithAllHeaderVariants();

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));
        var variants = data["headers"]!.AsArray()
            .Select(h => h!["variant"]!.GetValue<string>())
            .OrderBy(v => v, StringComparer.Ordinal)
            .ToList();

        Assert.Equal(["default", "even", "firstPage"], variants);
    }

    // -------------------------------------------------------------- refusals

    [Fact]
    public void Adding_an_existing_variant_is_invalid_args_with_edit_suggestion()
    {
        var file = CreateWithAllHeaderVariants();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/header[firstPage]","type":"header","props":{"text":"again"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("/header[firstPage]/p[1]", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Conflicting_path_variant_and_props_type_is_invalid_args()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/header[firstPage]","type":"header","props":{"type":"even","text":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
    }

    [Fact]
    public void Odd_type_is_invalid_args_explaining_the_ooxml_model()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/footer[1]","type":"footer","props":{"type":"odd","text":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("even", ex.Suggestion, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("default", ex.Candidates!);
    }

    [Fact]
    public void Resolving_a_missing_variant_is_invalid_path_with_creation_hint()
    {
        var file = CreateDoc(title: "Doc");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"base"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Handler.Get(Ctx(file, new JsonObject { ["path"] = "/header[even]" })));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("\"type\":\"header\"", ex.Message, StringComparison.Ordinal);
        Assert.Contains("/header[default]", ex.Candidates!);
    }

    // ------------------------------------------------------------ round trip

    [Fact]
    public void Open_then_save_with_variant_headers_keeps_every_part_byte_identical()
    {
        var file = CreateWithAllHeaderVariants();
        Edit(file, """[{"op":"add","path":"/footer[even]","type":"footer","props":{"text":"verso"}}]""");

        var before = File.ReadAllBytes(file);

        var ms = new MemoryStream();
        ms.Write(before);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            _ = doc.MainDocumentPart!.Document!.Body;
            _ = doc.MainDocumentPart.DocumentSettingsPart?.Settings;
            foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
            {
                _ = headerPart.Header;
            }

            foreach (var footerPart in doc.MainDocumentPart.FooterParts)
            {
                _ = footerPart.Footer;
            }
        }

        var after = ms.ToArray();

        using var zipBefore = new System.IO.Compression.ZipArchive(new MemoryStream(before));
        using var zipAfter = new System.IO.Compression.ZipArchive(new MemoryStream(after));
        foreach (var entry in zipBefore.Entries)
        {
            using var beforeStream = entry.Open();
            using var afterStream = zipAfter.GetEntry(entry.FullName)!.Open();
            using var bufferBefore = new MemoryStream();
            using var bufferAfter = new MemoryStream();
            beforeStream.CopyTo(bufferBefore);
            afterStream.CopyTo(bufferAfter);
            Assert.True(
                bufferBefore.ToArray().SequenceEqual(bufferAfter.ToArray()),
                $"Zip part '{entry.FullName}' changed on a no-edit open+save.");
        }
    }
}
