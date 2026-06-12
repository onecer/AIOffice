using System.IO.Compression;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace AIOffice.Word.Tests;

public sealed class HeaderFooterTests : WordTestBase
{
    /// <summary>A doc with one default header and one default footer, created through the public grammar.</summary>
    private string CreateWithHeaderFooter()
    {
        var file = CreateDoc(title: "Report");
        var envelope = Edit(file, """
            [
              {"op":"add","path":"/header[1]","type":"header","props":{"text":"Confidential","alignment":"center"}},
              {"op":"add","path":"/footer[1]","type":"footer","props":{"text":"Page footer"}}
            ]
            """);
        Assert.True(envelope.IsOk, envelope.ToJson());
        return file;
    }

    private JsonNode Get(string file, string path) =>
        Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = path })));

    // -------------------------------------------------------------- creation

    [Fact]
    public void Add_header_creates_default_header_with_one_paragraph()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(file, """
            [{"op":"add","path":"/header[1]","type":"header","props":{"text":"Top secret"}}]
            """);

        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("/header[1]", summary["path"]!.GetValue<string>());
        Assert.Equal("/header[1]/p[1]", summary["paragraph"]!.GetValue<string>());

        // Reopen-verify through the same public surface a fresh agent would use.
        var props = Get(file, "/header[1]/p[1]")["properties"]!;
        Assert.Equal("Top secret", props["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_footer_creates_default_footer_with_one_paragraph()
    {
        var file = CreateDoc(title: "Doc");

        Edit(file, """
            [{"op":"add","path":"/footer[1]","type":"footer","props":{"text":"p. 1","alignment":"right"}}]
            """);

        var props = Get(file, "/footer[1]/p[1]")["properties"]!;
        Assert.Equal("p. 1", props["text"]!.GetValue<string>());
        Assert.Equal("right", props["alignment"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_header_when_one_exists_is_invalid_args_with_edit_suggestion()
    {
        var file = CreateWithHeaderFooter();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"again"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("/header[1]/p[1]", ex.Suggestion, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("first")]
    [InlineData("even")]
    public void First_page_and_even_odd_types_are_unsupported_naming_default(string headerType)
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() => Edit(
            file,
            $$$"""[{"op":"add","path":"/header[1]","type":"header","props":{"type":"{{{headerType}}}","text":"x"}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("default", ex.Suggestion, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("M2", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Explicit_default_type_is_accepted()
    {
        var file = CreateDoc(title: "Doc");

        var envelope = Edit(
            file,
            """[{"op":"add","path":"/footer[1]","type":"footer","props":{"type":"default","text":"ok"}}]""");

        Assert.True(envelope.IsOk, envelope.ToJson());
        Assert.Equal("ok", Get(file, "/footer[1]/p[1]")["properties"]!["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Header_creation_must_target_header_1()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/header[2]","type":"header"}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("/header[1]", ex.Candidates!);
    }

    [Fact]
    public void Editing_a_missing_header_is_invalid_path_with_creation_hint()
    {
        var file = CreateDoc(title: "Doc");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/header[1]/p[1]","props":{"text":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("\"type\":\"header\"", ex.Message, StringComparison.Ordinal); // names the creation op
    }

    // ----------------------------------------- paragraph CRUD inside headers

    [Fact]
    public void Add_paragraph_inside_header_lands_at_scoped_path()
    {
        var file = CreateWithHeaderFooter();

        var envelope = Edit(file, """
            [{"op":"add","path":"/header[1]","type":"p","props":{"text":"Second line","italic":true}}]
            """);

        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("/header[1]/p[2]", summary["path"]!.GetValue<string>());
        Assert.Equal("Second line", Get(file, "/header[1]/p[2]")["properties"]!["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Set_header_paragraph_formatting_round_trips_through_get()
    {
        var file = CreateWithHeaderFooter();

        Edit(file, """
            [{"op":"set","path":"/header[1]/p[1]","props":{"text":"Restricted","bold":true,"fontSize":9}}]
            """);

        var props = Get(file, "/header[1]/p[1]")["properties"]!;
        Assert.Equal("Restricted", props["text"]!.GetValue<string>());
        Assert.True(props["bold"]!.GetValue<bool>());
        Assert.Equal(9, props["fontSize"]!.GetValue<double>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Remove_header_paragraph_shifts_following_paths()
    {
        var file = CreateWithHeaderFooter();
        Edit(file, """[{"op":"add","path":"/header[1]","type":"p","props":{"text":"Keep me"}}]""");

        Edit(file, """[{"op":"remove","path":"/header[1]/p[1]"}]""");

        Assert.Equal("Keep me", Get(file, "/header[1]/p[1]")["properties"]!["text"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Removing_the_last_header_paragraph_is_refused_with_a_workaround()
    {
        var file = CreateWithHeaderFooter();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"remove","path":"/header[1]/p[1]"}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("\"text\":\"\"", ex.Suggestion, StringComparison.Ordinal);
        AssertValidatesClean(file); // file untouched and still valid
    }

    [Fact]
    public void Removing_a_whole_header_is_unsupported_with_blank_workaround()
    {
        var file = CreateWithHeaderFooter();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"remove","path":"/header[1]"}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("/header[1]/p[1]", ex.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Set_on_the_header_root_points_at_its_paragraph()
    {
        var file = CreateWithHeaderFooter();

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"set","path":"/header[1]","props":{"text":"x"}}]"""));

        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.Contains("/header[1]/p[1]", ex.Suggestion, StringComparison.Ordinal);
    }

    // ------------------------------------------------------- query + read

    [Fact]
    public void Query_covers_header_and_footer_content_with_scoped_paths()
    {
        var file = CreateWithHeaderFooter();

        var data = Data(Handler.Query(Ctx(file, new JsonObject { ["selector"] = "p:contains('Confidential')" })));

        Assert.Equal(1, data["count"]!.GetValue<int>());
        Assert.Equal("/header[1]/p[1]", data["matches"]!.AsArray()[0]!["path"]!.GetValue<string>());

        var allPaths = Data(Handler.Query(Ctx(file, new JsonObject { ["selector"] = "*" })))["matches"]!.AsArray()
            .Select(m => m!["path"]!.GetValue<string>())
            .ToList();
        Assert.Contains("/footer[1]/p[1]", allPaths);
        Assert.Contains("/body/p[1]", allPaths);
    }

    [Fact]
    public void Structure_view_lists_headers_and_footers()
    {
        var file = CreateWithHeaderFooter();

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));

        var header = Assert.Single(data["headers"]!.AsArray())!;
        Assert.Equal("/header[1]", header["path"]!.GetValue<string>());
        Assert.Equal("header", header["type"]!.GetValue<string>());
        Assert.Equal("Confidential", header["text"]!.GetValue<string>());
        var footer = Assert.Single(data["footers"]!.AsArray())!;
        Assert.Equal("/footer[1]", footer["path"]!.GetValue<string>());
    }

    [Fact]
    public void Structure_view_omits_header_footer_lists_when_absent()
    {
        var file = CreateDoc(title: "Plain");

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));

        Assert.Null(data["headers"]);
        Assert.Null(data["footers"]);
    }

    // ----------------------------------------------------------- rendering

    [Fact]
    public void Render_html_wraps_headers_and_footers_with_paths()
    {
        var file = CreateWithHeaderFooter();

        var html = Data(Handler.Render(Ctx(file, new JsonObject { ["to"] = "html" })))["content"]!.GetValue<string>();

        Assert.Contains("""<header data-aio-path="/header[1]">""", html, StringComparison.Ordinal);
        Assert.Contains("""<p data-aio-path="/header[1]/p[1]">Confidential</p>""", html, StringComparison.Ordinal);
        Assert.Contains("""<footer data-aio-path="/footer[1]">""", html, StringComparison.Ordinal);
        Assert.Contains("""<p data-aio-path="/footer[1]/p[1]">Page footer</p>""", html, StringComparison.Ordinal);

        // Visual order: header block, then body content, then footer block.
        var headerAt = html.IndexOf("<header", StringComparison.Ordinal);
        var bodyAt = html.IndexOf("""<h1 data-aio-path="/body/p[1]">""", StringComparison.Ordinal);
        var footerAt = html.IndexOf("<footer", StringComparison.Ordinal);
        Assert.True(headerAt >= 0 && bodyAt > headerAt && footerAt > bodyAt, html);
    }

    [Fact]
    public void Render_scope_header_renders_only_that_header()
    {
        var file = CreateWithHeaderFooter();

        var html = Data(Handler.Render(Ctx(file, new JsonObject
        {
            ["to"] = "html",
            ["scope"] = "/header[1]",
        })))["content"]!.GetValue<string>();

        Assert.StartsWith("""<header data-aio-path="/header[1]">""", html, StringComparison.Ordinal);
        Assert.DoesNotContain("/body/", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<footer", html, StringComparison.Ordinal);
    }

    // ------------------------------------------------------ round-trip law

    /// <summary>
    /// The round-trip law extended to header/footer parts: a no-edit open+save
    /// of a document with created headers/footers keeps every part byte-identical.
    /// </summary>
    [Fact]
    public void Open_then_save_without_edits_keeps_header_footer_parts_byte_identical()
    {
        var file = CreateWithHeaderFooter();
        var before = File.ReadAllBytes(file);

        var ms = new MemoryStream();
        ms.Write(before);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            _ = doc.MainDocumentPart!.Document!.Body;
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

        using var zipBefore = new ZipArchive(new MemoryStream(before), ZipArchiveMode.Read);
        using var zipAfter = new ZipArchive(new MemoryStream(after), ZipArchiveMode.Read);
        var names = zipBefore.Entries.Select(e => e.FullName).OrderBy(n => n, StringComparer.Ordinal).ToList();
        Assert.Equal(names, zipAfter.Entries.Select(e => e.FullName).OrderBy(n => n, StringComparer.Ordinal).ToList());
        Assert.Contains(names, n => n.Contains("header", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(names, n => n.Contains("footer", StringComparison.OrdinalIgnoreCase));
        foreach (var name in names)
        {
            Assert.True(
                ReadEntry(zipBefore, name).SequenceEqual(ReadEntry(zipAfter, name)),
                $"Zip part '{name}' changed on a no-edit open+save.");
        }
    }

    private static byte[] ReadEntry(ZipArchive zip, string name)
    {
        using var stream = zip.GetEntry(name)!.Open();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }
}
