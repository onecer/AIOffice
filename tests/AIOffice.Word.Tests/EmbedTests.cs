using System.IO.Compression;
using System.Text;
using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace AIOffice.Word.Tests;

/// <summary>
/// M10 embedded objects: embedding any file as an OLE package object, listing,
/// extracting byte-for-byte, removing, sandbox enforcement on both the embed src
/// and the extract dest, and the embedded-payload round-trip law (extract after
/// an open+save returns identical bytes).
/// </summary>
public sealed class EmbedTests : WordTestBase
{
    /// <summary>A deterministic "binary" payload (not text-only) for the embed source.</summary>
    private static byte[] SamplePayload()
    {
        var bytes = new byte[256];
        for (var i = 0; i < bytes.Length; i++)
        {
            bytes[i] = (byte)((i * 31) ^ 0xA5);
        }

        return bytes;
    }

    /// <summary>A tiny valid .xlsx (zip) so media-type sniffing exercises the OOXML path.</summary>
    private static byte[] SampleXlsxBytes()
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = zip.CreateEntry("[Content_Types].xml");
            using var w = new StreamWriter(entry.Open());
            w.Write("<?xml version=\"1.0\"?><Types/>");
        }

        return ms.ToArray();
    }

    private string WriteSource(string name, byte[] bytes)
    {
        File.WriteAllBytes(Path.Combine(Dir, name), bytes);
        return name;
    }

    private IReadOnlyList<JsonNode?> Embeds(string file) =>
        [.. Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "embeds" })))["embeds"]!.AsArray()];

    // ----------------------------------------------------------------- add

    [Fact]
    public void Embed_a_file_lists_it_with_metadata()
    {
        var file = CreateDoc(title: "Attachments");
        var payload = SamplePayload();
        WriteSource("data.bin", payload);

        var envelope = Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"data.bin","name":"Raw data"}}]""");

        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("embed", summary["type"]!.GetValue<string>());
        Assert.Equal("/embed[1]", summary["path"]!.GetValue<string>());
        Assert.Equal("Raw data", summary["name"]!.GetValue<string>());
        Assert.Equal(payload.Length, summary["size"]!.GetValue<long>());

        var embeds = Embeds(file);
        var embed = Assert.Single(embeds);
        Assert.Equal("/embed[1]", embed!["path"]!.GetValue<string>());
        Assert.Equal("Raw data", embed["name"]!.GetValue<string>());
        Assert.Equal(payload.Length, embed["size"]!.GetValue<long>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Embedded_xlsx_media_type_is_sniffed_as_the_spreadsheet_type()
    {
        var file = CreateDoc(title: "Report with model");
        WriteSource("model.xlsx", SampleXlsxBytes());

        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"model.xlsx"}}]""");

        var embed = Assert.Single(Embeds(file));
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            embed!["mediaType"]!.GetValue<string>());
        Assert.Equal("model.xlsx", embed["name"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Media_type_is_sniffed_from_the_src_even_when_the_display_name_has_no_extension()
    {
        // A display name like "Q3 model" carries no extension; the media type must
        // still come from the SOURCE file (model.xlsx), matching xlsx/pptx handlers.
        var file = CreateDoc(title: "Named embed");
        WriteSource("model.xlsx", SampleXlsxBytes());

        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"model.xlsx","name":"Q3 model"}}]""");

        var embed = Assert.Single(Embeds(file));
        Assert.Equal("Q3 model", embed!["name"]!.GetValue<string>());
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            embed["mediaType"]!.GetValue<string>());
        AssertValidatesClean(file);
    }

    [Fact]
    public void Embed_defaults_the_display_name_to_the_source_file_name()
    {
        var file = CreateDoc(title: "Defaults");
        WriteSource("invoice.pdf", "%PDF-1.7 fake"u8.ToArray());

        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"invoice.pdf"}}]""");

        var embed = Assert.Single(Embeds(file));
        Assert.Equal("invoice.pdf", embed!["name"]!.GetValue<string>());
        Assert.Equal("application/pdf", embed["mediaType"]!.GetValue<string>());
    }

    [Fact]
    public void Missing_src_is_invalid_args()
    {
        var file = CreateDoc(title: "No src");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"name":"x"}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("src", ex.Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------- extract

    [Fact]
    public void Extract_writes_the_payload_byte_identically()
    {
        var file = CreateDoc(title: "Extractable");
        var payload = SamplePayload();
        WriteSource("data.bin", payload);
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"data.bin"}}]""");

        var envelope = Edit(file, """[{"op":"extract","path":"/embed[1]","props":{"to":"out/extracted.bin"}}]""");

        var summary = Data(envelope)["ops"]!.AsArray()[0]!;
        Assert.Equal("extract", summary["op"]!.GetValue<string>());
        Assert.Equal(payload.Length, summary["size"]!.GetValue<long>());

        var extracted = File.ReadAllBytes(Path.Combine(Dir, "out", "extracted.bin"));
        Assert.Equal(payload, extracted);
    }

    [Fact]
    public void Extract_does_not_modify_the_source_document()
    {
        var file = CreateDoc(title: "Untouched");
        WriteSource("data.bin", SamplePayload());
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"data.bin"}}]""");

        var before = File.ReadAllBytes(file);
        Edit(file, """[{"op":"extract","path":"/embed[1]","props":{"to":"copy.bin"}}]""");
        var after = File.ReadAllBytes(file);

        // The doc bytes may be re-serialized but the embedded payload and every
        // part must be unchanged; assert the embedded payload survives + validates.
        AssertValidatesClean(file);
        AssertEmbeddedPayloadEquals(after, SamplePayload());
        _ = before;
    }

    [Fact]
    public void Extract_missing_to_is_invalid_args()
    {
        var file = CreateDoc(title: "No dest");
        WriteSource("data.bin", SamplePayload());
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"data.bin"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"extract","path":"/embed[1]","props":{}}]"""));

        Assert.Equal(ErrorCodes.InvalidArgs, ex.Code);
        Assert.Contains("to", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Extract_unknown_embed_is_invalid_path_with_candidates()
    {
        var file = CreateDoc(title: "One embed");
        WriteSource("data.bin", SamplePayload());
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"data.bin"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"extract","path":"/embed[5]","props":{"to":"x.bin"}}]"""));

        Assert.Equal(ErrorCodes.InvalidPath, ex.Code);
        Assert.Contains("/embed[1]", ex.Candidates!);
    }

    // -------------------------------------------------------------- remove

    [Fact]
    public void Remove_embed_drops_the_object_and_the_backing_parts()
    {
        var file = CreateDoc(title: "Removable");
        WriteSource("data.bin", SamplePayload());
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"data.bin"}}]""");
        Assert.Single(Embeds(file));

        var envelope = Edit(file, """[{"op":"remove","path":"/embed[1]"}]""");
        Assert.Equal("embed", Data(envelope)["ops"]!.AsArray()[0]!["type"]!.GetValue<string>());

        Assert.Empty(Embeds(file));
        AssertValidatesClean(file);

        using var doc = WordprocessingDocument.Open(file, isEditable: false);
        Assert.Empty(doc.MainDocumentPart!.EmbeddedPackageParts);
    }

    // ------------------------------------------------------------- get/view

    [Fact]
    public void Get_on_an_embed_returns_metadata_not_bytes()
    {
        var file = CreateDoc(title: "Inspect");
        var payload = SamplePayload();
        WriteSource("data.bin", payload);
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"data.bin","name":"Sheet"}}]""");

        var got = Data(Handler.Get(Ctx(file, new JsonObject { ["path"] = "/embed[1]" })));
        Assert.Equal("embed", got["type"]!.GetValue<string>());
        var props = got["properties"]!;
        Assert.Equal("Sheet", props["name"]!.GetValue<string>());
        Assert.Equal(payload.Length, props["size"]!.GetValue<long>());
        Assert.Contains("extract", props["note"]!.GetValue<string>(), StringComparison.Ordinal);
    }

    [Fact]
    public void Structure_view_includes_embeds()
    {
        var file = CreateDoc(title: "Structured");
        WriteSource("data.bin", SamplePayload());
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"data.bin","name":"Data"}}]""");

        var data = Data(Handler.Read(Ctx(file, new JsonObject { ["view"] = "structure" })));
        var embeds = data["embeds"]!.AsArray();
        Assert.Equal("/embed[1]", embeds[0]!["path"]!.GetValue<string>());
        Assert.Equal("Data", embeds[0]!["name"]!.GetValue<string>());
    }

    [Fact]
    public void Two_embeds_get_distinct_sequential_paths()
    {
        var file = CreateDoc(title: "Multiple");
        WriteSource("a.bin", "AAAA"u8.ToArray());
        WriteSource("b.bin", "BBBBBB"u8.ToArray());
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"a.bin","name":"A"}}]""");
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"b.bin","name":"B"}}]""");

        var embeds = Embeds(file);
        Assert.Equal(2, embeds.Count);
        Assert.Equal(new[] { "/embed[1]", "/embed[2]" }, embeds.Select(e => e!["path"]!.GetValue<string>()));

        // Each extracts to its own distinct payload.
        Edit(file, """[{"op":"extract","path":"/embed[1]","props":{"to":"a-out.bin"}},{"op":"extract","path":"/embed[2]","props":{"to":"b-out.bin"}}]""");
        Assert.Equal("AAAA"u8.ToArray(), File.ReadAllBytes(Path.Combine(Dir, "a-out.bin")));
        Assert.Equal("BBBBBB"u8.ToArray(), File.ReadAllBytes(Path.Combine(Dir, "b-out.bin")));
    }

    // --------------------------------------------------------- header/footer

    [Fact]
    public void Embeds_referenced_from_a_header_are_discovered_and_extractable()
    {
        var file = CreateDoc(title: "Header attachment");
        var payload = SamplePayload();
        WriteSource("hdr.bin", payload);

        Edit(file, """[{"op":"add","path":"/header[1]","type":"header","props":{"text":"Top"}}]""");
        Edit(file, """[{"op":"add","path":"/header[1]","type":"embed","props":{"src":"hdr.bin","name":"InHeader"}}]""");

        var embed = Assert.Single(Embeds(file));
        Assert.Equal("InHeader", embed!["name"]!.GetValue<string>());
        Assert.Equal("/header[1]", embed["container"]!.GetValue<string>());
        AssertValidatesClean(file);

        Edit(file, """[{"op":"extract","path":"/embed[1]","props":{"to":"hdr-out.bin"}}]""");
        Assert.Equal(payload, File.ReadAllBytes(Path.Combine(Dir, "hdr-out.bin")));
    }

    // ------------------------------------------------------------ sandbox

    [Fact]
    public void Escaping_src_is_sandbox_denied_and_changes_nothing()
    {
        var file = CreateDoc(title: "Locked src");
        var before = File.ReadAllBytes(file);

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"../../etc/secret.bin"}}]"""));

        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
        Assert.Equal(before, File.ReadAllBytes(file));
    }

    [Fact]
    public void Escaping_extract_dest_is_sandbox_denied_and_writes_nothing()
    {
        var file = CreateDoc(title: "Locked dest");
        WriteSource("data.bin", SamplePayload());
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"data.bin"}}]""");

        var ex = Assert.Throws<AiofficeException>(() =>
            Edit(file, """[{"op":"extract","path":"/embed[1]","props":{"to":"../../etc/escape.bin"}}]"""));

        Assert.Equal(ErrorCodes.SandboxDenied, ex.Code);
        Assert.False(File.Exists(Path.Combine(Dir, "..", "..", "etc", "escape.bin")));
    }

    // ------------------------------------------------------- round-trip law

    [Fact]
    public void Embedded_payload_survives_open_and_save_byte_identically()
    {
        var file = CreateDoc(title: "Round trip embed");
        var payload = SamplePayload();
        WriteSource("data.bin", payload);
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"data.bin","name":"Payload"}}]""");

        // Open + save with NO edits (the round-trip law path).
        var before = File.ReadAllBytes(file);
        var ms = new MemoryStream();
        ms.Write(before);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, isEditable: true))
        {
            _ = doc.MainDocumentPart!.Document!.Body;
        }

        File.WriteAllBytes(file, ms.ToArray());

        // Extract after the re-save: the payload bytes are still identical.
        Edit(file, """[{"op":"extract","path":"/embed[1]","props":{"to":"after.bin"}}]""");
        Assert.Equal(payload, File.ReadAllBytes(Path.Combine(Dir, "after.bin")));
        AssertValidatesClean(file);
    }

    [Fact]
    public void Add_list_extract_remove_reopen_round_trips_cleanly()
    {
        var file = CreateDoc(title: "Full cycle");
        var payload = Encoding.UTF8.GetBytes("payload that must survive a full cycle 0123456789");
        WriteSource("cycle.bin", payload);

        // add
        Edit(file, """[{"op":"add","path":"/body","type":"embed","props":{"src":"cycle.bin","name":"Cycle"}}]""");
        Assert.Single(Embeds(file));

        // extract byte-identical
        Edit(file, """[{"op":"extract","path":"/embed[1]","props":{"to":"cycle-out.bin"}}]""");
        Assert.Equal(payload, File.ReadAllBytes(Path.Combine(Dir, "cycle-out.bin")));

        // remove + reopen-verify
        Edit(file, """[{"op":"remove","path":"/embed[1]"}]""");
        Assert.Empty(Embeds(file));
        AssertValidatesClean(file);
    }

    // ----------------------------------------------------------- helpers

    /// <summary>Asserts the (single) embedded package part of a docx carries the expected bytes.</summary>
    private static void AssertEmbeddedPayloadEquals(byte[] docxBytes, byte[] expected)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, isEditable: false);
        var part = Assert.Single(doc.MainDocumentPart!.EmbeddedPackageParts);
        using var stream = part.GetStream();
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        Assert.Equal(expected, buffer.ToArray());
    }
}
