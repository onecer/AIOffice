using System.Text.Json.Nodes;
using AIOffice.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace AIOffice.Pptx.Tests;

/// <summary>
/// M10 embedded objects: a file embedded on a slide as an OLE/package object via
/// a p:graphicFrame referencing an EmbeddedPackagePart. The tests cover the whole
/// surface — add, read --view embeds, get, extract (byte-identical), remove,
/// reopen-verify, sandbox denial on src and dest, validator-clean, the convert
/// Dropped note, media-type sniffing and addressing — all platform-independent.
/// </summary>
public sealed class EmbedTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly PptxHandler _handler = new();

    public void Dispose() => _ws.Dispose();

    private void Create() =>
        TestEnv.AssertOk(_handler.Create(_ws.Ctx("deck.pptx", ("title", "Report"))));

    private JsonObject Edit(params EditOp[] ops) =>
        TestEnv.AssertOk(_handler.Edit(_ws.Ctx("deck.pptx"), ops));

    /// <summary>Writes a payload file into the workspace and returns its workspace-relative name.</summary>
    private string WritePayload(string name, byte[] bytes)
    {
        File.WriteAllBytes(_ws.PathOf(name), bytes);
        return name;
    }

    // ---- add / list / get ---------------------------------------------------

    [Fact]
    public void AddEmbed_EmbedsPackagePart_AndReturnsCanonicalIdPath()
    {
        Create();
        var payload = SampleXlsx();
        WritePayload("data.xlsx", payload);

        var data = Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(
            ("src", "data.xlsx"),
            ("name", "Q3 numbers"))));
        var canonical = data["results"]![0]!["target"]!.GetValue<string>();
        Assert.Matches(@"^/slide\[1\]/embed\[@id=[0-9]+\]$", canonical);

        using var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false);
        var slidePart = doc.PresentationPart!.SlideParts.Single();
        var pkg = Assert.Single(slidePart.Parts.Select(p => p.OpenXmlPart).OfType<EmbeddedPackagePart>());
        Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", pkg.ContentType);
        using var stream = pkg.GetStream();
        Assert.Equal(payload.Length, stream.Length); // bytes embedded verbatim

        // The wire form is an OLE graphicFrame whose p:oleObject points at the part.
        var ole = Assert.Single(slidePart.Slide!.Descendants<P.OleObject>());
        Assert.Equal(slidePart.GetIdOfPart(pkg), ole.Id!.Value);
        Assert.Equal("Q3 numbers", ole.Name!.Value);

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    [Fact]
    public void ReadEmbedsView_ListsEveryEmbed_WithMetadataNotBytes()
    {
        Create();
        WritePayload("data.xlsx", SampleXlsx());
        WritePayload("notes.txt", "hello"u8.ToArray());

        Edit(
            TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "data.xlsx"), ("name", "Workbook"))),
            TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "notes.txt"))));

        var view = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "embeds"))));
        Assert.Equal("embeds", view["view"]!.GetValue<string>());
        Assert.Equal(2, view["count"]!.GetValue<int>());

        var embeds = view["embeds"]!.AsArray();
        Assert.Equal("Workbook", embeds[0]!["name"]!.GetValue<string>());
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            embeds[0]!["mediaType"]!.GetValue<string>());
        Assert.Equal("/slide[1]", embeds[0]!["container"]!.GetValue<string>());
        Assert.True(embeds[0]!["size"]!.GetValue<long>() > 0);

        // The .txt name falls back to the file name; its media type is text/plain.
        Assert.Equal("notes.txt", embeds[1]!["name"]!.GetValue<string>());
        Assert.Equal("text/plain", embeds[1]!["mediaType"]!.GetValue<string>());
    }

    [Fact]
    public void GetEmbedPath_ReturnsMetadata_NotBytes()
    {
        Create();
        var payload = SampleXlsx();
        WritePayload("data.xlsx", payload);
        var canonical = Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(
            ("src", "data.xlsx"), ("name", "Sheet"))))["results"]![0]!["target"]!.GetValue<string>();

        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", canonical))));
        Assert.Equal("Sheet", detail["name"]!.GetValue<string>());
        Assert.Equal(payload.Length, detail["size"]!.GetValue<long>());
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            detail["mediaType"]!.GetValue<string>());
        Assert.Null(detail["bytes"]); // metadata only, never the payload
    }

    [Fact]
    public void StructureView_IncludesSlideEmbeds()
    {
        Create();
        WritePayload("data.xlsx", SampleXlsx());
        Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "data.xlsx"), ("name", "Book"))));

        var structure = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "structure"))));
        var slide = structure["slides"]!.AsArray()[0]!;
        var embeds = slide["embeds"]!.AsArray();
        Assert.Equal("Book", Assert.Single(embeds)!["name"]!.GetValue<string>());
    }

    [Fact]
    public void AddEmbed_WithIconPreview_EmbedsImagePart_ValidatorClean()
    {
        Create();
        WritePayload("data.xlsx", SampleXlsx());
        WritePayload("preview.png", TestImages.Png(32, 32));

        var canonical = Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(
            ("src", "data.xlsx"),
            ("icon", "preview.png"))))["results"]![0]!["target"]!.GetValue<string>();
        Assert.Matches(@"^/slide\[1\]/embed\[@id=[0-9]+\]$", canonical);

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            // The icon is a real ImagePart; the package payload is still the embed.
            Assert.Single(slidePart.ImageParts);
            Assert.Single(slidePart.Parts.Select(p => p.OpenXmlPart).OfType<EmbeddedPackagePart>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- extract ------------------------------------------------------------

    [Fact]
    public void ExtractEmbed_WritesByteIdenticalPayload_WithoutChangingDeck()
    {
        Create();
        var payload = SampleXlsx();
        WritePayload("data.xlsx", payload);
        var canonical = Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(
            ("src", "data.xlsx"))))["results"]![0]!["target"]!.GetValue<string>();

        var deckBefore = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        Edit(TestEnv.Op("extract", canonical, props: TestEnv.Props(("to", "out/extracted.xlsx"))));

        Assert.Equal(payload, File.ReadAllBytes(_ws.PathOf("out/extracted.xlsx")));
        // extract reads, never writes the source deck.
        Assert.Equal(deckBefore, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    [Fact]
    public void ExtractEmbed_ByOrdinalPath_AlsoWorks()
    {
        Create();
        var payload = SampleXlsx();
        WritePayload("data.xlsx", payload);
        Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "data.xlsx"))));

        Edit(TestEnv.Op("extract", "/slide[1]/embed[1]", props: TestEnv.Props(("to", "got.xlsx"))));

        Assert.Equal(payload, File.ReadAllBytes(_ws.PathOf("got.xlsx")));
    }

    /// <summary>The byte-survival law: bytes survive the open+save round-trip and re-extract identically.</summary>
    [Fact]
    public void ExtractAfterReopen_ReturnsIdenticalBytes()
    {
        Create();
        var payload = SampleXlsx();
        WritePayload("data.xlsx", payload);
        Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "data.xlsx"))));

        // A second, unrelated mutating op forces another full open+save of the deck.
        Edit(TestEnv.Op("add", "/slide[2]", type: "slide"));

        var list = ((IEmbedHost)_handler).ListEmbeds(_ws.Ctx("deck.pptx"));
        var dest = _ws.PathOf("after-reopen.xlsx");
        ((IEmbedHost)_handler).ExtractEmbed(_ws.Ctx("deck.pptx"), list[0].Path, dest);

        Assert.Equal(payload, File.ReadAllBytes(dest));
    }

    // ---- remove -------------------------------------------------------------

    [Fact]
    public void RemoveEmbed_DropsFrameAndPayloadPart_ValidatorClean()
    {
        Create();
        WritePayload("data.xlsx", SampleXlsx());
        var canonical = Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(
            ("src", "data.xlsx"))))["results"]![0]!["target"]!.GetValue<string>();

        Edit(TestEnv.Op("remove", canonical));

        var view = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "embeds"))));
        Assert.Equal(0, view["count"]!.GetValue<int>());

        using (var doc = PresentationDocument.Open(_ws.PathOf("deck.pptx"), false))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            Assert.Empty(slidePart.Parts.Select(p => p.OpenXmlPart).OfType<EmbeddedPackagePart>()); // payload not orphaned
            Assert.Empty(slidePart.Slide!.Descendants<P.OleObject>());
        }

        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- sandbox ------------------------------------------------------------

    [Fact]
    public void EmbedSrcOutsideWorkspace_IsSandboxDenied()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "../escapee.xlsx"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.SandboxDenied);
        Assert.Equal(ExitCodes.SandboxDenied, envelope.ExitCode);
    }

    [Fact]
    public void ExtractDestOutsideWorkspace_IsSandboxDenied_AndWritesNothing()
    {
        Create();
        WritePayload("data.xlsx", SampleXlsx());
        var canonical = Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(
            ("src", "data.xlsx"))))["results"]![0]!["target"]!.GetValue<string>();
        var deckBefore = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("extract", canonical, props: TestEnv.Props(("to", "../escape.xlsx"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.SandboxDenied);
        Assert.Equal(deckBefore, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    [Fact]
    public void MissingSrc_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("name", "x"))),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("src", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MissingSrcFile_IsFileNotFound()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "ghost.xlsx"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.FileNotFound);
    }

    [Fact]
    public void ExtractWithoutTo_IsInvalidArgs()
    {
        Create();
        WritePayload("data.xlsx", SampleXlsx());
        var canonical = Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(
            ("src", "data.xlsx"))))["results"]![0]!["target"]!.GetValue<string>();

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("extract", canonical),
        ]);

        var error = TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains("to", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractNonEmbedPath_IsInvalidArgs()
    {
        Create();
        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("extract", "/slide[1]", props: TestEnv.Props(("to", "x.bin"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.InvalidArgs);
    }

    // ---- media-type sniffing ------------------------------------------------

    [Fact]
    public void MediaType_IsSniffedByHeader_NotExtension()
    {
        Create();
        // A PDF whose extension lies: the "%PDF-" header decides.
        WritePayload("disguised.dat", [.. "%PDF-1.4"u8, (byte)'\n', (byte)'x']);

        Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "disguised.dat"))));

        var view = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "embeds"))));
        Assert.Equal("application/pdf", view["embeds"]![0]!["mediaType"]!.GetValue<string>());
    }

    [Fact]
    public void UnknownBinary_FallsBackToOctetStream()
    {
        Create();
        WritePayload("blob.bin", [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07]);

        Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "blob.bin"))));

        var view = TestEnv.AssertOk(_handler.Read(_ws.Ctx("deck.pptx", ("view", "embeds"))));
        Assert.Equal("application/octet-stream", view["embeds"]![0]!["mediaType"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- convert honesty ----------------------------------------------------

    [Fact]
    public void Convert_ExportNeutral_ReportsEmbedAsDroppedNote()
    {
        Create();
        WritePayload("data.xlsx", SampleXlsx());
        Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "data.xlsx"))));

        _ = _handler.ExportNeutral(_ws.Ctx("deck.pptx"), out var dropped);
        Assert.Contains(dropped, d => d.Contains("embedded object", StringComparison.OrdinalIgnoreCase));
    }

    // ---- atomicity ----------------------------------------------------------

    [Fact]
    public void FailedEmbedOp_IsAtomic_NoWrite()
    {
        Create();
        WritePayload("data.xlsx", SampleXlsx());
        var before = File.ReadAllBytes(_ws.PathOf("deck.pptx"));

        var envelope = _handler.Edit(_ws.Ctx("deck.pptx"), [
            TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "data.xlsx"))),
            TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(("src", "../outside.xlsx"))),
        ]);

        TestEnv.AssertFail(envelope, ErrorCodes.SandboxDenied);
        Assert.Equal(before, File.ReadAllBytes(_ws.PathOf("deck.pptx")));
    }

    [Fact]
    public void TwoEmbedsOnOneSlide_KeepStableIdPaths()
    {
        Create();
        WritePayload("a.xlsx", SampleXlsx());
        WritePayload("b.txt", "second"u8.ToArray());
        var first = Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(
            ("src", "a.xlsx"))))["results"]![0]!["target"]!.GetValue<string>();
        var second = Edit(TestEnv.Op("add", "/slide[1]", type: "embed", props: TestEnv.Props(
            ("src", "b.txt"))))["results"]![0]!["target"]!.GetValue<string>();
        Assert.NotEqual(first, second);

        // Removing the first leaves the second addressable by its stable id.
        Edit(TestEnv.Op("remove", first));
        var detail = TestEnv.AssertOk(_handler.Get(_ws.Ctx("deck.pptx", ("path", second))));
        Assert.Equal("text/plain", detail["mediaType"]!.GetValue<string>());
        TestEnv.AssertValid(_ws, "deck.pptx");
    }

    // ---- fixtures -----------------------------------------------------------

    /// <summary>A small but genuinely valid .xlsx package used as an embed payload.</summary>
    private static byte[] SampleXlsx()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell(new InlineString(new Text("hello"))) { CellReference = "A1", DataType = CellValues.InlineString })
                { RowIndex = 1u }));
            workbookPart.Workbook = new Workbook(new Sheets(new Sheet
            {
                Name = "Sheet1",
                SheetId = 1u,
                Id = workbookPart.GetIdOfPart(worksheetPart),
            }));
        }

        return stream.ToArray();
    }
}
