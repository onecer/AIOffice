using System.IO.Compression;
using System.Text.Json.Nodes;
using AIOffice.Core;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace AIOffice.Excel.Tests;

/// <summary>
/// The M10 embedded-objects slice: any file embedded as an
/// <see cref="EmbeddedPackagePart"/> on a worksheet, with display/source/anchor
/// metadata in a workbook custom property. Covers embed/list/extract
/// (byte-identical)/remove, reopen-verify, sandbox denial on src AND dest, the
/// validator oracle, and the round-trip law (payload survives open+save exactly).
/// </summary>
public sealed class EmbedTests : ExcelTestBase
{
    /// <summary>A tiny ZIP payload (so the media sniff lands on the zip family).</summary>
    private static byte[] ZipBytes()
    {
        using var memory = new MemoryStream();
        using (var archive = new ZipArchive(memory, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = archive.CreateEntry("hello.txt");
            using var writer = new StreamWriter(entry.Open());
            writer.Write("payload contents");
        }

        return memory.ToArray();
    }

    /// <summary>A small valid .xlsx the user might attach as a source workbook.</summary>
    private byte[] SourceXlsxBytes()
    {
        var path = Path.Combine(Dir, "source-" + Guid.NewGuid().ToString("N") + ".xlsx");
        using (var workbook = new XLWorkbook())
        {
            workbook.AddWorksheet("Data").Cell("A1").Value = "embedded";
            workbook.SaveAs(path);
        }

        var bytes = File.ReadAllBytes(path);
        File.Delete(path);
        return bytes;
    }

    private string WriteWorkspaceFile(string name, byte[] bytes)
    {
        File.WriteAllBytes(Path.Combine(Dir, name), bytes);
        return name; // relative to the workspace root — the sandbox resolves it
    }

    private static EditOp ExtractOp(string path, string to) => new()
    {
        Op = "extract",
        Path = path,
        Props = new JsonObject { ["to"] = to },
    };

    /// <summary>The first sheet's embedded package parts, read raw with the OpenXml SDK.</summary>
    private static List<EmbeddedPackagePart> RawEmbeds(SpreadsheetDocument document, string sheetName = "Sheet1")
    {
        var workbookPart = document.WorkbookPart!;
        var sheet = workbookPart.Workbook!.Descendants<S.Sheet>()
            .First(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return worksheetPart.EmbeddedPackageParts.OrderBy(p => p.Uri.ToString(), StringComparer.Ordinal).ToList();
    }

    private static byte[] RawEmbedBytes(string file, int index0, string sheetName = "Sheet1")
    {
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        var part = RawEmbeds(document, sheetName)[index0];
        using var stream = part.GetStream();
        var bytes = new byte[stream.Length];
        stream.ReadExactly(bytes);
        return bytes;
    }

    // ----- add / list ---------------------------------------------------------

    [Fact]
    public void Add_embed_stores_payload_and_is_validator_clean()
    {
        var file = CreateWorkbook();
        var payload = SourceXlsxBytes();
        var src = WriteWorkspaceFile("report.xlsx", payload);

        var envelope = EditOps(file, AddOp("/Sheet1", "embed", ("src", src), ("name", "Q3 source")));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal("/Sheet1/embed[1]", detail["path"]!.GetValue<string>());
        Assert.Equal("embed", detail["type"]!.GetValue<string>());
        Assert.Equal("Q3 source", detail["name"]!.GetValue<string>());
        Assert.Equal(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            detail["mediaType"]!.GetValue<string>());

        AssertValidatorClean(file);

        // The package part holds exactly the source bytes.
        Assert.Equal(payload, RawEmbedBytes(file, 0));
    }

    [Fact]
    public void Default_name_is_the_source_file_name()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("budget.zip", ZipBytes());

        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/embed[1]"))));
        Assert.Equal("budget.zip", data["name"]!.GetValue<string>());
        Assert.Equal("application/zip", data["mediaType"]!.GetValue<string>());
        Assert.Equal("budget.zip", data["source"]!.GetValue<string>());
    }

    [Fact]
    public void Media_type_is_sniffed_by_header_not_extension()
    {
        var file = CreateWorkbook();
        // Real ZIP bytes behind a misleading .bin extension: sniff wins.
        var src = WriteWorkspaceFile("mystery.bin", ZipBytes());

        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/embed[1]"))));
        Assert.Equal("application/zip", data["mediaType"]!.GetValue<string>());
    }

    [Fact]
    public void Anchor_is_recorded_and_round_trips()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("doc.zip", ZipBytes());

        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src), ("anchor", "e2"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/embed[1]"))));
        Assert.Equal("E2", data["anchor"]!.GetValue<string>()); // normalized uppercase
    }

    [Fact]
    public void Bad_anchor_is_invalid_args()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("doc.zip", ZipBytes());

        var envelope = EditOps(file, AddOp("/Sheet1", "embed", ("src", src), ("anchor", "not-a-cell")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    [Fact]
    public void Read_view_embeds_lists_every_embed_in_path_order()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("a.zip", ZipBytes());
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "embed", ("src", src), ("name", "First")),
            AddOp("/Sheet1", "embed", ("src", src), ("name", "Second"))).IsOk);

        var data = OkData(Handler.Read(Ctx(file, ("view", "embeds"))));

        var embeds = data["embeds"]!.AsArray();
        Assert.Equal(2, embeds.Count);
        Assert.Equal("/Sheet1/embed[1]", embeds[0]!["path"]!.GetValue<string>());
        Assert.Equal("First", embeds[0]!["name"]!.GetValue<string>());
        Assert.Equal("/Sheet1/embed[2]", embeds[1]!["path"]!.GetValue<string>());
        Assert.Equal("Second", embeds[1]!["name"]!.GetValue<string>());
        Assert.Equal(2, data["totals"]!["count"]!.GetValue<int>());
    }

    [Fact]
    public void Structure_view_lists_embeds_per_sheet_and_sheet_get_counts_them()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("a.zip", ZipBytes());
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src), ("name", "One"))).IsOk);

        var structure = OkData(Handler.Read(Ctx(file, ("view", "structure"))));
        var embeds = structure["sheets"]!.AsArray()[0]!["embeds"]!.AsArray();
        Assert.Single(embeds);
        Assert.Equal("One", embeds[0]!["name"]!.GetValue<string>());

        // Polish: get on a sheet surfaces an embed count.
        var sheet = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1"))));
        Assert.Equal(1, sheet["embeds"]!.GetValue<int>());
    }

    [Fact]
    public void Get_on_an_embed_returns_metadata_not_bytes()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("a.zip", ZipBytes());
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src), ("name", "Doc"))).IsOk);

        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/embed[1]"))));

        Assert.Equal("/Sheet1/embed[1]", data["path"]!.GetValue<string>());
        Assert.Equal("embed", data["kind"]!.GetValue<string>());
        Assert.Equal("Doc", data["name"]!.GetValue<string>());
        Assert.Equal((long)ZipBytes().Length, data["size"]!.GetValue<long>());
        Assert.Null(data["bytes"]); // metadata only — never the payload
    }

    [Fact]
    public void Get_missing_embed_is_invalid_path_with_candidates()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("a.zip", ZipBytes());
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src))).IsOk);

        var envelope = Handler.Get(Ctx(file, ("path", "/Sheet1/embed[2]")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1/embed[1]", envelope.Error.Candidates!);
    }

    // ----- extract ------------------------------------------------------------

    [Fact]
    public void Extract_writes_byte_identical_payload_and_leaves_source_unchanged()
    {
        var file = CreateWorkbook();
        var payload = SourceXlsxBytes();
        var src = WriteWorkspaceFile("report.xlsx", payload);
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src), ("name", "Src"))).IsOk);

        var sourceRevBefore = Rev.OfFile(file);

        var envelope = EditOps(file, ExtractOp("/Sheet1/embed[1]", "out/extracted.xlsx"));

        Assert.True(envelope.IsOk, envelope.ToJson());
        var detail = Json(envelope)["data"]!["ops"]![0]!;
        Assert.Equal("/Sheet1/embed[1]", detail["path"]!.GetValue<string>());
        Assert.Equal("embed", detail["extracted"]!.GetValue<string>());

        var extracted = File.ReadAllBytes(Path.Combine(Dir, "out", "extracted.xlsx"));
        Assert.Equal(payload, extracted); // byte-identical

        // The source workbook is untouched (extract does not modify the document).
        Assert.Equal(sourceRevBefore, Rev.OfFile(file));
        AssertValidatorClean(file);
    }

    [Fact]
    public void Extract_after_reopen_save_returns_identical_bytes()
    {
        var file = CreateWorkbook();
        var payload = SourceXlsxBytes();
        var src = WriteWorkspaceFile("report.xlsx", payload);
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src))).IsOk);

        // A no-edit ClosedXML resave between embed and extract: the payload must
        // survive the round-trip exactly.
        using (var workbook = new XLWorkbook(file))
        {
            workbook.Save();
        }

        AssertValidatorClean(file);
        Assert.True(EditOps(file, ExtractOp("/Sheet1/embed[1]", "got.xlsx")).IsOk);
        Assert.Equal(payload, File.ReadAllBytes(Path.Combine(Dir, "got.xlsx")));
    }

    [Fact]
    public void Extract_via_embed_host_interface_matches_the_op()
    {
        var file = CreateWorkbook();
        var payload = ZipBytes();
        var src = WriteWorkspaceFile("a.zip", payload);
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src))).IsOk);

        // The IEmbedHost surface the verb layer (read --view embeds / extract) uses.
        var host = (IEmbedHost)Handler;
        var dest = Workspace.Resolve("host-out.zip");
        host.ExtractEmbed(Ctx(file), "/Sheet1/embed[1]", dest);
        Assert.Equal(payload, File.ReadAllBytes(dest));

        var listed = host.ListEmbeds(Ctx(file));
        var one = Assert.Single(listed);
        Assert.Equal("/Sheet1/embed[1]", one.Path);
        Assert.Equal("/Sheet1", one.Container);
        Assert.Equal((long)payload.Length, one.Size);
    }

    [Fact]
    public void Extract_missing_embed_is_invalid_path()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, ExtractOp("/Sheet1/embed[1]", "out.zip"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
    }

    [Fact]
    public void Extract_needs_a_destination()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("a.zip", ZipBytes());
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src))).IsOk);

        var envelope = EditOps(file, new EditOp { Op = "extract", Path = "/Sheet1/embed[1]" });

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
    }

    // ----- remove -------------------------------------------------------------

    [Fact]
    public void Remove_embed_shifts_indices_and_stays_clean()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("a.zip", ZipBytes());
        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "embed", ("src", src), ("name", "One")),
            AddOp("/Sheet1", "embed", ("src", src), ("name", "Two"))).IsOk);

        var remove = EditOps(file, RemoveOp("/Sheet1/embed[1]"));
        Assert.True(remove.IsOk, remove.ToJson());
        AssertValidatorClean(file);

        // "Two" shifted into [1]; only one embed remains.
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/embed[1]"))));
        Assert.Equal("Two", data["name"]!.GetValue<string>());
        Assert.Single(OkData(Handler.Read(Ctx(file, ("view", "embeds"))))["embeds"]!.AsArray());

        Assert.True(EditOps(file, RemoveOp("/Sheet1/embed[1]")).IsOk);
        AssertValidatorClean(file);
        using var document = SpreadsheetDocument.Open(file, isEditable: false);
        Assert.Empty(RawEmbeds(document));
    }

    [Fact]
    public void Remove_missing_embed_is_invalid_path_with_candidates()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("a.zip", ZipBytes());
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src))).IsOk);

        var envelope = EditOps(file, RemoveOp("/Sheet1/embed[9]"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("/Sheet1/embed[1]", envelope.Error.Candidates!);
    }

    [Fact]
    public void Set_on_an_embed_is_a_typed_unsupported_feature()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("a.zip", ZipBytes());
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src))).IsOk);

        var envelope = EditOps(file, SetOp("/Sheet1/embed[1]", ("value", 1)));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.UnsupportedFeature, envelope.Error!.Code);
        Assert.Contains("remove", envelope.Error.Suggestion, StringComparison.OrdinalIgnoreCase);
    }

    // ----- sandbox ------------------------------------------------------------

    [Fact]
    public void Src_escaping_the_sandbox_is_denied_without_reading_the_file()
    {
        var file = CreateWorkbook();
        var outside = Path.Combine(Path.GetTempPath(), "aioffice-outside-" + Guid.NewGuid().ToString("N") + ".zip");
        File.WriteAllBytes(outside, ZipBytes());
        try
        {
            foreach (var src in new[] { outside, "../" + Path.GetFileName(outside) })
            {
                var envelope = EditOps(file, AddOp("/Sheet1", "embed", ("src", src)));
                Assert.False(envelope.IsOk);
                Assert.Equal(ErrorCodes.SandboxDenied, envelope.Error!.Code);
            }
        }
        finally
        {
            File.Delete(outside);
        }
    }

    [Fact]
    public void Extract_dest_escaping_the_sandbox_is_denied()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("a.zip", ZipBytes());
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src))).IsOk);

        var outside = Path.Combine(Path.GetTempPath(), "aioffice-escape-" + Guid.NewGuid().ToString("N") + ".zip");
        foreach (var dest in new[] { outside, "../escape.zip" })
        {
            var envelope = EditOps(file, ExtractOp("/Sheet1/embed[1]", dest));
            Assert.False(envelope.IsOk);
            Assert.Equal(ErrorCodes.SandboxDenied, envelope.Error!.Code);
        }

        Assert.False(File.Exists(outside)); // nothing was written outside the sandbox
    }

    [Fact]
    public void Missing_src_file_is_file_not_found()
    {
        var file = CreateWorkbook();

        var envelope = EditOps(file, AddOp("/Sheet1", "embed", ("src", "nope.zip")));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.FileNotFound, envelope.Error!.Code);
    }

    // ----- survival across an unrelated edit ----------------------------------

    [Fact]
    public void Embed_survives_a_later_unrelated_edit_byte_identical()
    {
        var file = CreateWorkbook();
        var payload = SourceXlsxBytes();
        var src = WriteWorkspaceFile("report.xlsx", payload);
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src), ("name", "Keep"))).IsOk);

        // A normal cell edit goes through ClosedXML; the embed and its registry
        // must survive byte-identical.
        Assert.True(EditOps(file, SetOp("/Sheet1/A1", ("value", "touched"))).IsOk);

        AssertValidatorClean(file);
        var data = OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/embed[1]"))));
        Assert.Equal("Keep", data["name"]!.GetValue<string>());
        Assert.Equal(payload, RawEmbedBytes(file, 0));
    }

    [Fact]
    public void Embed_registry_coexists_with_the_cell_style_registry()
    {
        var file = CreateWorkbook();
        var src = WriteWorkspaceFile("a.zip", ZipBytes());

        // A named cell style also lives in a workbook custom property; the embed
        // registry must not clobber it (both land in docProps/custom.xml).
        Assert.True(EditOps(file, AddOp("/styles", "cellStyle",
            ("name", "Currency-Red"), ("numberFormat", "$#,##0.00"), ("color", "#FF0000"))).IsOk);
        Assert.True(EditOps(file, AddOp("/Sheet1", "embed", ("src", src), ("name", "Doc"))).IsOk);

        AssertValidatorClean(file);

        // Both registries are intact and readable.
        var styles = OkData(Handler.Read(Ctx(file, ("view", "styles"))))["styles"]!.AsArray();
        Assert.Contains(styles, s => s!["name"]!.GetValue<string>() == "Currency-Red");
        var embeds = OkData(Handler.Read(Ctx(file, ("view", "embeds"))))["embeds"]!.AsArray();
        Assert.Equal("Doc", Assert.Single(embeds)!["name"]!.GetValue<string>());
    }

    [Fact]
    public void Embeds_on_two_sheets_are_indexed_independently()
    {
        var file = CreateWorkbook();
        Assert.True(EditOps(file, AddOp("/Two", "sheet")).IsOk);
        var src = WriteWorkspaceFile("a.zip", ZipBytes());

        Assert.True(EditOps(
            file,
            AddOp("/Sheet1", "embed", ("src", src), ("name", "S1")),
            AddOp("/Two", "embed", ("src", src), ("name", "S2"))).IsOk);

        Assert.Equal("S1", OkData(Handler.Get(Ctx(file, ("path", "/Sheet1/embed[1]"))))["name"]!.GetValue<string>());
        Assert.Equal("S2", OkData(Handler.Get(Ctx(file, ("path", "/Two/embed[1]"))))["name"]!.GetValue<string>());
    }
}
