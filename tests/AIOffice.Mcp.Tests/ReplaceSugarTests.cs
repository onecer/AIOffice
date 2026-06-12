using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using AIOffice.Excel;
using AIOffice.Pptx;
using AIOffice.Word;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// The M4 document-wide find/replace sugar shared by the CLI
/// (<c>--find/--replace</c>) and MCP (<c>office_edit</c> replace op with
/// path <c>"/"</c>): a root-scoped replace op fans out over the format's
/// default scopes — docx body + every header/footer, every sheet, every slide
/// including its notes — and the per-scope results are folded into one
/// document-level <c>{replacements, locations}</c> aggregate with collapsed
/// <c>find_no_match</c> warnings.
/// </summary>
public sealed class ReplaceSugarTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly CommandService _service;

    public ReplaceSugarTests()
    {
        var snapshots = new SnapshotStore(_ws.SnapshotDir);
        var registry = new HandlerRegistry();
        registry.Register(new WordHandler(snapshots), ".docx");
        registry.Register(new ExcelHandler(snapshots), ".xlsx");
        registry.Register(new PptxHandler(snapshots), ".pptx");
        _service = _ws.NewService(
            registry,
            new HashSet<DocumentKind> { DocumentKind.Docx, DocumentKind.Xlsx, DocumentKind.Pptx });
    }

    public void Dispose() => _ws.Dispose();

    // ---------------------------------------------------------------- helpers

    private JsonObject Ok(Envelope envelope)
    {
        Assert.True(envelope.IsOk, $"expected ok, got: {envelope.Error?.Code} {envelope.Error?.Message}");
        return (JsonObject)JsonSerializer.SerializeToNode(envelope.Data!, JsonDefaults.Options)!;
    }

    private Envelope Edit(string file, JsonArray ops, bool track = false) =>
        _service.Edit(new JsonObject { ["file"] = file, ["ops"] = ops, ["track"] = track });

    private static JsonObject Op(string op, string path, string? type = null, JsonObject? props = null)
    {
        var node = new JsonObject { ["op"] = op, ["path"] = path };
        if (type is not null)
        {
            node["type"] = type;
        }

        if (props is not null)
        {
            node["props"] = props;
        }

        return node;
    }

    private static JsonObject RootReplace(string find, string replace, bool regex = false) =>
        Op("replace", "/", props: new JsonObject
        {
            ["find"] = find,
            ["replace"] = replace,
            ["regex"] = regex,
        });

    private static IReadOnlyList<Warning> WarningsOf(Envelope envelope) => envelope.Meta.Warnings ?? [];

    // ------------------------------------------------------------------- docx

    private void BuildReportDocx()
    {
        Ok(_service.Create(new JsonObject { ["file"] = "report.docx" }));
        Ok(Edit("report.docx", new JsonArray(
            Op("set", "/body/p[1]", props: new JsonObject { ["text"] = "Budget 2025 intro" }),
            Op("add", "/body", "p", new JsonObject { ["text"] = "Plain middle paragraph" }),
            Op("add", "/body", "p", new JsonObject { ["text"] = "2025 closing, 2025 twice" }),
            Op("add", "/header[1]", "header", new JsonObject { ["text"] = "FY 2025 header" }),
            Op("add", "/footer[1]", "footer", new JsonObject { ["text"] = "footer for 2025" }))));
    }

    [Fact]
    public void Docx_root_replace_covers_body_headers_and_footers_and_aggregates()
    {
        BuildReportDocx();

        var envelope = Edit("report.docx", new JsonArray(RootReplace("2025", "2026")));
        var data = Ok(envelope);

        // 2 body hits in p[3] + 1 in p[1] + 1 header + 1 footer = 5 total.
        Assert.Equal(5, data["replacements"]!.GetValue<int>());
        var locations = data["locations"]!.AsArray().Select(l => l!.GetValue<string>()).ToList();
        Assert.Contains("/body/p[1]", locations);
        Assert.Contains("/header[1]/p[1]", locations);
        Assert.Contains("/footer[1]/p[1]", locations);

        // Something matched -> the per-scope find_no_match warnings are gone.
        Assert.DoesNotContain(WarningsOf(envelope), w => w.Code == ErrorCodes.FindNoMatch);

        var header = Ok(_service.Get(new JsonObject { ["file"] = "report.docx", ["path"] = "/header[1]/p[1]" }));
        Assert.Equal("FY 2026 header", header["properties"]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void Docx_tracked_root_replace_stays_body_scoped_and_records_revisions()
    {
        BuildReportDocx();

        var data = Ok(Edit("report.docx", new JsonArray(RootReplace("2025", "2026")), track: true));

        // Tracked replace is body-only by contract: 3 body hits, header/footer untouched.
        Assert.Equal(3, data["replacements"]!.GetValue<int>());

        var revisions = Ok(_service.Read(new JsonObject { ["file"] = "report.docx", ["view"] = "revisions" }));
        Assert.Equal(6, revisions["count"]!.GetValue<int>()); // 3 del + 3 ins pairs

        var header = Ok(_service.Get(new JsonObject { ["file"] = "report.docx", ["path"] = "/header[1]/p[1]" }));
        Assert.Equal("FY 2025 header", header["properties"]!["text"]!.GetValue<string>());
    }

    [Fact]
    public void No_match_anywhere_is_ok_with_one_document_level_warning()
    {
        BuildReportDocx();
        var before = Rev.OfFile(_ws.Workspace.Resolve("report.docx"));

        var envelope = Edit("report.docx", new JsonArray(RootReplace("absent text", "x")));
        var data = Ok(envelope);

        Assert.Equal(0, data["replacements"]!.GetValue<int>());
        Assert.Empty(data["locations"]!.AsArray());

        var warning = Assert.Single(WarningsOf(envelope), w => w.Code == ErrorCodes.FindNoMatch);
        Assert.Contains("absent text", warning.Message, StringComparison.Ordinal);
        Assert.Equal(before, Rev.OfFile(_ws.Workspace.Resolve("report.docx"))); // nothing changed... rev-wise
    }

    // ------------------------------------------------------------------- xlsx

    [Fact]
    public void Xlsx_root_replace_covers_every_sheet()
    {
        Ok(_service.Create(new JsonObject { ["file"] = "book.xlsx" }));
        Ok(Edit("book.xlsx", new JsonArray(
            Op("set", "/Sheet1/A1", props: new JsonObject { ["value"] = "draft note" }),
            Op("add", "/'Q3 Data'", "sheet"),
            Op("set", "/'Q3 Data'/B2", props: new JsonObject { ["value"] = "another draft" }))));

        var data = Ok(Edit("book.xlsx", new JsonArray(RootReplace("draft", "final"))));

        Assert.Equal(2, data["replacements"]!.GetValue<int>());
        var locations = data["locations"]!.AsArray().Select(l => l!.GetValue<string>()).ToList();
        Assert.Contains("/Sheet1/A1", locations);
        Assert.Contains("/'Q3 Data'/B2", locations);

        var cell = Ok(_service.Get(new JsonObject { ["file"] = "book.xlsx", ["path"] = "/'Q3 Data'/B2" }));
        Assert.Equal("another final", cell["value"]!.GetValue<string>());
    }

    // ------------------------------------------------------------------- pptx

    [Fact]
    public void Pptx_root_replace_covers_every_slide_including_notes()
    {
        Ok(_service.Create(new JsonObject { ["file"] = "deck.pptx" }));
        Ok(Edit("deck.pptx", new JsonArray(
            Op("add", "/slide[2]", "slide"),
            Op("add", "/slide[1]", "shape", new JsonObject { ["text"] = "Q3 review" }),
            Op("add", "/slide[2]", "shape", new JsonObject { ["text"] = "More Q3 numbers" }),
            Op("set", "/slide[2]/notes", props: new JsonObject { ["text"] = "remember Q3 context" }))));

        var data = Ok(Edit("deck.pptx", new JsonArray(RootReplace("Q3", "Q4"))));

        Assert.Equal(3, data["replacements"]!.GetValue<int>());
        var locations = data["locations"]!.AsArray().Select(l => l!.GetValue<string>()).ToList();
        Assert.Contains(locations, l => l.StartsWith("/slide[1]/", StringComparison.Ordinal));
        Assert.Contains(locations, l => l.StartsWith("/slide[2]/", StringComparison.Ordinal));
        Assert.Contains("/slide[2]/notes", locations);

        var notes = Ok(_service.Get(new JsonObject { ["file"] = "deck.pptx", ["path"] = "/slide[2]/notes" }));
        Assert.Equal("remember Q4 context", notes["text"]!.GetValue<string>());
    }

    // ------------------------------------------------------------------ guard

    [Fact]
    public void Root_path_is_rejected_for_non_replace_ops()
    {
        Ok(_service.Create(new JsonObject { ["file"] = "report.docx" }));

        var envelope = Edit("report.docx", new JsonArray(
            Op("set", "/", props: new JsonObject { ["text"] = "x" })));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.False(string.IsNullOrWhiteSpace(envelope.Error.Suggestion));
    }
}
