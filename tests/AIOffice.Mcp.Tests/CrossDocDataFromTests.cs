using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using AIOffice.Excel;
using AIOffice.Pptx;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// The M3 cross-document bridge: a pptx <c>add chart</c> op carrying
/// <c>props.dataFrom = "metrics.xlsx!Sheet1/A1:C5"</c> is expanded at the
/// command layer (CLI edit and MCP office_edit share it) into literal
/// categories/series read from the REAL workbook via the REAL xlsx handler —
/// first column = categories, header row = series names, remaining columns =
/// series values. Failures must be typed, carry candidates where the spec
/// promises them, and write nothing.
/// </summary>
public sealed class CrossDocDataFromTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly CommandService _service;

    public CrossDocDataFromTests()
    {
        // The xlsx handler manages its own pre-image snapshots; point its ring
        // at the test dir so the suite never touches ~/.aioffice.
        var snapshots = new SnapshotStore(_ws.SnapshotDir);
        var registry = new HandlerRegistry();
        registry.Register(new ExcelHandler(snapshots), ".xlsx");
        registry.Register(new PptxHandler(), ".pptx");
        _service = _ws.NewService(registry, new HashSet<DocumentKind> { DocumentKind.Xlsx });
        BuildMetricsWorkbook();
        EnvelopeOk(_service.Create(new JsonObject { ["file"] = "deck.pptx" }));
    }

    public void Dispose() => _ws.Dispose();

    // ---------------------------------------------------------------- helpers

    private void BuildMetricsWorkbook()
    {
        EnvelopeOk(_service.Create(new JsonObject { ["file"] = "metrics.xlsx" }));
        (string Cell, JsonNode Value)[] cells =
        [
            ("A1", "Quarter"), ("B1", "North"), ("C1", "South"),
            ("A2", "Q1"), ("B2", 10), ("C2", 20),
            ("A3", "Q2"), ("B3", 15), ("C3", 25),
            ("A4", "Q3"), ("B4", 12), ("C4", 22),
            ("A5", "Q4"), ("B5", 18), ("C5", 28),
            ("E2", "note"), // non-numeric stray for the bad-range test
        ];
        var ops = new JsonArray();
        foreach (var (cell, value) in cells)
        {
            ops.Add(new JsonObject
            {
                ["op"] = "set",
                ["path"] = "/Sheet1/" + cell,
                ["props"] = new JsonObject { ["value"] = value },
            });
        }

        EnvelopeOk(_service.Edit(new JsonObject { ["file"] = "metrics.xlsx", ["ops"] = ops }));
    }

    private static JsonObject ChartAddOps(string dataFrom, string kind = "bar") => new()
    {
        ["file"] = "deck.pptx",
        ["ops"] = new JsonArray(new JsonObject
        {
            ["op"] = "add",
            ["path"] = "/slide[1]",
            ["type"] = "chart",
            ["props"] = new JsonObject
            {
                ["kind"] = kind,
                ["title"] = "From workbook",
                ["dataFrom"] = dataFrom,
            },
        }),
    };

    private static JsonObject EnvelopeOk(Envelope envelope)
    {
        Assert.True(envelope.IsOk, $"expected ok, got: {envelope.Error?.Code} {envelope.Error?.Message}");
        return (JsonObject)JsonSerializer.SerializeToNode(envelope.Data!, JsonDefaults.Options)!;
    }

    private string DeckRev() => Rev.OfFile(_ws.Workspace.Resolve("deck.pptx"));

    // ------------------------------------------------------------- happy path

    [Fact]
    public void DataFrom_expands_into_a_real_chart_with_workbook_data()
    {
        var envelope = _service.Edit(ChartAddOps("metrics.xlsx!Sheet1/A1:C5"));
        EnvelopeOk(envelope);

        var chart = EnvelopeOk(_service.Get(new JsonObject { ["file"] = "deck.pptx", ["path"] = "/slide[1]/chart[1]" }));
        Assert.Equal("bar", chart["chartKind"]!.GetValue<string>());
        Assert.Equal("From workbook", chart["title"]!.GetValue<string>());
        Assert.Equal(
            ["Q1", "Q2", "Q3", "Q4"],
            chart["categories"]!.AsArray().Select(c => c!.GetValue<string>()));

        var series = chart["series"]!.AsArray();
        Assert.Equal(2, series.Count);
        Assert.Equal("North", series[0]!["name"]!.GetValue<string>());
        Assert.Equal("South", series[1]!["name"]!.GetValue<string>());
        Assert.Equal(
            [10d, 15d, 12d, 18d],
            series[0]!["values"]!.AsArray().Select(v => v!.GetValue<double>()));

        // The deck still validates after the cross-doc write.
        var valid = EnvelopeOk(_service.Validate(new JsonObject { ["file"] = "deck.pptx" }));
        Assert.True(valid["valid"]!.GetValue<bool>());
    }

    [Fact]
    public void DataFrom_supports_quoted_sheet_names_and_leading_slash()
    {
        var envelope = _service.Edit(ChartAddOps("metrics.xlsx!/'Sheet1'/A1:B5", kind: "line"));
        EnvelopeOk(envelope);

        var chart = EnvelopeOk(_service.Get(new JsonObject { ["file"] = "deck.pptx", ["path"] = "/slide[1]/chart[1]" }));
        var series = Assert.Single(chart["series"]!.AsArray());
        Assert.Equal("North", series!["name"]!.GetValue<string>());
    }

    // ------------------------------------------------------------ failure map

    [Fact]
    public void Wrong_sheet_fails_typed_with_candidates_and_writes_nothing()
    {
        var revBefore = DeckRev();

        var envelope = _service.Edit(ChartAddOps("metrics.xlsx!Sheeet1/A1:C5"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidPath, envelope.Error!.Code);
        Assert.Contains("dataFrom", envelope.Error.Message, StringComparison.Ordinal);
        Assert.NotNull(envelope.Error.Candidates);
        Assert.NotEmpty(envelope.Error.Candidates!);
        Assert.Equal(revBefore, DeckRev());
    }

    [Fact]
    public void Single_cell_fails_with_the_used_range_as_candidate()
    {
        var envelope = _service.Edit(ChartAddOps("metrics.xlsx!Sheet1/B2"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        var candidate = Assert.Single(envelope.Error.Candidates!);
        Assert.Equal("metrics.xlsx!Sheet1/A1:E5", candidate); // the sheet's real used range
    }

    [Fact]
    public void Too_small_range_fails_with_the_used_range_as_candidate()
    {
        var envelope = _service.Edit(ChartAddOps("metrics.xlsx!Sheet1/A1:A5")); // one column = no series

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("at least two columns", envelope.Error.Message, StringComparison.Ordinal);
        Assert.NotEmpty(envelope.Error.Candidates!);
    }

    [Fact]
    public void Non_numeric_series_cell_fails_naming_the_offender()
    {
        var envelope = _service.Edit(ChartAddOps("metrics.xlsx!Sheet1/D1:E5")); // E2 holds "note"

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("not numeric", envelope.Error.Message, StringComparison.Ordinal);
        Assert.Contains("note", envelope.Error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Spec_without_a_bang_fails_with_the_shape_suggestion()
    {
        var envelope = _service.Edit(ChartAddOps("metrics.xlsx Sheet1/A1:C5"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.InvalidArgs, envelope.Error!.Code);
        Assert.Contains("book.xlsx!Sheet1/A1:B5", envelope.Error.Suggestion, StringComparison.Ordinal);
    }

    [Fact]
    public void Workbook_path_is_sandbox_resolved()
    {
        var envelope = _service.Edit(ChartAddOps("../escape.xlsx!Sheet1/A1:C5"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.SandboxDenied, envelope.Error!.Code);
    }

    [Fact]
    public void Missing_workbook_fails_with_file_not_found()
    {
        var envelope = _service.Edit(ChartAddOps("nope.xlsx!Sheet1/A1:C5"));

        Assert.False(envelope.IsOk);
        Assert.Equal(ErrorCodes.FileNotFound, envelope.Error!.Code);
    }

    // -------------------------------------------------------- expansion scope

    [Fact]
    public void Non_pptx_targets_pass_through_unexpanded()
    {
        var registry = new HandlerRegistry();
        registry.Register(new ExcelHandler(), ".xlsx");
        IReadOnlyList<EditOp> ops =
        [
            new EditOp
            {
                Op = "add",
                Path = "/Sheet1",
                Type = "chart",
                Props = new JsonObject { ["dataFrom"] = "metrics.xlsx!Sheet1/A1:C5" },
            },
        ];

        var result = CrossDocDataFrom.Expand(ops, DocumentKind.Xlsx, _ws.Workspace, registry);

        Assert.Same(ops, result); // untouched: dataFrom is a pptx-chart prop
    }

    [Fact]
    public void Pptx_ops_without_dataFrom_pass_through_unexpanded()
    {
        var registry = new HandlerRegistry();
        IReadOnlyList<EditOp> ops =
        [
            new EditOp { Op = "set", Path = "/slide[1]", Props = new JsonObject { ["background"] = "112233" } },
        ];

        var result = CrossDocDataFrom.Expand(ops, DocumentKind.Pptx, _ws.Workspace, registry);

        Assert.Same(ops, result);
    }

    // -------------------------------------------------------------- over MCP

    [Fact]
    public async Task DataFrom_works_identically_over_the_MCP_wire()
    {
        await using var srv = await McpTestServer.StartAsync(_service);

        var envelope = await srv.CallAsync("office_edit", new Dictionary<string, object?>
        {
            ["file"] = "deck.pptx",
            ["ops"] = new[]
            {
                new Dictionary<string, object?>
                {
                    ["op"] = "add",
                    ["path"] = "/slide[1]",
                    ["type"] = "chart",
                    ["props"] = new Dictionary<string, string>
                    {
                        ["kind"] = "bar",
                        ["title"] = "MCP chart",
                        ["dataFrom"] = "metrics.xlsx!Sheet1/A1:C5",
                    },
                },
            },
        });

        EnvelopeAssert.Ok(envelope);

        var chart = EnvelopeAssert.Ok(await srv.CallAsync("office_get", new Dictionary<string, object?>
        {
            ["file"] = "deck.pptx",
            ["path"] = "/slide[1]/chart[1]",
        }));
        Assert.Equal("MCP chart", chart.GetProperty("title").GetString());
        Assert.Equal("North", chart.GetProperty("series")[0].GetProperty("name").GetString());
        Assert.Equal(4, chart.GetProperty("categories").GetArrayLength());
    }
}
