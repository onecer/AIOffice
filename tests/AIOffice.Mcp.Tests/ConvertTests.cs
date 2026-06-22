using System.Text.Json;
using AIOffice.Core;
using AIOffice.Excel;
using AIOffice.Pptx;
using AIOffice.Word;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// The M9 <c>convert</c> verb end-to-end, driving the REAL docx/xlsx/pptx
/// handlers through the shared <see cref="CommandService"/> and the MCP server's
/// <c>office_convert</c> tool. Covers cross-format transfer via the neutral
/// model (docx↔pptx, xlsx→docx), the text bridges (docx↔md, xlsx→csv), the
/// render route (docx→pdf is skipped without a browser), the honest
/// <c>convert_lossy</c> warning, and the two error gates (same-extension /
/// unknown-target). Asserts stay platform-independent: structural content, not
/// byte sizes.
/// </summary>
public sealed class ConvertTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly CommandService _service;

    public ConvertTests()
    {
        var snapshots = new SnapshotStore(_ws.SnapshotDir);
        var registry = new HandlerRegistry();
        registry.Register(new WordHandler(snapshots), ".docx");
        registry.Register(new ExcelHandler(snapshots), ".xlsx");
        registry.Register(new PptxHandler(snapshots), ".pptx");
        _service = _ws.NewService(registry,
            new HashSet<DocumentKind> { DocumentKind.Docx, DocumentKind.Xlsx, DocumentKind.Pptx });
    }

    public void Dispose() => _ws.Dispose();

    // ─────────────────────────────────────────────────── helpers

    private JsonElement Convert(string src, string dest)
    {
        var args = new Dictionary<string, object?> { ["src"] = src, ["dest"] = dest };
        return ToElement(_service.Convert(System.Text.Json.Nodes.JsonNode.Parse(JsonSerializer.Serialize(args))!.AsObject()));
    }

    private static JsonElement ToElement(Envelope envelope)
    {
        using var doc = JsonDocument.Parse(envelope.ToJson());
        return doc.RootElement.Clone();
    }

    private string ReadView(string file, string view)
    {
        var args = new System.Text.Json.Nodes.JsonObject { ["file"] = file, ["view"] = view };
        var data = EnvelopeAssert.Ok(ToElement(_service.Read(args)));
        return data.GetRawText();
    }

    /// <summary>Builds a report.docx: title, 3 headings, bullets under each, and a table.</summary>
    private string BuildReport(string name = "report.docx")
    {
        var create = new System.Text.Json.Nodes.JsonObject { ["file"] = name, ["title"] = "Quarterly Report" };
        EnvelopeAssert.Ok(ToElement(_service.Create(create)));

        var ops = """
            [
              {"op":"add","path":"/body","type":"p","props":{"text":"Sales","style":"Heading1"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"North up 12%","list":"bullet"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"South flat","list":"bullet"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Costs","style":"Heading1"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Rent steady","list":"bullet"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Outlook","style":"Heading1"}},
              {"op":"add","path":"/body","type":"p","props":{"text":"Grow EU","list":"bullet"}},
              {"op":"add","path":"/body","type":"table","props":{"rows":2,"columns":2}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"Region"}},
              {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"Total"}},
              {"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"text":"North"}},
              {"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"text":"100"}}
            ]
            """;
        var edit = new System.Text.Json.Nodes.JsonObject
        {
            ["file"] = name,
            ["ops"] = System.Text.Json.Nodes.JsonNode.Parse(ops),
        };
        EnvelopeAssert.Ok(ToElement(_service.Edit(edit)));
        return name;
    }

    /// <summary>A 2-sheet workbook with a formula, for xlsx→docx. The first sheet stays Sheet1.</summary>
    private string BuildWorkbook(string name = "data.xlsx")
    {
        EnvelopeAssert.Ok(ToElement(_service.Create(new System.Text.Json.Nodes.JsonObject
        {
            ["file"] = name,
        })));

        var ops = """
            [
              {"op":"add","path":"/Totals","type":"sheet"},
              {"op":"set","path":"/Sheet1/A1:B3","props":{"values":[["Region","Units"],["North","10"],["South","20"]]}},
              {"op":"set","path":"/Totals/A1","props":{"value":"=SUM(Sheet1!B2:B3)"}}
            ]
            """;
        EnvelopeAssert.Ok(ToElement(_service.Edit(new System.Text.Json.Nodes.JsonObject
        {
            ["file"] = name,
            ["ops"] = System.Text.Json.Nodes.JsonNode.Parse(ops),
        })));
        return name;
    }

    // ─────────────────────────────────────────────────── cross-format

    [Fact]
    public void Docx_to_pptx_makes_a_slide_per_heading_with_bullets()
    {
        BuildReport();
        var data = EnvelopeAssert.Ok(Convert("report.docx", "deck.pptx"));

        Assert.Equal("docx", data.GetProperty("from").GetString());
        Assert.Equal("pptx", data.GetProperty("to").GetString());
        Assert.True(data.GetProperty("blocksWritten").GetInt32() > 0);

        // The deck's outline shows three slides (one per heading) carrying the bullets.
        var outline = ReadView("deck.pptx", "outline");
        Assert.Contains("Sales", outline, StringComparison.Ordinal);
        Assert.Contains("Costs", outline, StringComparison.Ordinal);
        Assert.Contains("Outlook", outline, StringComparison.Ordinal);

        var text = ReadView("deck.pptx", "text");
        Assert.Contains("North up 12%", text, StringComparison.Ordinal);

        // The deck validates clean.
        var valid = EnvelopeAssert.Ok(ToElement(_service.Validate(
            new System.Text.Json.Nodes.JsonObject { ["file"] = "deck.pptx" })));
        Assert.True(valid.GetProperty("valid").GetBoolean());
    }

    [Fact]
    public void Xlsx_to_docx_writes_a_table_per_sheet_with_cached_values()
    {
        BuildWorkbook();
        var data = EnvelopeAssert.Ok(Convert("data.xlsx", "data.docx"));
        Assert.Equal("xlsx", data.GetProperty("from").GetString());
        Assert.True(data.GetProperty("blocksWritten").GetInt32() > 0);

        // The docx carries a heading + table per sheet; the formula crossed as 30.
        var text = ReadView("data.docx", "text");
        Assert.Contains("Sheet1", text, StringComparison.Ordinal);
        Assert.Contains("Totals", text, StringComparison.Ordinal);
        Assert.Contains("30", text, StringComparison.Ordinal);

        var valid = EnvelopeAssert.Ok(ToElement(_service.Validate(
            new System.Text.Json.Nodes.JsonObject { ["file"] = "data.docx" })));
        Assert.True(valid.GetProperty("valid").GetBoolean());
    }

    [Fact]
    public void Pptx_to_docx_carries_titles_and_bullets()
    {
        BuildReport();
        EnvelopeAssert.Ok(Convert("report.docx", "deck.pptx"));

        var data = EnvelopeAssert.Ok(Convert("deck.pptx", "outline.docx"));
        Assert.Equal("pptx", data.GetProperty("from").GetString());

        var text = ReadView("outline.docx", "text");
        Assert.Contains("Sales", text, StringComparison.Ordinal);
        Assert.Contains("North up 12%", text, StringComparison.Ordinal);
    }

    // ─────────────────────────────────────────────────── markdown round-trip

    [Fact]
    public void Docx_to_md_and_back_preserves_headings_bullets_and_table()
    {
        BuildReport();
        var toMd = EnvelopeAssert.Ok(Convert("report.docx", "report.md"));
        Assert.Equal("md", toMd.GetProperty("to").GetString());

        var markdown = _ws.ReadFile("report.md");
        // The document Title rides along as a leading H1 (distinct from the first
        // body heading "Sales"), so the converted markdown is self-describing.
        Assert.StartsWith("# Quarterly Report", markdown, StringComparison.Ordinal);
        Assert.Contains("# Sales", markdown, StringComparison.Ordinal);
        Assert.Contains("- North up 12%", markdown, StringComparison.Ordinal);
        Assert.Contains("| Region |", markdown, StringComparison.Ordinal);

        // md → docx rebuilds an equivalent outline.
        EnvelopeAssert.Ok(Convert("report.md", "roundtrip.docx"));
        var text = ReadView("roundtrip.docx", "text");
        Assert.Contains("Sales", text, StringComparison.Ordinal);
        Assert.Contains("North up 12%", text, StringComparison.Ordinal);
        Assert.Contains("Region", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Xlsx_to_md_uses_the_neutral_markdown_serializer()
    {
        BuildWorkbook();
        EnvelopeAssert.Ok(Convert("data.xlsx", "data.md"));

        var markdown = _ws.ReadFile("data.md");
        Assert.Contains("## Sheet1", markdown, StringComparison.Ordinal); // sheet = level-2 heading
        Assert.Contains("| Region | Units |", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Xlsx_to_csv_exports_the_first_sheet()
    {
        BuildWorkbook();
        EnvelopeAssert.Ok(Convert("data.xlsx", "data.csv"));

        var csv = _ws.ReadFile("data.csv").Replace("\r\n", "\n", StringComparison.Ordinal);
        Assert.Contains("Region,Units", csv, StringComparison.Ordinal);
        Assert.Contains("North,10", csv, StringComparison.Ordinal);
    }

    // ─────────────────────────────────────────────────── lossiness

    [Fact]
    public void Convert_surfaces_a_convert_lossy_warning_naming_dropped_content()
    {
        // A deck with a chart loses the chart on the way to docx; convert must say so.
        EnvelopeAssert.Ok(ToElement(_service.Create(
            new System.Text.Json.Nodes.JsonObject { ["file"] = "charted.pptx", ["title"] = "Charted" })));
        var chartOp = """
            [{"op":"add","path":"/slide[1]","type":"chart","props":{
              "kind":"bar","categories":["Q1","Q2"],
              "series":[{"name":"Sales","values":[10,20]}]}}]
            """;
        EnvelopeAssert.Ok(ToElement(_service.Edit(new System.Text.Json.Nodes.JsonObject
        {
            ["file"] = "charted.pptx",
            ["ops"] = System.Text.Json.Nodes.JsonNode.Parse(chartOp),
        })));

        var envelope = ToElement(_service.Convert(
            new System.Text.Json.Nodes.JsonObject { ["src"] = "charted.pptx", ["dest"] = "charted.docx" }));
        EnvelopeAssert.Ok(envelope);

        var warnings = envelope.GetProperty("meta").GetProperty("warnings").EnumerateArray().ToList();
        Assert.Contains(warnings, w => w.GetProperty("code").GetString() == "convert_lossy" &&
            w.GetProperty("message").GetString()!.Contains("chart", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Xlsx_source_chart_drop_fires_convert_lossy_and_populates_dropped()
    {
        // A workbook with a chart loses the chart converting to docx/md. The
        // loss must reach BOTH data.dropped and a convert_lossy warning — not be
        // hidden behind a body-only "[dropped]" note (the M9/M10 regression this
        // test guards against; xlsx export now mirrors the pptx exporter).
        EnvelopeAssert.Ok(ToElement(_service.Create(
            new System.Text.Json.Nodes.JsonObject { ["file"] = "charted.xlsx" })));
        var ops = """
            [
              {"op":"set","path":"/Sheet1/A1:B3","props":{"values":[["Region","Units"],["North","10"],["South","20"]]}},
              {"op":"add","path":"/Sheet1","type":"chart","props":{"kind":"bar","dataRange":"A1:B3","anchor":"D2"}}
            ]
            """;
        EnvelopeAssert.Ok(ToElement(_service.Edit(new System.Text.Json.Nodes.JsonObject
        {
            ["file"] = "charted.xlsx",
            ["ops"] = System.Text.Json.Nodes.JsonNode.Parse(ops),
        })));

        var envelope = ToElement(_service.Convert(
            new System.Text.Json.Nodes.JsonObject { ["src"] = "charted.xlsx", ["dest"] = "charted.docx" }));
        var data = EnvelopeAssert.Ok(envelope);

        var dropped = data.GetProperty("dropped").EnumerateArray().Select(d => d.GetString()!).ToList();
        Assert.Contains(dropped, d => d.Contains("chart", StringComparison.OrdinalIgnoreCase));

        var warnings = envelope.GetProperty("meta").GetProperty("warnings").EnumerateArray().ToList();
        Assert.Contains(warnings, w => w.GetProperty("code").GetString() == "convert_lossy" &&
            w.GetProperty("message").GetString()!.Contains("chart", StringComparison.OrdinalIgnoreCase));
    }

    // ─────────────────────────────────────────────────── error gates

    [Fact]
    public void Same_extension_is_invalid_args()
    {
        BuildReport("a.docx");
        var error = EnvelopeAssert.Fail(
            ToElement(_service.Convert(
                new System.Text.Json.Nodes.JsonObject { ["src"] = "a.docx", ["dest"] = "b.docx" })),
            "invalid_args");
        Assert.Contains("edit", error.GetProperty("suggestion").GetString()!, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Unknown_destination_extension_is_unsupported_feature_naming_targets()
    {
        BuildReport("a.docx");
        var error = EnvelopeAssert.Fail(
            ToElement(_service.Convert(
                new System.Text.Json.Nodes.JsonObject { ["src"] = "a.docx", ["dest"] = "a.xyz" })),
            "unsupported_feature");
        Assert.Contains(".pptx", error.GetProperty("suggestion").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public void Overwriting_an_existing_destination_warns_convert_overwrite()
    {
        BuildReport();
        EnvelopeAssert.Ok(Convert("report.docx", "deck.pptx"));

        // Convert again over the same dest: convert_overwrite rides along.
        var envelope = ToElement(_service.Convert(
            new System.Text.Json.Nodes.JsonObject { ["src"] = "report.docx", ["dest"] = "deck.pptx" }));
        EnvelopeAssert.Ok(envelope);
        var warnings = envelope.GetProperty("meta").GetProperty("warnings").EnumerateArray().ToList();
        Assert.Contains(warnings, w => w.GetProperty("code").GetString() == "convert_overwrite");
    }

    // ─────────────────────────────────────────────────── over the wire

    [Fact]
    public async Task Office_convert_docx_to_pptx_over_the_wire()
    {
        BuildReport();
        await using var srv = await McpTestServer.StartAsync(_service);

        Assert.Equal(19, (await srv.Client.ListToolsAsync()).Count);

        var data = EnvelopeAssert.Ok(await srv.CallAsync("office_convert",
            new Dictionary<string, object?> { ["src"] = "report.docx", ["dest"] = "wire.pptx" }));
        Assert.Equal("pptx", data.GetProperty("to").GetString());
        Assert.True(data.GetProperty("blocksWritten").GetInt32() > 0);

        var outline = ReadView("wire.pptx", "outline");
        Assert.Contains("Sales", outline, StringComparison.Ordinal);
    }
}
