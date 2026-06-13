using System.Text.Json;
using System.Text.Json.Nodes;
using AIOffice.Core;
using AIOffice.Excel;
using AIOffice.Pptx;
using AIOffice.Word;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// The M5 markdown/csv bridge wiring shared by the CLI (<c>create --from</c>,
/// <c>read --view markdown|csv</c>) and MCP (<c>office_create.from</c>):
/// extension routing, the import matrix, sandboxing of the source path, and
/// the single-format bridge-view guard.
/// </summary>
public sealed class BridgeTests : IDisposable
{
    private readonly TempWorkspace _ws = new();
    private readonly CommandService _service;

    public BridgeTests()
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

    private static JsonElement ToElement(Envelope envelope)
    {
        using var doc = JsonDocument.Parse(envelope.ToJson());
        return doc.RootElement.Clone();
    }

    // ----- create --from routing -------------------------------------------

    [Fact]
    public void Markdown_source_imports_into_docx_and_round_trips()
    {
        _ws.WriteFile("notes.md",
            """
            # Quarterly Notes

            Some **bold** intro with a [link](https://example.com).

            ## Items

            - one
            - two

            | a | b |
            | - | - |
            | 1 | 2 |
            """);

        var created = ToElement(_service.Create(new JsonObject
        {
            ["file"] = "report.docx",
            ["from"] = "notes.md",
        }));
        EnvelopeAssert.Ok(created);
        Assert.True(File.Exists(Path.Combine(_ws.Root, "report.docx")));

        var outline = EnvelopeAssert.Ok(ToElement(_service.Read(new JsonObject
        {
            ["file"] = "report.docx",
            ["view"] = "outline",
        })));
        Assert.Contains("Quarterly Notes", outline.GetRawText(), StringComparison.Ordinal);

        var markdown = EnvelopeAssert.Ok(ToElement(_service.Read(new JsonObject
        {
            ["file"] = "report.docx",
            ["view"] = "markdown",
        }))).GetProperty("markdown").GetString()!;
        Assert.Contains("# Quarterly Notes", markdown, StringComparison.Ordinal);
        Assert.Contains("## Items", markdown, StringComparison.Ordinal);
        Assert.Contains("**bold**", markdown, StringComparison.Ordinal);
        Assert.Contains("[link](https://example.com)", markdown, StringComparison.Ordinal);
        Assert.Contains("- one", markdown, StringComparison.Ordinal);
        Assert.Contains("| a | b |", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void Csv_source_imports_into_xlsx_keeping_leading_zero_codes_text()
    {
        _ws.WriteFile("orders.csv",
            "code,customer,date,total\n" +
            "007,\"Acme, Inc.\",2026-06-12,1234.5\n");

        EnvelopeAssert.Ok(ToElement(_service.Create(new JsonObject
        {
            ["file"] = "orders.xlsx",
            ["from"] = "orders.csv",
        })));

        var cell = EnvelopeAssert.Ok(ToElement(_service.Get(new JsonObject
        {
            ["file"] = "orders.xlsx",
            ["path"] = "/Sheet1/A2",
        })));
        Assert.Equal("007", cell.GetProperty("value").GetString());
        Assert.Equal("text", cell.GetProperty("type").GetString());

        var csv = EnvelopeAssert.Ok(ToElement(_service.Read(new JsonObject
        {
            ["file"] = "orders.xlsx",
            ["view"] = "csv",
        }))).GetProperty("content").GetString()!;
        Assert.Contains("007", csv, StringComparison.Ordinal);
        Assert.Contains("\"Acme, Inc.\"", csv, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("notes.md", "data.xlsx")] // md needs a docx target
    [InlineData("notes.md", "deck.pptx")] // pptx has no importer at all
    [InlineData("orders.csv", "report.docx")] // csv needs an xlsx target
    public void Mismatched_pairs_fail_with_the_import_matrix(string source, string target)
    {
        _ws.WriteFile(source, "x,y\n1,2\n");
        var envelope = ToElement(_service.Create(new JsonObject
        {
            ["file"] = target,
            ["from"] = source,
        }));

        var error = EnvelopeAssert.Fail(envelope, ErrorCodes.InvalidArgs);
        Assert.Contains(".md/.markdown", error.GetProperty("suggestion").GetString(), StringComparison.Ordinal);
        Assert.Contains(".csv/.tsv", error.GetProperty("suggestion").GetString(), StringComparison.Ordinal);
        Assert.False(File.Exists(Path.Combine(_ws.Root, target)), "a failed import must not create the target");
    }

    [Fact]
    public void Unknown_source_extension_fails_with_the_import_matrix()
    {
        _ws.WriteFile("data.json", "{}");
        var error = EnvelopeAssert.Fail(ToElement(_service.Create(new JsonObject
        {
            ["file"] = "report.docx",
            ["from"] = "data.json",
        })), ErrorCodes.InvalidArgs);
        Assert.Contains(".csv/.tsv", error.GetProperty("suggestion").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Missing_source_is_file_not_found_and_escaping_source_is_sandbox_denied()
    {
        EnvelopeAssert.Fail(ToElement(_service.Create(new JsonObject
        {
            ["file"] = "report.docx",
            ["from"] = "nope.md",
        })), ErrorCodes.FileNotFound);

        EnvelopeAssert.Fail(ToElement(_service.Create(new JsonObject
        {
            ["file"] = "report.docx",
            ["from"] = "../outside.md",
        })), ErrorCodes.SandboxDenied);
    }

    [Fact]
    public void Default_handler_hook_refuses_with_unsupported_feature()
    {
        // The IFormatHandler default: formats without an importer (pptx) refuse.
        IFormatHandler handler = new PptxHandler();
        var ex = Assert.Throws<AiofficeException>(() => handler.CreateFrom(
            new CommandContext { Workspace = _ws.Workspace, File = Path.Combine(_ws.Root, "deck.pptx") },
            Path.Combine(_ws.Root, "notes.md")));
        Assert.Equal(ErrorCodes.UnsupportedFeature, ex.Code);
        Assert.False(string.IsNullOrWhiteSpace(ex.Suggestion));
    }

    // ----- bridge view guard -------------------------------------------------

    [Theory]
    [InlineData("orders.xlsx", "markdown", "csv")] // markdown is docx-only; xlsx views offered instead
    [InlineData("report.docx", "csv", "markdown")] // csv is xlsx-only; docx views offered instead
    [InlineData("deck.pptx", "csv", "structure")] // pptx has neither bridge view
    public void Bridge_views_on_the_wrong_format_are_unsupported_with_valid_views(
        string file, string view, string expectedOffer)
    {
        EnvelopeAssert.Ok(ToElement(_service.Create(new JsonObject { ["file"] = file })));
        var error = EnvelopeAssert.Fail(ToElement(_service.Read(new JsonObject
        {
            ["file"] = file,
            ["view"] = view,
        })), ErrorCodes.UnsupportedFeature);
        Assert.Contains(expectedOffer, error.GetProperty("suggestion").GetString(), StringComparison.Ordinal);
    }

    // ----- over the wire ------------------------------------------------------

    [Fact]
    public async Task Office_create_with_from_and_a_dataValidation_op_work_over_the_wire()
    {
        _ws.WriteFile("orders.csv", "sku,qty\n007,3\n012,5\n");

        await using var srv = await McpTestServer.StartAsync(_service);

        var tools = await srv.Client.ListToolsAsync();
        Assert.Equal(16, tools.Count); // the bridge adds args, never tools
        var createSchema = tools.First(t => t.Name == "office_create").JsonSchema.GetRawText();
        Assert.Contains("\"from\"", createSchema, StringComparison.Ordinal);

        var created = await srv.CallAsync("office_create", new Dictionary<string, object?>
        {
            ["file"] = "orders.xlsx",
            ["from"] = "orders.csv",
        });
        EnvelopeAssert.Ok(created);

        var edited = await srv.CallAsync("office_edit", new Dictionary<string, object?>
        {
            ["file"] = "orders.xlsx",
            ["ops"] = new object[]
            {
                new Dictionary<string, object?>
                {
                    ["op"] = "add",
                    ["path"] = "/Sheet1/B2:B10",
                    ["type"] = "dataValidation",
                    ["props"] = new Dictionary<string, object?>
                    {
                        ["kind"] = "wholeNumber",
                        ["operator"] = "between",
                        ["value"] = 1,
                        ["value2"] = 100,
                    },
                },
            },
        });
        EnvelopeAssert.Ok(edited);

        var read = await srv.CallAsync("office_read", new Dictionary<string, object?>
        {
            ["file"] = "orders.xlsx",
            ["view"] = "csv",
        });
        var content = EnvelopeAssert.Ok(read).GetProperty("content").GetString()!;
        Assert.Contains("007", content, StringComparison.Ordinal);
    }

    // M6: add a docx equation over MCP, then read its stored LaTeX back. Exercises
    // the new "equation" op type and the /body/p[i]/omath[j] addressing form
    // through the real server + EditOp.ParseBatch gate (tool count is 16 since M8).
    [Fact]
    public async Task Office_edit_adds_a_docx_equation_over_the_wire()
    {
        await using var srv = await McpTestServer.StartAsync(_service);

        Assert.Equal(16, (await srv.Client.ListToolsAsync()).Count);

        EnvelopeAssert.Ok(await srv.CallAsync("office_create", new Dictionary<string, object?>
        {
            ["file"] = "eq.docx",
        }));

        var edited = await srv.CallAsync("office_edit", new Dictionary<string, object?>
        {
            ["file"] = "eq.docx",
            ["ops"] = new object[]
            {
                new Dictionary<string, object?>
                {
                    ["op"] = "add",
                    ["path"] = "/body/p[1]",
                    ["type"] = "equation",
                    ["props"] = new Dictionary<string, object?> { ["latex"] = "E = mc^2" },
                },
            },
        });
        var op = EnvelopeAssert.Ok(edited).GetProperty("ops")[0];
        Assert.Equal("/body/p[1]/omath[1]", op.GetProperty("path").GetString());

        var got = await srv.CallAsync("office_get", new Dictionary<string, object?>
        {
            ["file"] = "eq.docx",
            ["path"] = "/body/p[1]/omath[1]",
        });
        Assert.Equal("E = mc^2",
            EnvelopeAssert.Ok(got).GetProperty("properties").GetProperty("latex").GetString());
    }
}
