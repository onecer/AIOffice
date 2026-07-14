using System.Text.Json.Nodes;
using AIOffice.Cli;
using AIOffice.Core;
using AIOffice.Core.Cli;
using Xunit;

namespace AIOffice.Mcp.Tests;

/// <summary>
/// End-to-end CLI plumbing tests: they drive the REAL <see cref="FileVerbs"/>
/// (the same object <c>Program</c> dispatches to) via <see cref="ArgParser"/> and
/// read the returned envelope back, so a flag that is DECLARED yet never
/// forwarded into the handler ctx — the v1.23.1 <c>read --delimiter</c> silent
/// drop, and the create/get gaps this suite lands with — fails here even though
/// the handler-level tests (which inject the ctx directly) stay green.
///
/// Each test proves a CLI flag actually reaches the handler AND changes behavior
/// (a range that windows, a sheet that selects, a dry-run that doesn't persist,
/// a track that marks a revision, a wrong expect-rev that refuses to write).
/// </summary>
public sealed class CliIntegrationTests : IDisposable
{
    private readonly string _root;
    private readonly FileVerbs _verbs;

    public CliIntegrationTests()
    {
        _root = Path.Combine(Path.GetTempPath(), "aioffice-cli-integration", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_root);
        var snapshots = new SnapshotStore();
        // Fully qualified: AIOffice.Mcp also declares a HandlerDiscovery, and this
        // test lives under the AIOffice.Mcp.Tests namespace.
        _verbs = new FileVerbs(new Workspace(_root), AIOffice.Cli.HandlerDiscovery.Discover(snapshots), snapshots);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_root, recursive: true);
        }
        catch (IOException)
        {
            // Best-effort cleanup.
        }
    }

    // ----- helpers ---------------------------------------------------------

    private Envelope Run(Func<ParsedArgs, Envelope> verb, params string[] argv) =>
        verb(ArgParser.Parse(argv));

    private static JsonObject Ok(Envelope env)
    {
        Assert.True(env.IsOk, env.ToJson());
        return JsonNode.Parse(env.ToJson())!["data"]!.AsObject();
    }

    private static string[] WarningCodes(Envelope env)
    {
        var meta = JsonNode.Parse(env.ToJson())!["meta"];
        var warnings = meta?["warnings"]?.AsArray();
        return warnings is null ? [] : [.. warnings.Select(w => w!["code"]!.GetValue<string>())];
    }

    /// <summary>
    /// The failure code, whichever layer rejected: pre-handler validation in
    /// <see cref="FileVerbs"/> throws <see cref="AiofficeException"/> (Program /
    /// the MCP service catch it and shape the envelope), while handler-level
    /// validation returns an error envelope. Tests should not care which.
    /// </summary>
    private string RejectCode(Func<ParsedArgs, Envelope> verb, params string[] argv)
    {
        try
        {
            var env = verb(ArgParser.Parse(argv));
            Assert.False(env.IsOk, "expected a rejection, got: " + env.ToJson());
            return env.Error!.Code;
        }
        catch (AiofficeException ex)
        {
            return ex.Code;
        }
    }

    private void SeedDocxParagraphs(string file)
    {
        Assert.True(_verbs.Create(ArgParser.Parse(["create", file])).IsOk);
        Assert.True(_verbs.Edit(ArgParser.Parse([
            "edit", file, "--ops",
            """[{"op":"add","path":"/body","type":"p","props":{"text":"Para One"}},{"op":"add","path":"/body","type":"p","props":{"text":"Para Two"}},{"op":"add","path":"/body","type":"p","props":{"text":"Para Three"}}]""",
        ])).IsOk);
    }

    private void SeedXlsxColumn(string file)
    {
        Assert.True(_verbs.Create(ArgParser.Parse(["create", file])).IsOk);
        Assert.True(_verbs.Edit(ArgParser.Parse([
            "edit", file, "--ops",
            """[{"op":"set","path":"/Sheet1/A1","props":{"values":[["1"],["2"],["3"],["4"],["5"]]}}]""",
        ])).IsOk);
    }

    // ===== read ============================================================

    [Fact]
    public void Read_text_range_windows_the_paragraphs()
    {
        SeedDocxParagraphs("doc.docx");

        var full = Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "text"));
        Assert.Equal("1..4", full["range"]!.GetValue<string>());
        Assert.Equal(4, full["lines"]!.AsArray().Count);

        // --range must actually reach the handler and clip the emitted window.
        var windowed = Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "text", "--range", "2..3"));
        Assert.Equal("2..3", windowed["range"]!.GetValue<string>());
        var lines = windowed["lines"]!.AsArray();
        Assert.Equal(2, lines.Count);
        Assert.Equal("/body/p[2]", lines[0]!["path"]!.GetValue<string>());
        Assert.Equal("/body/p[3]", lines[1]!["path"]!.GetValue<string>());
    }

    [Fact]
    public void Read_maxBytes_truncates_and_warns()
    {
        SeedDocxParagraphs("doc.docx");

        var env = Run(_verbs.Read, "read", "doc.docx", "--view", "text", "--max-bytes", "20");
        Assert.True(env.IsOk, env.ToJson());
        // The CLI 'max-bytes' -> ctx 'maxBytes' rename must survive: if it dropped,
        // no truncation would fire and the warning would be absent.
        Assert.Contains("result_truncated", WarningCodes(env));
    }

    [Fact]
    public void Read_csv_sheet_selects_the_named_sheet()
    {
        Assert.True(_verbs.Create(ArgParser.Parse(["create", "m.xlsx"])).IsOk);
        Assert.True(_verbs.Edit(ArgParser.Parse([
            "edit", "m.xlsx", "--ops",
            """[{"op":"set","path":"/Sheet1/A1","props":{"values":[["one"]]}},{"op":"add","path":"/Second","type":"sheet"},{"op":"set","path":"/Second/A1","props":{"values":[["two"]]}}]""",
        ])).IsOk);

        // Default = first sheet.
        Assert.Equal("one\r\n", Ok(Run(_verbs.Read, "read", "m.xlsx", "--view", "csv"))["content"]!.GetValue<string>());
        // --sheet must reach the handler and switch which grid is emitted.
        Assert.Equal("two\r\n",
            Ok(Run(_verbs.Read, "read", "m.xlsx", "--view", "csv", "--sheet", "Second"))["content"]!.GetValue<string>());
    }

    [Fact]
    public void Read_structure_depth_bounds_the_tree()
    {
        SeedDocxParagraphs("doc.docx");
        // A paragraph with text carries a run child, so depth has something to prune.
        static JsonNode P2(JsonObject data) =>
            data["root"]!["children"]!.AsArray().Single(n => n!["path"]!.GetValue<string>() == "/body/p[2]")!;

        // Default depth (3): the run child comes back in a populated children array.
        var deepP2 = P2(Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "structure")));
        Assert.True(deepP2["childCount"]!.GetValue<int>() >= 1);
        Assert.NotNull(deepP2["children"]);
        Assert.NotEmpty(deepP2["children"]!.AsArray());

        // --depth 1 must reach the handler and prune the nested children out (the
        // childCount is still reported, but the children array is gone).
        var shallowP2 = P2(Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "structure", "--depth", "1")));
        Assert.True(shallowP2["childCount"]!.GetValue<int>() >= 1);
        Assert.Null(shallowP2["children"]);
    }

    [Fact]
    public void Read_stream_forces_the_streaming_path_on_a_small_workbook()
    {
        SeedXlsxColumn("grid.xlsx");

        // Without --stream a small workbook never touches the SAX path.
        Assert.DoesNotContain("stream_fallback", WarningCodes(Run(_verbs.Read, "read", "grid.xlsx", "--view", "outline")));
        // --stream must reach the handler: outline can't stream, so the explicit
        // request surfaces as a stream_fallback warning even on a tiny file.
        Assert.Contains("stream_fallback", WarningCodes(Run(_verbs.Read, "read", "grid.xlsx", "--view", "outline", "--stream")));
    }

    // ===== create ==========================================================

    [Fact]
    public void Create_title_names_the_first_sheet()
    {
        var data = Ok(Run(_verbs.Create, "create", "titled.xlsx", "--title", "MySheet"));
        var sheets = data["sheets"]!.AsArray();
        Assert.Equal("MySheet", sheets[0]!.GetValue<string>());
    }

    [Fact]
    public void Create_kind_override_routes_to_the_named_handler()
    {
        // The .xlsx extension would resolve to the xlsx handler on its own; --kind
        // docx must reach ResolveHandler and win, so a docx document is produced.
        var data = Ok(Run(_verbs.Create, "create", "override.xlsx", "--kind", "docx"));
        Assert.Equal("docx", data["kind"]!.GetValue<string>());
    }

    [Fact]
    public void Create_from_delimiter_overrides_the_sniffer()
    {
        // A semicolon-separated file: the sniffer picks ';' on its own.
        File.WriteAllText(Path.Combine(_root, "src.csv"), "name;note\nACME;value\n");

        // Sniffed: two columns split on the semicolon.
        var sniffed = _verbs.Create(ArgParser.Parse(["create", "sniff.xlsx", "--from", "src.csv"]));
        Assert.True(sniffed.IsOk, sniffed.ToJson());
        var sniffRow = Ok(Run(_verbs.Get, "get", "sniff.xlsx", "/Sheet1/A1:B1"))["values"]!.AsArray()[0]!.AsArray();
        Assert.Equal("name", sniffRow[0]!.GetValue<string>());
        Assert.Equal("note", sniffRow[1]!.GetValue<string>());

        // --delimiter comma must reach the importer and force comma-splitting, so
        // the whole "name;note" line lands in a single cell (behavior CHANGES).
        var forced = _verbs.Create(ArgParser.Parse(["create", "comma.xlsx", "--from", "src.csv", "--delimiter", "comma"]));
        Assert.True(forced.IsOk, forced.ToJson());
        var forcedA1 = Ok(Run(_verbs.Get, "get", "comma.xlsx", "/Sheet1/A1"))["value"]!.GetValue<string>();
        Assert.Equal("name;note", forcedA1);
    }

    // ===== get / query =====================================================

    [Fact]
    public void Get_path_returns_the_cell_value()
    {
        SeedXlsxColumn("grid.xlsx");
        var cell = Ok(Run(_verbs.Get, "get", "grid.xlsx", "/Sheet1/A1"));
        Assert.Equal("cell", cell["kind"]!.GetValue<string>());
        Assert.Equal(1, cell["value"]!.GetValue<int>());
    }

    [Fact]
    public void Get_maxCells_caps_the_range_and_warns()
    {
        SeedXlsxColumn("grid.xlsx");

        // Default cap (1000) emits all five rows.
        var full = Ok(Run(_verbs.Get, "get", "grid.xlsx", "/Sheet1/A1:A5"));
        Assert.Equal(5, full["values"]!.AsArray().Count);
        Assert.False(full["truncated"]!.GetValue<bool>());

        // --max-cells 2 (CLI 'max-cells' -> ctx 'maxCells') must reach the handler
        // and clip the emitted rows + raise the result_truncated warning it cites.
        var capped = Run(_verbs.Get, "get", "grid.xlsx", "/Sheet1/A1:A5", "--max-cells", "2");
        var cappedData = Ok(capped);
        Assert.Equal(2, cappedData["values"]!.AsArray().Count);
        Assert.True(cappedData["truncated"]!.GetValue<bool>());
        Assert.Contains("result_truncated", WarningCodes(capped));
    }

    [Fact]
    public void Get_stream_forces_the_streaming_path_on_a_small_workbook()
    {
        SeedXlsxColumn("grid.xlsx");

        // Without --stream a small workbook is served by the DOM path, whose cell
        // shape carries no 'streamed' marker.
        var dom = Ok(Run(_verbs.Get, "get", "grid.xlsx", "/Sheet1/A1"));
        Assert.Null(dom["streamed"]);

        // --stream must reach the handler and route the get through the SAX path,
        // whose cell shape sets streamed=true (mutation-drop of the forward fails here).
        var streamed = Ok(Run(_verbs.Get, "get", "grid.xlsx", "/Sheet1/A1", "--stream"));
        Assert.True(streamed["streamed"]!.GetValue<bool>());
    }

    [Fact]
    public void Query_selector_returns_matching_paths()
    {
        SeedDocxParagraphs("doc.docx");
        var data = Ok(Run(_verbs.Query, "query", "doc.docx", "p"));
        Assert.Equal(4, data["count"]!.GetValue<int>());
        Assert.Equal(4, data["matches"]!.AsArray().Count);
    }

    // ===== edit ============================================================

    [Fact]
    public void Edit_set_sugar_applies_the_property()
    {
        SeedDocxParagraphs("doc.docx");
        var data = Ok(Run(_verbs.Edit, "edit", "doc.docx", "--set", "/body/p[2]", "text=Updated"));
        Assert.Equal(1, data["applied"]!.GetValue<int>());

        var text = Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "text"))["lines"]!.AsArray()[1]!["text"]!.GetValue<string>();
        Assert.Equal("Updated", text);
    }

    [Fact]
    public void Edit_add_then_remove_sugar_change_the_body()
    {
        SeedDocxParagraphs("doc.docx");

        Assert.True(Run(_verbs.Edit, "edit", "doc.docx", "--add", "/body", "--type", "p", "text=Fourth").IsOk);
        Assert.Equal(5, Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "text"))["totalParagraphs"]!.GetValue<int>());

        Assert.True(Run(_verbs.Edit, "edit", "doc.docx", "--remove", "/body/p[5]").IsOk);
        Assert.Equal(4, Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "text"))["totalParagraphs"]!.GetValue<int>());
    }

    [Fact]
    public void Edit_dryRun_reports_but_does_not_persist()
    {
        SeedDocxParagraphs("doc.docx");

        var data = Ok(Run(_verbs.Edit, "edit", "doc.docx", "--set", "/body/p[2]", "text=SHOULD-NOT-STICK", "--dry-run"));
        Assert.True(data["dryRun"]!.GetValue<bool>());

        // The file on disk is unchanged: p[2] is still the seeded text.
        var text = Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "text"))["lines"]!.AsArray()[1]!["text"]!.GetValue<string>();
        Assert.Equal("Para One", text);
    }

    [Fact]
    public void Edit_track_and_author_record_a_revision()
    {
        SeedDocxParagraphs("doc.docx");

        Assert.True(Run(_verbs.Edit, "edit", "doc.docx", "--set", "/body/p[2]", "text=Tracked", "--track", "--author", "Alice").IsOk);

        var revisions = Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "revisions"));
        Assert.True(revisions["count"]!.GetValue<int>() >= 1);
        // --author must reach the handler and stamp the revision.
        Assert.All(revisions["revisions"]!.AsArray(), r => Assert.Equal("Alice", r!["author"]!.GetValue<string>()));
    }

    [Fact]
    public void Edit_expectRev_wrong_fails_before_any_write()
    {
        SeedDocxParagraphs("doc.docx");
        var before = Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "text"))["lines"]!.AsArray()[1]!["text"]!.GetValue<string>();

        Assert.Equal(ErrorCodes.StaleAddress,
            RejectCode(_verbs.Edit, "edit", "doc.docx", "--set", "/body/p[2]", "text=Nope", "--expect-rev", "deadbeef"));

        // Nothing was written: the stale-rev guard fires before the edit.
        var after = Ok(Run(_verbs.Read, "read", "doc.docx", "--view", "text"))["lines"]!.AsArray()[1]!["text"]!.GetValue<string>();
        Assert.Equal(before, after);
    }

    [Fact]
    public void Edit_find_replace_reports_aggregate_replacements()
    {
        SeedDocxParagraphs("doc.docx");
        var data = Ok(Run(_verbs.Edit, "edit", "doc.docx", "--find", "Para", "--replace", "Section"));
        // All three seeded paragraphs contain "Para" (the empty p[1] does not).
        Assert.Equal(3, data["replacements"]!.GetValue<int>());
        Assert.Equal(3, data["locations"]!.AsArray().Count);
    }

    [Fact]
    public void Edit_find_combined_with_ops_is_invalid_args()
    {
        SeedDocxParagraphs("doc.docx");
        Assert.Equal(ErrorCodes.InvalidArgs,
            RejectCode(_verbs.Edit, "edit", "doc.docx", "--find", "Para", "--ops", "[]"));
    }

    [Fact]
    public void Edit_set_and_add_together_is_invalid_args()
    {
        SeedDocxParagraphs("doc.docx");
        Assert.Equal(ErrorCodes.InvalidArgs,
            RejectCode(_verbs.Edit, "edit", "doc.docx", "--set", "/body/p[1]", "--add", "/body", "--type", "p"));
    }

    // ===== render ==========================================================

    [Fact]
    public void Render_to_text_scope_renders_only_the_scoped_node()
    {
        SeedDocxParagraphs("doc.docx");
        var data = Ok(Run(_verbs.Render, "render", "doc.docx", "--to", "text", "--scope", "/body/p[2]"));
        // --to and --scope both reach the handler: only p[2]'s text comes back.
        Assert.Equal("/body/p[2]", data["scope"]!.GetValue<string>());
        Assert.Equal("Para One", data["content"]!.GetValue<string>());
    }

    [Fact]
    public void Render_o_writes_the_output_file()
    {
        SeedDocxParagraphs("doc.docx");
        var data = Ok(Run(_verbs.Render, "render", "doc.docx", "--to", "html", "-o", "out.html"));
        var written = data["written"]!.GetValue<string>();
        Assert.True(File.Exists(written), $"expected {written} on disk");
    }

    [Fact]
    public void Render_o_writes_the_output_file_for_xlsx()
    {
        // Regression guard: xlsx render used to ignore -o (returned inline content,
        // wrote nothing) while docx/pptx honored it. It must now write the file and
        // report the resolved path under 'written', exactly like docx.
        SeedXlsxColumn("grid.xlsx");
        var data = Ok(Run(_verbs.Render, "render", "grid.xlsx", "--to", "html", "-o", "grid.html"));
        var written = data["written"]!.GetValue<string>();
        Assert.True(File.Exists(written), $"expected {written} on disk");
        // The written file carries the same html the inline envelope returned.
        Assert.Equal(data["content"]!.GetValue<string>(), File.ReadAllText(written));
    }

    [Fact]
    public void Render_without_o_stays_inline_for_xlsx()
    {
        // Byte-safety: with no -o the xlsx envelope is unchanged — inline content,
        // and no 'written' field is emitted.
        SeedXlsxColumn("grid.xlsx");
        var data = Ok(Run(_verbs.Render, "render", "grid.xlsx", "--to", "html"));
        Assert.NotNull(data["content"]);
        Assert.Null(data["written"]);
    }

    [Fact]
    public void Render_o_writes_the_output_file_for_pptx()
    {
        // pptx reports the written path under 'output' (its existing WithOptionalOutput
        // shape), not 'written'; the parity that matters is that -o actually writes a
        // file for all three formats.
        Assert.True(_verbs.Create(ArgParser.Parse(["create", "deck.pptx"])).IsOk);
        var data = Ok(Run(_verbs.Render, "render", "deck.pptx", "--to", "html", "-o", "deck.html"));
        var written = data["output"]!.GetValue<string>();
        Assert.True(File.Exists(written), $"expected {written} on disk");
    }

    [Fact]
    public void Render_bad_engine_is_invalid_args()
    {
        SeedDocxParagraphs("doc.docx");
        Assert.Equal(ErrorCodes.InvalidArgs,
            RejectCode(_verbs.Render, "render", "doc.docx", "--to", "png", "--engine", "bogus"));
    }

    [Fact]
    public void Render_svg_with_soffice_engine_surfaces_the_engine_fallback_warning()
    {
        // The soffice engine renders png+pdf only, so a native svg target falls
        // back to the native engine with an engine_fallback warning — end-to-end
        // through the real render dispatch. This needs NO LibreOffice on the box:
        // Parse("soffice") and the native-target path never probe soffice, so the
        // warning fires deterministically on every runner (incl. CI).
        Assert.True(_verbs.Create(ArgParser.Parse(["create", "deck.pptx"])).IsOk);
        var env = Run(_verbs.Render, "render", "deck.pptx", "--to", "svg", "--engine", "soffice");
        Assert.True(env.IsOk, env.ToJson());
        Assert.Contains("engine_fallback", WarningCodes(env));
    }

    // ===== diff ============================================================

    [Fact]
    public void Diff_snapshot_lists_the_changes()
    {
        // Create + edit; the pre-edit snapshot 1 is the empty document.
        SeedDocxParagraphs("doc.docx");
        var data = Ok(Run(_verbs.Diff, "diff", "doc.docx", "--snapshot", "1"));
        Assert.Equal("detailed", data["view"]!.GetValue<string>());
        Assert.True(data["changes"]!.AsArray().Count >= 1);
    }

    [Fact]
    public void Diff_view_summary_trims_each_change()
    {
        SeedDocxParagraphs("doc.docx");

        var detailed = Ok(Run(_verbs.Diff, "diff", "doc.docx", "--snapshot", "1"))["changes"]!.AsArray();
        var summary = Ok(Run(_verbs.Diff, "diff", "doc.docx", "--snapshot", "1", "--view", "summary"))["changes"]!.AsArray();

        // detailed carries a per-change 'detail'; --view summary trims to kind+path.
        Assert.NotNull(detailed[0]!["detail"]);
        Assert.Null(summary[0]!["detail"]);
    }

    [Fact]
    public void Diff_without_a_baseline_is_invalid_args()
    {
        SeedDocxParagraphs("doc.docx");
        Assert.Equal(ErrorCodes.InvalidArgs, RejectCode(_verbs.Diff, "diff", "doc.docx"));
    }

    // ===== audit / validate ===============================================

    [Fact]
    public void Audit_category_and_severity_reach_the_auditor()
    {
        SeedDocxParagraphs("doc.docx");

        var accessibility = Ok(Run(_verbs.Audit, "audit", "doc.docx", "--category", "accessibility"));
        Assert.NotNull(accessibility["findings"]);
        Assert.NotNull(accessibility["summary"]);

        // --severity error hides info+warning: the reported infos count is zeroed.
        var errorsOnly = Ok(Run(_verbs.Audit, "audit", "doc.docx", "--severity", "error"));
        Assert.Equal(0, errorsOnly["summary"]!["infos"]!.GetValue<int>());
    }

    [Fact]
    public void Audit_fix_reports_fixed_and_remaining()
    {
        SeedDocxParagraphs("doc.docx");
        var data = Ok(Run(_verbs.Audit, "audit", "doc.docx", "--fix"));
        // --fix flips the auditor into fix mode: the payload carries fixed/remaining.
        Assert.NotNull(data["fixed"]);
        Assert.NotNull(data["remaining"]);
    }

    [Fact]
    public void Audit_unknown_category_is_invalid_args()
    {
        SeedDocxParagraphs("doc.docx");
        Assert.Equal(ErrorCodes.InvalidArgs,
            RejectCode(_verbs.Audit, "audit", "doc.docx", "--category", "bogus"));
    }

    [Fact]
    public void Validate_returns_the_issue_list()
    {
        SeedDocxParagraphs("doc.docx");
        var data = Ok(Run(_verbs.Validate, "validate", "doc.docx"));
        Assert.NotNull(data["valid"]);
        Assert.NotNull(data["issues"]);
    }
}
