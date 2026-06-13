using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// The machine-readable description of the whole aioffice command surface —
/// the single source behind both <c>office_schema</c> and <c>aioffice schema</c>.
/// 16 verbs, mirroring the 16 MCP tools (the <c>preview</c> verb carries
/// both <c>preview_open</c> and <c>preview_selection</c>) plus the <c>mcp</c>
/// verb itself. M8 added <c>diff</c> / <c>office_diff</c>.
/// </summary>
public static class SurfaceSchema
{
    /// <summary>One verb in the surface (serialized camelCase by the envelope writer).</summary>
    public sealed record VerbInfo(
        string Name,
        string Summary,
        IReadOnlyDictionary<string, string> Args,
        IReadOnlyList<string> Errors,
        IReadOnlyList<Example> Examples,
        string? McpTool);

    /// <summary>A copy-pasteable CLI invocation with a one-line note.</summary>
    public sealed record Example(string Call, string Note);

    private static VerbInfo Verb(
        string name, string summary, string? mcpTool,
        (string Arg, string Doc)[] args, string[] errors, (string Call, string Note)[] examples) =>
        new(
            name,
            summary,
            args.ToDictionary(a => a.Arg, a => a.Doc, StringComparer.Ordinal),
            errors,
            [.. examples.Select(e => new Example(e.Call, e.Note))],
            mcpTool);

    private static readonly IReadOnlyList<VerbInfo> Verbs =
    [
        Verb("create", "Create a blank document (kind inferred from the extension), or import via --from (md→docx, csv→xlsx).", "office_create",
            [("file", "target path ending in .docx/.xlsx/.pptx"),
             ("--from", "import source: .md/.markdown → .docx, .csv/.tsv → .xlsx; mismatched pairs → invalid_args with the matrix"),
             ("--kind", "docx|xlsx|pptx, overrides extension inference"),
             ("--title", "document title written to core properties"),
             ("--overwrite", "replace an existing file")],
            [ErrorCodes.InvalidArgs, ErrorCodes.SandboxDenied, ErrorCodes.UnsupportedFeature, ErrorCodes.InternalError],
            [("aioffice create out/q4.pptx --title 'Q4 Review'", "kind pptx inferred from extension"),
             ("aioffice create report.docx --from notes.md", "markdown bridge: headings/lists/tables/links/code become real docx")]),

        Verb("read", "Read a document as outline, plain text, stats, structure, properties, markdown (docx), csv (xlsx) and more.", "office_read",
            [("file", "document path"),
             ("--view", "outline|text|stats|structure|properties (default outline; properties = core+custom doc props, all formats); docx adds revisions|comments|styles|fields|markdown (fields = content controls); xlsx adds csv|styles (named cell styles)"),
             ("--range", "scope 'a..b' (1-based): paragraphs/slides/rows; csv view also takes A1:C10"),
             ("--sheet", "csv view: sheet name (default first)"),
             ("--max-bytes", "cap payload; truncation flagged in warnings")],
            [ErrorCodes.FileNotFound, ErrorCodes.SandboxDenied, ErrorCodes.FormatCorrupt, ErrorCodes.InvalidArgs, ErrorCodes.UnsupportedFeature],
            [("aioffice read report.docx --view outline", "headings/tables skeleton with canonical paths"),
             ("aioffice read report.docx --view markdown", "the whole body back as GFM markdown")]),

        Verb("query", "Find elements with a CSS-like selector; returns canonical paths.", "office_query",
            [("file", "document path"),
             ("selector", "e.g. p[style=Heading1], cell[value>100], shape:contains('Q3')")],
            [ErrorCodes.InvalidArgs, ErrorCodes.FileNotFound, ErrorCodes.SandboxDenied, ErrorCodes.FormatCorrupt],
            [("aioffice query deck.pptx \"shape:contains('Q3')\"", "paths feed office_get/office_edit")]),

        Verb("get", "Read one node and its properties by canonical path.", "office_get",
            [("file", "document path"),
             ("path", "canonical 1-based path, e.g. /body/p[3], /Sheet1/B2; '/' for document props")],
            [ErrorCodes.InvalidPath, ErrorCodes.FileNotFound, ErrorCodes.SandboxDenied, ErrorCodes.FormatCorrupt],
            [("aioffice get data.xlsx /Sheet1/B2", "value, valueType, formula, numberFormat")]),

        Verb("edit", "Apply an atomic batch of set/add/remove/move/replace ops: all-or-nothing single save, auto-snapshot, optional rev guard.", "office_edit",
            [("file", "document path"),
             ("--ops", "JSON array of ops, or @ops.json"),
             ("--set/--add/--remove", "single-op sugar: --set <path> k=v..., --add <path> --type T k=v..., --remove <path>"),
             ("--find/--replace", "document-wide find/replace sugar (docx body+headers+footers, every sheet, every slide+notes); modifiers: --regex, --match-case, --whole-word; returns aggregate {replacements, locations}"),
             ("--expect-rev", "12-hex rev from a previous meta.rev; mismatch -> stale_address before any write"),
             ("--dry-run", "validate the whole batch without writing")],
            [ErrorCodes.StaleAddress, ErrorCodes.InvalidPath, ErrorCodes.InvalidArgs, ErrorCodes.UnsupportedFeature,
             ErrorCodes.FileNotFound, ErrorCodes.SandboxDenied, ErrorCodes.FormatCorrupt],
            [("aioffice edit r.docx --set /body/p[1] style=Heading1", "single-op sugar"),
             ("aioffice edit r.docx --find 2025 --replace 2026 --track", "document-wide tracked find/replace"),
             ("aioffice edit r.docx --ops @ops.json --expect-rev a3f9c12be01d", "atomic batch with optimistic lock")]),

        Verb("render", "Render the document (or a subtree) to html/svg/text/png for inspection.", "office_render",
            [("file", "document path"),
             ("--to", "html|svg|text|png (default html; png screenshots via a local Chromium and is written to a file)"),
             ("--scope", "render only this subtree, e.g. /slide[3], /Sheet1/A1:F20; png on a multi-slide pptx defaults to /slide[1] with a warning"),
             ("-o", "output file or directory inside the workspace (png default: source path with .png)")],
            [ErrorCodes.UnsupportedFeature, ErrorCodes.InvalidPath, ErrorCodes.FileNotFound,
             ErrorCodes.SandboxDenied, ErrorCodes.FormatCorrupt, ErrorCodes.InternalError],
            [("aioffice render deck.pptx --to svg --scope /slide[3]", "one slide as svg, inlined when small"),
             ("aioffice render report.docx --to png -o look.png", "browser screenshot; needs Chrome/Edge/Chromium installed")]),

        Verb("validate", "OpenXmlValidator schema check plus lint; issues carry suggestions.", "office_validate",
            [("file", "document path")],
            [ErrorCodes.FileNotFound, ErrorCodes.SandboxDenied, ErrorCodes.FormatCorrupt],
            [("aioffice validate data.xlsx", "ok:true even when the doc has issues — read data.valid")]),

        Verb("audit", "Accessibility + quality lint; findings are data (ok:true, exit 0). --fix applies safe autofixes, then re-audits.", "office_audit",
            [("file", "document path (.docx/.xlsx/.pptx)"),
             ("--category", "accessibility|quality|all (default all)"),
             ("--severity", "error|warning|info — minimum level to report (default info)"),
             ("--fix", "apply only safe autofixes (alt text, table header, doc/slide title, orphan bookmark); result adds {fixed, remaining}. Codes + autofix table: aioffice help audit")],
            [ErrorCodes.InvalidArgs, ErrorCodes.FileNotFound, ErrorCodes.SandboxDenied, ErrorCodes.FormatCorrupt, ErrorCodes.UnsupportedFeature],
            [("aioffice audit report.docx", "every finding with code + path + summary{errors,warnings,infos}"),
             ("aioffice audit report.docx --fix", "fixes alt/header/title, reports {fixed:N, remaining:[…]}"),
             ("aioffice audit metrics.xlsx --category accessibility --severity warning", "only a11y findings at warning+")]),

        Verb("template", "Fill {{key}} placeholders in text runs from a JSON merge map.", "office_template",
            [("file", "template document containing {{key}} placeholders"),
             ("--data", "JSON object of string values, or @data.json"),
             ("-o", "result path (recommended); omit = merge in place with auto-snapshot"),
             ("--overwrite", "replace an existing output file")],
            [ErrorCodes.InvalidArgs, ErrorCodes.FileNotFound, ErrorCodes.SandboxDenied, ErrorCodes.FormatCorrupt],
            [("aioffice template contract.docx --data @acme.json -o out/acme.docx", "check data.leftoverPlaceholders afterwards")]),

        Verb("diff", "Semantically compare a document against a baseline (another same-format file, or one of its own snapshots): a sorted, deterministic change list. Changes are data — ok:true / exit 0.", "office_diff",
            [("file", "the current/new document"),
             ("otherFile", "baseline: the OLD same-format document (sandbox-resolved); a format mismatch is invalid_args. Mutually exclusive with --snapshot"),
             ("--snapshot", "baseline: snapshot N of <file> from its pre-edit ring; a missing index is invalid_args naming the available numbers. Mutually exclusive with otherFile"),
             ("--view", "summary | detailed (default): summary trims to counts + path+kind per change, detailed keeps full before/after")],
            [ErrorCodes.InvalidArgs, ErrorCodes.FileNotFound, ErrorCodes.SandboxDenied, ErrorCodes.FormatCorrupt, ErrorCodes.UnsupportedFeature],
            [("aioffice diff new.docx old.docx", "every added/removed/modified/moved block, sorted by (path, kind)"),
             ("aioffice diff report.docx --snapshot 1", "what changed since the last edit (snapshot 1 of the file's own ring)"),
             ("aioffice diff metrics.xlsx baseline.xlsx --view summary", "just the counts + a terse path+kind line per change")]),

        Verb("snapshot", "List or restore the automatic pre-edit snapshot ring (20 per file).", "file_snapshot",
            [("action", "list | restore (default list)"),
             ("file", "document path"),
             ("n", "restore: snapshot number from list (omit = latest)")],
            [ErrorCodes.InvalidArgs, ErrorCodes.FileNotFound, ErrorCodes.SandboxDenied],
            [("aioffice snapshot restore report.docx 7", "restore is itself snapshotted — undo is undoable")]),

        Verb("doctor", "Health check: runtime, workspace sandbox, snapshot store.", "office_status",
            [],
            [ErrorCodes.InternalError],
            [("aioffice doctor", "run when commands start failing oddly")]),

        Verb("schema", "Machine-readable JSON of this whole command surface.", "office_schema",
            [("verb", "optional filter: one of the 16 verb names")],
            [ErrorCodes.InvalidArgs],
            [("aioffice schema edit", "just the edit verb")]),

        Verb("help", "Progressive docs: addressing, selectors, per-element property tables.", "office_help",
            [("topic", "omit for the index; e.g. addressing, selectors, docx/paragraph")],
            [ErrorCodes.InvalidArgs],
            [("aioffice help selectors", "full selector grammar")]),

        Verb("preview", "Live browser preview: open serves the document on localhost (click = select), selection reads the clicked paths, close stops the server.", "preview_open",
            [("action", "open (blocking server; prints {url,port,pid} before blocking) | selection | close"),
             ("file", "document to preview (.docx/.xlsx/.pptx)"),
             ("--port", "fixed port for open (default: first free port in 26500-26600)")],
            [ErrorCodes.InvalidArgs, ErrorCodes.PreviewNotRunning, ErrorCodes.FileNotFound,
             ErrorCodes.SandboxDenied, ErrorCodes.UnsupportedFeature],
            [("aioffice preview open report.docx", "blocks; click elements in the browser, then read them"),
             ("aioffice preview selection report.docx", "the clicked canonical paths, ready for get/edit")]),

        Verb("mcp", "Start the stdio MCP server exposing the same capabilities as the CLI (16 tools).", null,
            [("--workspace", "sandbox root (default cwd; also AIOFFICE_WORKSPACE)")],
            [ErrorCodes.InternalError],
            [("aioffice mcp --workspace ~/docs", "stdio transport; one JSON envelope per tool result")]),
    ];

    /// <summary>The 16 verb names, in surface order.</summary>
    public static IReadOnlyList<string> VerbNames { get; } = [.. Verbs.Select(v => v.Name)];

    /// <summary>
    /// Builds the <c>office_schema</c>/<c>aioffice schema</c> data payload, optionally
    /// filtered to one verb. Unknown verbs throw <c>invalid_args</c> with candidates.
    /// </summary>
    public static object Build(string? verb)
    {
        if (string.IsNullOrWhiteSpace(verb))
        {
            return new { version = Meta.ToolVersion, verbs = Verbs };
        }

        var match = Verbs.FirstOrDefault(v => v.Name.Equals(verb.Trim(), StringComparison.OrdinalIgnoreCase));
        if (match is null)
        {
            throw new AiofficeException(
                ErrorCodes.InvalidArgs,
                $"Unknown verb: '{verb}'.",
                "Call office_schema with no verb to list the whole surface.",
                candidates: VerbNames);
        }

        return new { version = Meta.ToolVersion, verbs = new[] { match } };
    }
}
