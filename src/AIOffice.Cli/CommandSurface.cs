using System.Text.Json.Serialization;
using AIOffice.Core;

namespace AIOffice.Cli;

/// <summary>One positional argument of a verb (for schema/help).</summary>
public sealed record PositionalSpec(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("required")] bool Required,
    [property: JsonPropertyName("description")] string Description);

/// <summary>One option/flag of a verb (for schema/help).</summary>
public sealed record OptionSpec(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("value")] string? Value,
    [property: JsonPropertyName("description")] string Description,
    [property: JsonPropertyName("repeatable")] bool Repeatable = false);

/// <summary>One verb of the surface. CLI verbs mirror the MCP tools 1:1.</summary>
public sealed record VerbSpec(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("summary")] string Summary,
    [property: JsonPropertyName("usage")] string Usage,
    [property: JsonPropertyName("positionals")] IReadOnlyList<PositionalSpec> Positionals,
    [property: JsonPropertyName("options")] IReadOnlyList<OptionSpec> Options,
    [property: JsonPropertyName("mcpTool")] string? McpTool);

/// <summary>
/// The single source of truth for the aioffice command surface (v0). The verb
/// dispatcher, <c>schema</c> and <c>help</c> all read from this table so the
/// documented surface can never drift from the implemented one.
/// </summary>
public static class CommandSurface
{
    public static readonly IReadOnlyList<OptionSpec> GlobalFlags =
    [
        new("json", null, "Force compact JSON output (the default when stdout is not a TTY)."),
        new("pretty", null, "Pretty-print the JSON envelope."),
        new("workspace", "<dir>", "Sandbox root for all file access (default: current directory; also AIOFFICE_WORKSPACE)."),
        new("quiet", null, "Suppress the success envelope on stdout; errors are still printed. Exit codes are unaffected."),
    ];

    public static readonly IReadOnlyList<VerbSpec> Verbs =
    [
        new("create",
            "Create a new empty document (kind inferred from the file extension), or import one with --from (markdown→docx, csv→xlsx).",
            "aioffice create <file> [--from notes.md|data.csv] [--kind docx|xlsx|pptx] [--title T]",
            [new("file", true, "Path of the document to create, inside the workspace.")],
            [
                new("from", "<source>", "Import source: .md/.markdown → .docx (headings, lists, tables, links, code), .csv/.tsv → .xlsx (typed cells). Mismatched pairs fail with the valid matrix."),
                new("kind", "docx|xlsx|pptx", "Override the document kind when the extension is ambiguous."),
                new("title", "<text>", "Document title (core properties / first slide title / first sheet name)."),
            ],
            "office_create"),

        new("read",
            "Read a document as outline, plain text, stats, structure, properties or embeds (docx adds revisions/comments/styles/markdown/fields; xlsx adds csv/styles).",
            "aioffice read <file> [--view outline|text|stats|structure|properties|embeds|revisions|comments|styles|fields|markdown|csv] [--range a..b] [--sheet NAME] [--max-bytes N]",
            [new("file", true, "Document to read.")],
            [
                new("view", "outline|text|stats|structure|properties|embeds|revisions|comments|styles|fields|markdown|csv",
                    "Projection to return (default: outline). properties = core + custom document properties (docx/xlsx/pptx); embeds = embedded OLE/package objects (docx/xlsx/pptx); comments works on all three formats. revisions/markdown/fields are docx views (fields = content controls); styles is docx (style defs) or xlsx (named cell styles); csv is xlsx-only (one sheet, RFC 4180)."),
                new("range", "a..b", "Limit to an element range, e.g. paragraphs 3..10 or slides 1..4; the csv view also takes a cell range like A1:C10."),
                new("sheet", "<name>", "csv view: which sheet to emit (default: the first sheet)."),
                new("max-bytes", "N", "Truncate the response payload to at most N bytes."),
            ],
            "office_read"),

        new("query",
            "Find elements with a CSS-like selector; returns stable canonical paths.",
            "aioffice query <file> <selector>",
            [
                new("file", true, "Document to query."),
                new("selector", true, "Selector, e.g. p[style=Heading1], cell[value>100], shape:contains('Q3')."),
            ],
            [],
            "office_query"),

        new("get",
            "Get one node and its properties by canonical path.",
            "aioffice get <file> <path>",
            [
                new("file", true, "Document to inspect."),
                new("path", true, "Canonical path, e.g. /body/p[3] or /Sheet1/A1."),
            ],
            [],
            "office_get"),

        new("edit",
            "Apply an atomic batch of operations (set/add/remove/move/replace/accept/reject/extract). Auto-snapshots the pre-image.",
            "aioffice edit <file> --ops <json|@file> | --set <path> k=v... | --add <path> --type T k=v... | --remove <path> | --find X --replace Y [--regex] [--match-case] [--whole-word] [--track] [--author NAME] [--dry-run] [--expect-rev R]",
            [
                new("file", true, "Document to edit."),
                new("k=v", false, "Property assignments for the --set/--add sugar forms."),
            ],
            [
                new("ops", "<json|@file>", "JSON array of operations: [{op:set|add|remove|move|replace|accept|reject|extract, path, type?, props?, position?}]. extract writes an embedded object's bytes to a sandbox destination (props.to) — a producing op that does not modify the source."),
                new("set", "<path>", "Sugar: set properties (from trailing k=v pairs) on the node at <path>."),
                new("add", "<path>", "Sugar: add a new node of --type at <path> (trailing k=v pairs become props)."),
                new("type", "<element>", "Element type for --add, e.g. p, table, slide, shape, image, comment, reply, style, pivot, conditionalFormat, toc, watermark, endnote, sectionBreak, animation, note, row, col, field, dataValidation, sparkline, caption (docx), crossRef (docx), slicer (xlsx)."),
                new("remove", "<path>", "Sugar: remove the node at <path>.", Repeatable: true),
                new("find", "<text>", "Sugar: document-wide find/replace (docx body+headers+footers, every sheet, every slide incl. notes); aggregate {replacements, locations} in the result."),
                new("replace", "<text>", "Replacement text for --find (default: empty = delete matches); in --regex mode $1 etc. substitute groups."),
                new("regex", null, ".NET regex matching for --find (2s match budget; timeout -> invalid_args)."),
                new("match-case", null, "Case-sensitive --find matching (default: case-insensitive)."),
                new("whole-word", null, "Match --find at word boundaries only."),
                new("position", "at|before|after|inside", "Placement for the --add sugar (format-specific; default: append)."),
                new("track", null, "docx: record text set/add/remove/replace ops as tracked revisions (w:ins/w:del); resolve later with accept/reject ops."),
                new("author", "<name>", "Author stamped on revisions/comments (op props.author wins; default: AIOFFICE_AUTHOR env, then 'AIOffice')."),
                new("dry-run", null, "Validate and report what would change without writing."),
                new("expect-rev", "R", "Fail with stale_address before any write unless the file's rev is R."),
            ],
            "office_edit"),

        new("render",
            "Render a document (or a scoped node) to html, svg, text, png or pdf.",
            "aioffice render <file> [--to html|svg|text|png|pdf] [--scope <path>] [-o out]",
            [new("file", true, "Document to render.")],
            [
                new("to", "html|svg|text|png|pdf", "Output format. html for docx/xlsx, svg or html per slide for pptx; png screenshots via a local Chromium (pptx: one slide, default /slide[1] — use --scope); pdf prints paged output via the same Chromium (pptx: the whole deck, one page per slide)."),
                new("scope", "<path>", "Render only the node at this path, e.g. /slide[2]."),
                new("o", "<file>", "Write the rendering to this file instead of inlining it in the envelope (png/pdf default: source path with .png/.pdf)."),
            ],
            "office_render"),

        new("validate",
            "Validate OOXML conformance plus aioffice lint rules; returns issues with suggestions.",
            "aioffice validate <file>",
            [new("file", true, "Document to validate.")],
            [],
            "office_validate"),

        new("template",
            "Merge {{key}} placeholders in text runs with JSON data.",
            "aioffice template <file> --data <json|@file> [-o out]",
            [new("file", true, "Template document containing {{key}} placeholders.")],
            [
                new("data", "<json|@file>", "JSON object with the merge values."),
                new("o", "<file>", "Write the merged document here (default: in place)."),
            ],
            "office_template"),

        new("audit",
            "Audit a document for accessibility + quality findings (findings are data, exit 0); --fix applies the safe autofixes.",
            "aioffice audit <file> [--category accessibility|quality|all] [--severity error|warning|info] [--fix]",
            [new("file", true, "Document to audit (.docx/.xlsx/.pptx).")],
            [
                new("category", "accessibility|quality|all", "Which checks to run (default: all). accessibility = alt text, headings, table headers, contrast, titles, reading order, tiny fonts, merged data cells; quality = broken refs/links, formula errors, empty/duplicate ids, off-canvas, empty placeholders."),
                new("severity", "error|warning|info", "Minimum severity to report (default: info reports everything; warning hides info; error hides info+warning)."),
                new("fix", null, "Apply only the safe, non-destructive autofixes (placeholder alt text, mark a table header row, set a doc/slide title, drop an orphan bookmark). Reports {fixed:N, remaining:[…]}. See 'aioffice help audit' for the codes and which are autofixable."),
            ],
            "office_audit"),

        new("diff",
            "Semantically compare a document against a baseline (another file, or one of its own snapshots): a sorted, deterministic change list.",
            "aioffice diff <file> [<otherFile>] [--snapshot N] [--view summary|detailed]",
            [
                new("file", true, "The current/new document."),
                new("otherFile", false, "Baseline: the OLD document to compare against (same format). Mutually exclusive with --snapshot."),
            ],
            [
                new("snapshot", "N", "Baseline: snapshot N of <file> from its automatic pre-edit ring (restored to a temp file). Mutually exclusive with otherFile."),
                new("view", "summary|detailed", "detailed (default) carries full before/after per change; summary trims to counts + one terse path+kind line per change."),
            ],
            "office_diff"),

        new("convert",
            "Convert a document between formats: docx/xlsx/pptx ↔ each other (via a content-neutral model), docx↔md, xlsx↔csv, any→pdf/png/svg/html. Inherently lossy — what didn't survive is named in a convert_lossy warning.",
            "aioffice convert <src> <dest> [--json]",
            [
                new("src", true, "Source document (.docx/.xlsx/.pptx/.md/.csv); read-only."),
                new("dest", true, "Destination, created fresh: .docx/.xlsx/.pptx/.md/.csv (content) or .pdf/.png/.svg/.html/.txt (render). Same extension as src is invalid_args (use edit)."),
            ],
            [],
            "office_convert"),

        new("snapshot",
            "List or restore the automatic pre-edit snapshot ring (keeps the last 20 per file).",
            "aioffice snapshot <list|restore> <file> [n]",
            [
                new("action", true, "list or restore."),
                new("file", true, "The document whose snapshot ring to use."),
                new("n", false, "Snapshot number to restore (default: the newest)."),
            ],
            [],
            "file_snapshot"),

        new("doctor",
            "Diagnose the environment: runtime, workspace, handlers, dependencies.",
            "aioffice doctor",
            [],
            [],
            "office_status"),

        new("schema",
            "Machine-readable JSON of the whole command surface (agents introspect instead of guessing).",
            "aioffice schema [verb]",
            [new("verb", false, "Limit the schema to one verb.")],
            [],
            "office_schema"),

        new("help",
            "Progressive documentation: addressing grammar, selectors, per-format properties, errors.",
            "aioffice help [topic]",
            [new("topic", false, "One of: addressing, selectors, properties-docx, properties-xlsx, properties-pptx, errors, equations, embeds, rtl, sections, audit, diff, convert (or any verb name). Omit for the index.")],
            [],
            "office_help"),

        new("preview",
            "Live browser preview: open serves the document on localhost (click = select), selection reads the clicked paths, close stops the server.",
            "aioffice preview <open|selection|close> <file> [--port N]",
            [
                new("action", true, "open (blocking server; prints {url,port,pid} first), selection or close."),
                new("file", true, "Document to preview (.docx/.xlsx/.pptx)."),
            ],
            [new("port", "N", "Fixed port for open (default: first free port in 26500-26600).")],
            "preview_open"),

        new("mcp",
            "Run the stdio MCP server exposing the same 17 tools as the CLI.",
            "aioffice mcp",
            [],
            [],
            null),

        new("version",
            "Print the aioffice version.",
            "aioffice version",
            [],
            [],
            null),
    ];

    /// <summary>Finds a verb spec by exact name.</summary>
    public static VerbSpec? Find(string name) =>
        Verbs.FirstOrDefault(v => string.Equals(v.Name, name, StringComparison.Ordinal));

    /// <summary>All verb names, for candidates lists.</summary>
    public static IReadOnlyList<string> VerbNames => [.. Verbs.Select(v => v.Name)];

    /// <summary>Nearest verbs by edit distance, for invalid verb candidates.</summary>
    public static IReadOnlyList<string> NearestVerbs(string attempted)
    {
        return [.. Verbs
            .Select(v => (v.Name, Distance: Levenshtein(attempted, v.Name)))
            .OrderBy(p => p.Distance)
            .ThenBy(p => p.Name, StringComparer.Ordinal)
            .Take(3)
            .Select(p => p.Name)];
    }

    private static int Levenshtein(string a, string b)
    {
        var previous = new int[b.Length + 1];
        var current = new int[b.Length + 1];
        for (var j = 0; j <= b.Length; j++)
        {
            previous[j] = j;
        }

        for (var i = 1; i <= a.Length; i++)
        {
            current[0] = i;
            for (var j = 1; j <= b.Length; j++)
            {
                var substitution = a[i - 1] == b[j - 1] ? 0 : 1;
                current[j] = Math.Min(
                    Math.Min(current[j - 1] + 1, previous[j] + 1),
                    previous[j - 1] + substitution);
            }

            (previous, current) = (current, previous);
        }

        return previous[b.Length];
    }
}

/// <summary>Static descriptions of the addressing/selector grammars surfaced by <c>schema</c>.</summary>
public static class GrammarPointers
{
    public static readonly object Addressing = new
    {
        summary = "Slash-separated 1-based paths. docx: /body/p[3], /body/table[1]/tr[2]/tc[1], /header[1]/p[1], " +
                  "/revision[@id=3], /comment[@id=1], /style[@id=Callout], /caption[@label=Figure][1], " +
                  "/body/p[3]/omath[1], /embed[1], /properties. " +
                  "xlsx: /Sheet1/A1, /Sheet1/A1:C10, /Sheet1/row[3], /'Q3 Data'/B2, /Pivot/pivot[@name=SalesPivot], " +
                  "/Sheet1/conditionalFormat[1], /Sheet1/image[1], /Sheet1/slicer[1], /Sheet1/embed[1], /properties. " +
                  "pptx: /slide[2], /slide[2]/shape[3], /slide[2]/shape[3]/p[1], /slide[2]/notes, " +
                  "/slide[2]/shape[@id=9]/omath[1], /slide[2]/embed[@id=7], /properties.",
        helpTopic = "addressing",
        examples = new[]
        {
            "/body/p[3]", "/revision[@id=3]", "/comment[@id=1]", "/style[@id=Callout]",
            "/caption[@label=Figure][1]", "/body/p[3]/omath[1]", "/embed[1]", "/Sheet1/A1:C10",
            "/Pivot/pivot[@name=SalesPivot]", "/Sheet1/conditionalFormat[1]", "/Sheet1/slicer[1]",
            "/Sheet1/embed[1]", "/slide[2]/shape[3]", "/slide[2]/notes",
            "/slide[2]/shape[@id=9]/omath[1]", "/slide[2]/embed[@id=7]",
        },
    };

    public static readonly object Selectors = new
    {
        summary = "element[attr OP value]:contains('text') where OP is = != > >= < <= *= and element may be *. " +
                  "query returns canonical paths that get/edit accept.",
        helpTopic = "selectors",
        examples = new[]
        {
            "p[style=Heading1]", "cell[value>100]", "shape:contains('Q3')", "*[name*=Total]",
        },
    };
}
