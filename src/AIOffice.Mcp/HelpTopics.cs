using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// The progressive documentation behind <c>office_help</c> / <c>aioffice help</c>.
/// Everything too big for tool schemas (property tables, selector grammar,
/// addressing details) lives here so the always-resident schema surface stays
/// inside its token budget.
/// </summary>
public static class HelpTopics
{
    private const string IndexTopic = "index";

    private static readonly IReadOnlyDictionary<string, (string Doc, string[] Related)> Topics =
        new Dictionary<string, (string, string[])>(StringComparer.OrdinalIgnoreCase)
        {
            [IndexTopic] = (
                """
                ## aioffice help — topics
                - addressing      path grammar for docx/xlsx/pptx (1-based)
                - selectors       CSS-like query grammar for office_query
                - edit-ops        op kinds, position forms, atomicity, expect_rev
                - envelope        the {ok,data,error,meta} result shape
                - errors          all error codes and how to recover
                - docx/paragraph  paragraph element: settable props
                - xlsx/cell       cell element: settable props
                - pptx/shape      shape element: settable props
                Call office_help {topic:"<name>"} (CLI: aioffice help <name>).
                """,
                ["addressing", "selectors", "edit-ops"]),

            ["addressing"] = (
                """
                ## Addressing grammar (1-based indices everywhere)
                docx:  /body/p[3]   /body/table[1]/tr[2]/tc[1]   /body/p[3]/run[2]   /header[1]/p[1]
                xlsx:  /Sheet1/A1   /Sheet1/A1:C10   /Sheet1/row[3]   sheet names with spaces or specials are quoted: /'Q3 Data'/B2 (escape ' as '')
                pptx:  /slide[2]   /slide[2]/shape[3]   /slide[2]/shape[3]/p[1]   (master/layout addressing is reserved for M1)
                "/" alone addresses document-level properties (office_get and office_edit set only).
                office_query returns canonical paths; office_get / office_edit accept them verbatim.
                Positional indices DRIFT after inserts/removes — re-run office_query instead of reusing old indices.
                """,
                ["selectors", "edit-ops"]),

            ["selectors"] = (
                """
                ## Selector grammar (office_query)
                form:       element predicates*        element = p | run | table | cell | row | sheet | slide | shape | image | * (any)
                predicates: [attr OP value]            OP = "=", "!=", ">", ">=", "<", "<=", "*=" (text contains)
                            :contains('text')          escape ' inside the quotes as ''
                examples:   p[style=Heading1]          cell[value>100]          shape:contains('Q3')
                            run[bold=true]             p:contains('budget')     *[style=Caption]
                Numeric comparisons apply when the value parses as a number; otherwise ordinal text comparison.
                xlsx note: value predicates on formula cells without a cached value add a formula_not_evaluated warning to meta.warnings.
                Results carry canonical paths (max 50; data.total is the full hit count, data.truncated says if the list was cut).
                """,
                ["addressing", "errors"]),

            ["edit-ops"] = (
                """
                ## office_edit ops — the only mutation surface
                kinds: set (change props), add (insert child; path = PARENT; needs type), remove (delete), move (reposition).
                position (add/move): 1-based integer within parent | "before:<path>" | "after:<path>" | omit = append.
                props are string-valued: {"text":"Hi","bold":"true","size":"12pt","fill":"FF0000"}; sizes unit-qualified (12pt, 2cm); colors hex or named.
                Atomicity: the whole ops[] batch is one save — any failure means NOTHING is written.
                Ops resolve at execution time: an add shifts the indices later ops see. Put set/remove targets BEFORE inserts, or split into two edits and re-query.
                expect_rev: pass the last meta.rev; mismatch -> stale_address before any write. dry_run:true validates without writing.
                Every successful edit auto-snapshots the pre-image (file_snapshot lists/restores).
                """,
                ["docx/paragraph", "xlsx/cell", "pptx/shape", "addressing"]),

            ["envelope"] = (
                """
                ## Result envelope — every tool, every time
                { "ok": bool, "data": {...}|null, "error": {code,message,suggestion,candidates?}|null,
                  "meta": { file?, rev?, elapsedMs, version, warnings?: [{code,message}] } }
                error.suggestion is never empty — it is the next action to take.
                rev = first 12 hex of SHA-256 of the file bytes AFTER the call; feed it to office_edit expect_rev.
                meta.warnings carries non-fatal issues (e.g. formula_not_evaluated); check it even when ok=true.
                """,
                ["errors"]),

            ["errors"] = (
                """
                ## Error codes
                invalid_args         bad parameter/selector/topic/snapshot number; suggestion shows the correct shape, candidates list near matches
                file_not_found       check spelling relative to the workspace root, or office_create it
                sandbox_denied       path escapes the workspace (symlinks count); widen with --workspace / AIOFFICE_WORKSPACE
                invalid_path         element path does not resolve; candidates[] holds the nearest canonical paths — pick one
                stale_address        expect_rev mismatch: file changed since you read it; re-read, then retry (nothing was written)
                unsupported_feature  capability not in this milestone; suggestion names the workaround
                format_corrupt       not valid OOXML; try office_validate and file_snapshot restore
                internal_error       our bug; report with office_status output
                formula_not_evaluated  WARNING in meta.warnings, not an error: formula text returned instead of a value
                """,
                ["envelope"]),

            ["docx/paragraph"] = (
                """
                ## docx paragraph — props (office_edit set/add, type "paragraph")
                text     replaces the paragraph's text content
                style    named style, e.g. Heading1, Heading2, Normal, Quote
                bold / italic / underline   "true" | "false"
                size     unit-qualified, e.g. "12pt"
                color    hex "FF0000" or named
                align    left | center | right | justify
                add child runs with type "run" for mixed formatting inside one paragraph.
                """,
                ["edit-ops", "addressing"]),

            ["xlsx/cell"] = (
                """
                ## xlsx cell — props (office_edit set; address cells as /Sheet1/B2)
                value         cell value; numbers/dates auto-typed, anything else stays text
                formula       e.g. "=SUM(A1:A10)" (calculated on save where supported)
                numberFormat  e.g. "0.00", "yyyy-mm-dd"
                bold / italic "true" | "false"
                fill          background color, hex "FFFF00"
                office_get returns value, valueType, formula, numberFormat. Reading a formula cell without
                a cached value adds a formula_not_evaluated warning and returns the formula text.
                """,
                ["edit-ops", "addressing"]),

            ["pptx/shape"] = (
                """
                ## pptx shape — props (office_edit set/add, type "shape")
                text          replaces the shape's text content
                x / y / w / h position and size, unit-qualified ("2cm", "1.5in") or EMU integers
                fill          hex "4472C4" or named
                fontSize      unit-qualified, e.g. "18pt"
                bold / italic "true" | "false"
                add slides with type "slide" on path "/"; render→look→fix: office_render {to:"svg", scope:"/slide[N]"} after each visual change.
                """,
                ["edit-ops", "addressing"]),
        };

    /// <summary>All topic names (index first, then alphabetical).</summary>
    public static IReadOnlyList<string> Names { get; } =
        [IndexTopic, .. Topics.Keys.Where(k => k != IndexTopic).Order(StringComparer.Ordinal)];

    /// <summary>
    /// Resolves a topic (null/empty → index) to the <c>office_help</c> data payload,
    /// or throws <c>invalid_args</c> with nearest-match candidates.
    /// </summary>
    public static object Get(string? topic)
    {
        var key = string.IsNullOrWhiteSpace(topic) ? IndexTopic : topic.Trim();
        if (Topics.TryGetValue(key, out var entry))
        {
            return new { topic = key.ToLowerInvariant(), doc = entry.Doc, related = entry.Related };
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Unknown help topic: '{topic}'.",
            "Call office_help with no topic to list all topics.",
            candidates: Nearest(key));
    }

    private static IReadOnlyList<string> Nearest(string requested)
    {
        var hits = Names
            .Where(n => n.Contains(requested, StringComparison.OrdinalIgnoreCase) ||
                        requested.Contains(n, StringComparison.OrdinalIgnoreCase))
            .ToList();
        return hits.Count > 0 ? hits : Names;
    }
}
