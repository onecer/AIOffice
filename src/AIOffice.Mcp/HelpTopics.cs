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
                - docx/revisions  tracked changes: track:true, accept/reject ops
                - docx/comment    comments + threaded replies: add/read/remove
                - docx/style      style definitions: add/set/remove, apply
                - docx/image      inline pictures: add type "image"
                - docx/list       numbered/bulleted/nested list items (M3)
                - docx/link       hyperlinks, bookmarks, footnotes (M3)
                - docx/section    page size/orientation/margins (M3)
                - xlsx/cell       cell element: settable props
                - xlsx/pivot      pivot tables: add type "pivot"
                - xlsx/conditionalFormat  cellIs/colorScale/dataBar/containsText
                - xlsx/image      anchored pictures: add type "image"
                - xlsx/sheet      freeze/autoFilter/print setup, defined names, streaming (M3)
                - pptx/shape      shape element: settable props incl. preset geometry/z-order
                - pptx/slide      slide props incl. background + transition
                - pptx/notes      speaker notes: /slide[i]/notes
                - pptx/image      pictures: add type "image"
                - pptx/chart      native charts + cross-doc dataFrom (M3)
                Call office_help {topic:"<name>"} (CLI: aioffice help <name>).
                """,
                ["addressing", "selectors", "edit-ops"]),

            ["addressing"] = (
                """
                ## Addressing grammar (1-based indices everywhere)
                docx:  /body/p[3]   /body/table[1]/tr[2]/tc[1]   /body/p[3]/run[2]   /header[1]/p[1]   /footer[1]/p[1]
                xlsx:  /Sheet1/A1   /Sheet1/A1:C10   /Sheet1/row[3]   /Sheet1/chart[1]   sheet names with spaces or specials are quoted: /'Q3 Data'/B2 (escape ' as '')
                pptx:  /slide[2]   /slide[2]/shape[3]   /slide[2]/shape[@id=7] (stable-id form, canonical in results)   /slide[2]/shape[3]/p[1]
                       /master[1]   /master[1]/layout[2]   /master[1]/shape[1]   (masters/layouts are read-only: get/query; editing them is M2)
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
                file_too_large       file exceeds the opt-in AIOFFICE_MAX_FILE_MB cap (default unlimited); raise/unset it or split the file
                format_corrupt       not valid OOXML; try office_validate and file_snapshot restore
                internal_error       our bug; report with office_status output
                formula_not_evaluated  WARNING in meta.warnings, not an error: formula text returned instead of a value
                """,
                ["envelope"]),

            ["docx/revisions"] = (
                """
                ## docx tracked changes (M2)
                Record: office_edit {track:true, author:"Reviewer", ops:[…]} — text set/add/remove ops become w:ins/w:del.
                Only TEXT changes can be tracked ({"props":{"text":…}}); tracked formatting -> invalid_args naming the workaround.
                Author resolution: op props.author > tool arg author > AIOFFICE_AUTHOR env > "AIOffice". Date = now UTC.
                Read:    office_read {view:"revisions"} -> [{id, kind:insert|delete, author, date, text, path}].
                Resolve: {"op":"accept","path":"/revision[@id=3]"} applies one; path "/body" resolves every revision in scope.
                         {"op":"reject",…} undoes instead. Revisions are never "removed" — only accepted or rejected.
                """,
                ["docx/comment", "edit-ops"]),

            ["docx/comment"] = (
                """
                ## docx comments (M2)
                Add:    {"op":"add","path":"/body/p[2]","type":"comment","props":{"text":"…","author":"Reviewer"?}}
                        (path = anchored content: a paragraph or a run; result reports /comment[@id=N] + anchor).
                Read:   office_read {view:"comments"} -> [{id, author, date, text, anchorPath, anchorText}].
                Get:    office_get {path:"/comment[@id=2]"}. Remove: {"op":"remove","path":"/comment[@id=2]"}.
                """,
                ["docx/revisions", "edit-ops"]),

            ["docx/style"] = (
                """
                ## docx style definitions (M2)
                Add:    {"op":"add","path":"/styles","type":"style","props":{"id":"Callout","kind":"paragraph"?,
                        "name"?,"basedOn"?,"bold"?,"italic"?,"underline"?,"color"?,"fontSize"?,"alignment"?,
                        "spacingBefore"?,"spacingAfter"?}}
                Modify: {"op":"set","path":"/style[@id=Callout]","props":{…}} (id/kind are fixed).
                Apply:  {"op":"set","path":"/body/p[2]","props":{"style":"Callout"}}.
                Read:   office_read {view:"styles"}; office_get {path:"/style[@id=Callout]"}.
                Remove: {"op":"remove","path":"/style[@id=Callout]"} (custom styles only; built-ins are modified, not removed).
                """,
                ["docx/paragraph", "edit-ops"]),

            ["docx/image"] = (
                """
                ## docx inline pictures (M2)
                Add: {"op":"add","path":"/body","type":"image","props":{"src":"logo.png","width":"10cm"?,"height":"3cm"?}}
                     position "before:<path>"/"after:<path>" places it; omit = append. PNG/JPEG; src resolves through the
                     workspace sandbox (escaping it -> sandbox_denied). Omitted width/height keep the natural aspect ratio.
                """,
                ["edit-ops", "addressing"]),

            ["xlsx/pivot"] = (
                """
                ## xlsx pivot tables (M2)
                Add: {"op":"add","path":"/Sheet1","type":"pivot","props":{
                       "sourceRange":"A1:E7","targetSheet":"Pivot","name"?,"targetAnchor"?,
                       "rows":["Region"],"columns"?,"filters"?,
                       "values":[{"field":"Sales","agg":"sum|average|count|min|max"}]}}
                     path = SOURCE sheet; sourceRange needs a header row. targetSheet is created when absent.
                Get: office_get {path:"/Pivot/pivot[1]"} or {path:"/Pivot/pivot[@name=SalesPivot]"} (the TARGET sheet).
                Remove: {"op":"remove","path":"/Pivot/pivot[@name=SalesPivot]"}. Excel recomputes on open (refreshOnLoad).
                """,
                ["xlsx/cell", "edit-ops"]),

            ["xlsx/conditionalFormat"] = (
                """
                ## xlsx conditional formatting (M2)
                Add (path = the range): {"op":"add","path":"/Sheet1/A1:C10","type":"conditionalFormat","props":{…}}
                kinds: cellIs      {operator:"> >= < <= == != between", value, value2 (between only), fill?, color?, bold?}
                       colorScale  {minColor, maxColor, midColor?}   dataBar {color}   containsText {text, fill?, color?, bold?}
                Get: office_get {path:"/Sheet1/conditionalFormat[1]"}; remove by the same path (later indices shift down).
                """,
                ["xlsx/cell", "edit-ops"]),

            ["xlsx/image"] = (
                """
                ## xlsx anchored pictures (M2)
                Add: {"op":"add","path":"/Sheet1","type":"image","props":{"src":"logo.png","anchor":"E2","name"?,
                     "widthPx"?,"heightPx"?}} — PNG/JPEG; src resolves through the workspace sandbox.
                Get: office_get {path:"/Sheet1/image[1]"} -> name, format, anchor, size. Remove by the same path.
                """,
                ["xlsx/cell", "edit-ops"]),

            ["pptx/slide"] = (
                """
                ## pptx slide (incl. M2 background, M3 transition)
                Add:        {"op":"add","path":"/slide[3]","type":"slide","props":{"title"?,"background"?}} — becomes slide 3.
                Background: {"op":"set","path":"/slide[1]","props":{"background":"0F172A"}} — a real p:bg solid fill (hex).
                Transition: {"op":"set","path":"/slide[1]","props":{"transition":"fade","transitionDuration":"0.5s"}} (see pptx/transition).
                Move/remove a slide by its /slide[i] path. render {to:"svg"|"png", scope:"/slide[N]"} to look after edits;
                render {to:"pdf"} prints the whole deck, one page per slide.
                """,
                ["pptx/shape", "pptx/notes", "pptx/transition"]),

            ["pptx/notes"] = (
                """
                ## pptx speaker notes (M2)
                Path /slide[i]/notes addresses the whole notes body (no /p[j] beneath it).
                Set:    {"op":"set","path":"/slide[2]/notes","props":{"text":"line1\nline2"}} (replaces; \n = new paragraph)
                Append: {"op":"add","path":"/slide[2]/notes","props":{"text":"follow-up"}}
                Remove: {"op":"remove","path":"/slide[2]/notes"}. office_read outline + office_get /slide[i]/notes read them back.
                """,
                ["pptx/slide", "edit-ops"]),

            ["pptx/image"] = (
                """
                ## pptx pictures (M2)
                Add: {"op":"add","path":"/slide[1]","type":"image","props":{"src":"logo.png","x"?,"y"?,"w"?,"h"?,"name"?}}
                     sizes unit-qualified ("6cm","1.5in") or EMU; omit w/h to keep natural size/aspect. PNG/JPEG;
                     src resolves through the workspace sandbox. Result path: /slide[1]/shape[@id=N] (remove by that path).
                """,
                ["pptx/slide", "edit-ops"]),

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
                shape (M3)    preset geometry on add: rect|roundRect|ellipse|triangle|diamond|arrow|line ("line" builds a connector; fill = stroke)
                flipH / flipV (M3) "true" mirrors the geometry
                z-order (M3): {"op":"move","path":"/slide[1]/shape[@id=5]","position":"front|back|forward|backward"} (front = painted last/topmost)
                add slides with type "slide" on path "/"; render→look→fix: office_render {to:"svg", scope:"/slide[N]"} after each visual change.
                """,
                ["edit-ops", "addressing"]),

            ["docx/list"] = (
                """
                ## docx lists (M3) — add type "p" with list props
                list         bullet | number
                level        0-based nesting (default 0)
                listRestart  "true" restarts a number list at this item
                {"op":"add","path":"/body","type":"p","props":{"text":"Step one","list":"number"}}
                Nested: {"props":{"text":"Detail","list":"number","level":"1"}}.
                read {view:"text"} shows "1." / "•" markers; render {to:"html"} emits real <ol>/<ul>.
                """,
                ["docx/paragraph", "edit-ops"]),

            ["docx/link"] = (
                """
                ## docx hyperlinks / bookmarks / footnotes (M3)
                link:     {"op":"add","path":"/body/p[3]","type":"link","props":{"text":"site","url":"https://…"}}
                          internal jump: {"props":{"text":"see intro","anchor":"Intro"}} (a bookmark name)
                bookmark: {"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Intro"}} (letter/_ start, ≤40 chars)
                footnote: {"op":"add","path":"/body/p[3]","type":"footnote","props":{"text":"the note"}} — reference lands at the paragraph end.
                """,
                ["docx/paragraph", "edit-ops"]),

            ["docx/section"] = (
                """
                ## docx page setup (M3) — set on /section[1]
                pageSize     A4 | Letter | Legal | A3 | A5
                orientation  portrait | landscape (width/height swap automatically)
                marginTop / marginBottom / marginLeft / marginRight   unit-qualified, e.g. "2cm"
                {"op":"set","path":"/section[1]","props":{"pageSize":"A4","orientation":"landscape"}}
                office_get /section[1] reflects everything (cm). Inserting NEW sections is not supported yet.
                """,
                ["docx/paragraph", "addressing"]),

            ["xlsx/sheet"] = (
                """
                ## xlsx sheet-level features (M3)
                freeze:     {"op":"set","path":"/Sheet1","props":{"freezeRows":"1","freezeCols":"2"}} (0 clears an axis)
                autoFilter: {"op":"set","path":"/Sheet1/A1:D20","props":{"autoFilter":"true"}} (one per sheet; false clears)
                print:      {"op":"set","path":"/Sheet1","props":{"orientation":"landscape","paperSize":"A4","fitToWidth":"1","printArea":"A1:F40"}}
                names:      {"op":"add","path":"/Sheet1/B2:C5","type":"name","props":{"name":"SalesData","scope"?:"workbook|sheet"}}
                            then formulas just use it: {"op":"set","path":"/Sheet1/D6","props":{"value":"=SUM(SalesData)"}} (evaluates).
                            get /name[@name=X] or /Sheet1/name[@name=X].
                Streaming reads: files over 20 MB (or args.stream:true) answer read stats/text and cell/range get
                via a SAX scan (no DOM). Reads only — edits still load the whole workbook.
                """,
                ["xlsx/cell", "edit-ops"]),

            ["pptx/chart"] = (
                """
                ## pptx charts (M3) — add type "chart" on /slide[i]
                kind        bar | line | pie (pie: exactly one series)
                categories  ["Q1","Q2",…]            series  [{"name":"Sales","values":[10,20]}] (null = gap)
                title, x, y, w, h   optional
                Data is cached literally (c:strLit/c:numLit, no embedded workbook): projections report dataEditable:false.
                Cross-doc dataFrom: {"props":{"kind":"bar","dataFrom":"book.xlsx!Sheet1/A1:B5"}} replaces categories/series —
                first column = categories, header row = series names, remaining columns = series values; the workbook is
                sandbox-resolved; a wrong range fails typed with candidates. Quote sheet names: book.xlsx!'Q3 Data'/A1:C5.
                Address /slide[i]/chart[k] for get (full data readback), set (title) and remove.
                """,
                ["pptx/slide", "edit-ops"]),

            ["pptx/transition"] = (
                """
                ## pptx slide transitions (M3) — slide-level set
                {"op":"set","path":"/slide[2]","props":{"transition":"fade|push|wipe|none","transitionDuration":"0.5s"}}
                Duration maps to p14:dur milliseconds (PowerPoint 2010+; older clients use the spd fallback).
                office_get /slide[2] and read {view:"outline"} report transition/transitionDuration.
                """,
                ["pptx/slide", "edit-ops"]),
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
