using System.Text.Json;
using ModelContextProtocol.Protocol;

namespace AIOffice.Mcp;

/// <summary>
/// The 19 MCP tools (17 verb-tools + preview_mark/preview_goto), mirroring the CLI verbs (docs/MCP.md is the spec).
/// M9 added <c>office_convert</c> (the 17th tool); M8 added <c>office_diff</c>;
/// M7 added <c>office_audit</c>; M1 added <c>preview_open</c> / <c>preview_selection</c>.
/// <para>
/// <c>preview_open</c> / <c>preview_selection</c> were reserved through v0 and
/// register here in M1, inside the token budget set aside for them in
/// docs/MCP.md §2.1.
/// </para>
/// <para>
/// Budget discipline: property tables, selector grammar and addressing details live in
/// <c>office_help</c>, never in these schemas. The schema-budget test asserts the whole
/// serialized surface stays ≤ 3500 tokens (chars/4).
/// </para>
/// </summary>
public static class ToolCatalog
{
    /// <summary>All tools, in spec order (17 as of M9).</summary>
    public static IReadOnlyList<Tool> Tools { get; } =
    [
        Make(
            "office_create",
            "Create a blank .docx/.xlsx/.pptx inside the workspace (kind inferred from the extension), or import one via 'from' (markdown→docx, csv→xlsx).",
            """
            {"type":"object","properties":{
              "file":{"type":"string","description":"Target path ending in .docx/.xlsx/.pptx, inside workspace"},
              "from":{"type":"string","description":"Import source: .md/.markdown -> .docx (headings/lists/tables/links/code), .csv/.tsv -> .xlsx (typed cells; leading-zero codes stay text). Mismatched pairs fail with the valid matrix"},
              "kind":{"type":"string","enum":["docx","xlsx","pptx"],"description":"Override when extension is non-standard"},
              "title":{"type":"string","description":"Document title written to core properties"},
              "overwrite":{"type":"boolean","default":false}},
             "required":["file"]}
            """),
        Make(
            "office_read",
            "Read a document: outline, plain text, stats or full structure tree. First step to understand a file.",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},
              "view":{"type":"string","enum":["outline","text","stats","structure","properties","embeds","revisions","comments","styles","fields","markdown","csv"],"default":"outline",
                "description":"outline: headings/slides/sheets skeleton with paths; text: plain text; stats: counters; structure: full element tree with paths+types; properties: core + custom doc properties under data.properties.{core,custom} (all formats); embeds: embedded OLE/package objects (all formats); comments: comment threads (all formats); revisions/markdown/fields: docx (fields = content controls); styles: docx style defs or xlsx named cell styles; csv: one xlsx sheet as RFC-4180 csv"},
              "range":{"type":"string","description":"Scope limit 'a..b' (1-based): paragraphs for docx, slides for pptx, rows for xlsx; csv view also takes 'A1:C10'"},
              "sheet":{"type":"string","description":"csv view: sheet name (default: first sheet)"},
              "max_bytes":{"type":"integer","description":"Cap payload size; truncation reported in meta.warnings and data.truncated"}},
             "required":["file"]}
            """),
        Make(
            "office_query",
            "Find elements with a CSS-like selector; returns canonical paths for office_get/office_edit.",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},
              "selector":{"type":"string","description":"CSS-like selector, e.g. \"p[style=Heading1]\", \"cell[value>100]\", \"shape:contains('Q3')\". Full grammar: office_help {topic:\"selectors\"}"}},
             "required":["file","selector"]}
            """),
        Make(
            "office_get",
            "Read one node and its properties by canonical path (inspect before editing).",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},
              "path":{"type":"string","description":"Canonical 1-based path, e.g. \"/body/p[3]\", \"/body/p[3]/omath[1]\", \"/Sheet1/B2\", \"/Sheet1/table[@name=Sales]\", \"/slide[2]/shape[3]\", \"/section[1]\", \"/master[1]/layout[2]\"; \"/\" = pptx slide-size/section root. Grammar: office_help {topic:\"addressing\"}"}},
             "required":["file","path"]}
            """),
        Make(
            "office_edit",
            "ALL mutations: apply an atomic batch of set/add/remove/move/replace/accept/reject/extract ops — all-or-nothing single save, auto-snapshot, optional rev guard.",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},
              "ops":{"type":"array","minItems":1,
                "description":"Atomic batch, applied in order. Use office_help {topic:\"<fmt>/<element>\"} for exact prop names and element types — do NOT guess.",
                "items":{"type":"object","properties":{
                  "op":{"type":"string","enum":["set","add","remove","move","replace","accept","reject","extract"],
                    "description":"accept/reject resolve docx tracked revisions (path: /revision[@id=N] or a scope like /body). replace = find/replace in scope: props {find,replace,regex?,matchCase?,wholeWord?}; path \"/\" = whole document (docx body+headers+footers, every sheet, every slide+notes); 0 matches -> ok + find_no_match warning. extract writes an embedded object's bytes to a sandbox destination (props.to) — a producing op that does NOT modify the source"},
                  "path":{"type":"string","description":"set/remove/move: target element. add: PARENT element, e.g. \"/body\", \"/slide[2]\", \"/Sheet1\". replace: container scope or \"/\""},
                  "type":{"type":"string","description":"add only: element type, e.g. paragraph, run, table (docx/pptx/xlsx ListObject), row, col, cell, slide, shape, image, comment, reply, note, style, header, footer, chart, pivot, conditionalFormat, toc, watermark, footnote, endnote, sectionBreak, equation (docx LaTeX), columnBreak, animation, section (pptx), layout (pptx clone), group (xlsx outline), field, dataValidation, sparkline, caption (docx), crossRef (docx), slicer (xlsx), source/citation/bibliography (docx), media (pptx audio/video), smartart/connector/group/ungroup (pptx), tableOfFigures/indexEntry/index/mergeField (docx), formControl (xlsx), dataTable (xlsx what-if), ifField (docx «IF» merge field), zoom (pptx slide/section/summary nav), scenario (xlsx what-if changing-cell set), buildingBlock/buildingBlockRef (docx reusable AutoText), font (pptx embedded .ttf), actionButton (pptx), linkedPicture (xlsx camera-tool snapshot of a range). A dynamic-array formula on a cell (=FILTER/UNIQUE/SORT/SORTBY/SEQUENCE/RANDARRAY/TRANSPOSE) spills automatically; scalar XLOOKUP/IFS/SWITCH/LET/TEXTJOIN now evaluate to cached values. set props gain goalSeek+applyScenario (xlsx), formula on table cells + lineNumbers + dropCap (docx), embedAll (pptx font); v1.7 also: xlsx print titles/pageBreak/fitToPage/printHeader/printFooter/center + calculationMode/iterativeCalc; docx watermark image + STYLEREF/SYMBOL/QUOTE fields + equation number:true; pptx /notesMaster + /handoutMaster + animation repeat/rewind/autoReverse + table-cell valign/margins. office_help {topic:\"formulas\"}"},
                  "props":{"type":"object","additionalProperties":{"type":"string"},
                    "description":"String-valued props, e.g. {\"text\":\"Hi\",\"bold\":\"true\",\"size\":\"12pt\",\"fill\":\"FF0000\"}. Sizes unit-qualified; colors hex/named. Table cells merge via {\"mergeRight\":\"2\"}/{\"mergeDown\":\"2\"}. pptx add chart: {\"dataFrom\":\"book.xlsx!Sheet1/A1:B5\"} pulls categories+series from a workbook (first col = categories, header row = series names) instead of literals"},
                  "position":{"type":["integer","string"],
                    "description":"add/move: 1-based index within parent, or \"before:<path>\" / \"after:<path>\"; omit = append"}},
                 "required":["op","path"]}},
              "track":{"type":"boolean","default":false,"description":"docx: record text set/add/remove ops as tracked revisions (w:ins/w:del); resolve later with accept/reject"},
              "author":{"type":"string","description":"Author stamped on revisions/comments (op props.author wins; default: AIOFFICE_AUTHOR env, then \"AIOffice\")"},
              "expect_rev":{"type":"string","description":"Optimistic lock: 12-hex rev from a previous meta.rev; mismatch fails with stale_address BEFORE any write"},
              "dry_run":{"type":"boolean","default":false,"description":"Validate the whole batch without writing"}},
             "required":["file","ops"]}
            """),
        Make(
            "office_render",
            "Render the document (or a subtree) to an inspectable artifact — the look step of render→look→fix. png also returns the image inline.",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},
              "to":{"type":"string","enum":["html","svg","text","png","pdf"],"default":"html",
                "description":"html: docx/xlsx/pptx; svg: pptx, one file per slide; text: plain text; png: browser screenshot, written next to source (pptx: one slide, default /slide[1] — pass scope); pdf: paged print via local Chromium, written next to source (pptx: whole deck, one page per slide)"},
              "scope":{"type":"string","description":"Render only this subtree, e.g. \"/slide[3]\", \"/Sheet1/A1:F20\", \"/body/table[1]\""},
              "engine":{"type":"string","enum":["chromium","soffice","auto"],"default":"chromium",
                "description":"png/pdf engine. soffice=LibreOffice true-fidelity (png also needs pdftoppm); auto=soffice if installed else chromium. office_help{topic:\"render-engines\"}"},
              "output":{"type":"string","description":"Output file or directory inside workspace (default: alongside source)"}},
             "required":["file"]}
            """),
        Make(
            "office_validate",
            "OpenXmlValidator schema check + lint; each issue carries a suggestion. Cheap post-edit health check.",
            """
            {"type":"object","properties":{"file":{"type":"string"}},"required":["file"]}
            """),
        Make(
            "office_audit",
            "Accessibility + quality lint: findings are DATA (ok:true even with errors). fix:true applies only safe autofixes (alt text, table header, doc/slide title, orphan bookmark). Codes: office_help {topic:\"audit\"}.",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},
              "category":{"type":"string","enum":["accessibility","quality","all"],"default":"all"},
              "severity":{"type":"string","enum":["error","warning","info"],"default":"info","description":"Minimum severity to report"},
              "fix":{"type":"boolean","default":false,"description":"Apply safe autofixes, then re-audit; result adds {fixed:N, remaining:[ids]}"}},
             "required":["file"]}
            """),
        Make(
            "office_diff",
            "Semantically compare a document against a baseline — another same-format file (other) OR one of its own snapshots (snapshot:N). Returns a sorted added/removed/modified/moved change list. Changes are data: ok:true.",
            """
            {"type":"object","properties":{
              "file":{"type":"string","description":"The current/new document"},
              "other":{"type":"string","description":"Baseline: the OLD same-format document. A format mismatch is invalid_args. Exactly one of other/snapshot required"},
              "snapshot":{"type":"integer","description":"Baseline: snapshot N of file from its auto pre-edit ring; a missing index is invalid_args naming the available numbers. Exactly one of other/snapshot required"},
              "view":{"type":"string","enum":["summary","detailed"],"default":"detailed","description":"detailed: full before/after per change; summary: counts + path+kind only"}},
             "required":["file"]}
            """),
        Make(
            "office_template",
            "Fill {{key}} placeholders (and docx MERGEFIELD / «IF» fields) in text runs from a merge map. An OBJECT data fills one document; an ARRAY of records runs a mail merge (v1.4). Placeholders split across runs are matched.",
            """
            {"type":"object","properties":{
              "file":{"type":"string","description":"Template document containing {{key}} placeholders (docx also fills MERGEFIELD / «IF» fields by name)"},
              "data":{"type":["object","array"],
                "description":"OBJECT {\"client\":\"ACME\"} fills ONE document; ARRAY [{...},{...}] of record objects runs a mail merge — one merge per record (v1.4)"},
              "output":{"type":"string","description":"Single fill: result path (recommended; omit = in place, snapshotted). Mail merge (array data): a path PATTERN — {n}=1-based record index, {Field}=that record's value, e.g. \"letter-{n}.docx\". Every expanded path is sandbox-resolved. Omit for one combined doc (a next-page section per record). office_help {topic:\"mail-merge\"}"},
              "overwrite":{"type":"boolean","default":false}},
             "required":["file","data"]}
            """),
        Make(
            "office_convert",
            "Convert a document between formats: docx/xlsx/pptx <-> each other (content-neutral model), docx<->md, xlsx<->csv, any->pdf/png/svg/html. Lossy across formats; dropped content is named in data.dropped + a convert_lossy warning. dest is created fresh.",
            """
            {"type":"object","properties":{
              "src":{"type":"string","description":"Source document (.docx/.xlsx/.pptx/.md/.csv); opened read-only"},
              "dest":{"type":"string","description":"Destination, created fresh: .docx/.xlsx/.pptx/.md/.csv (content) or .pdf/.png/.svg/.html/.txt (render). Same ext as src -> invalid_args (use office_edit); unknown ext -> unsupported_feature"}},
             "required":["src","dest"]}
            """),
        Make(
            "file_snapshot",
            "List or restore the automatic pre-edit snapshots (ring of 20 per file).",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},
              "action":{"type":"string","enum":["list","restore"],"default":"list"},
              "n":{"type":"integer","description":"restore: snapshot number from list (omit = latest)"}},
             "required":["file"]}
            """),
        Make(
            "office_status",
            "Health check: runtime, workspace sandbox, snapshot store. Call when tools start failing oddly.",
            """
            {"type":"object","properties":{}}
            """),
        Make(
            "office_help",
            "Progressive docs: addressing grammar, selector syntax, per-element property tables. Look up — never guess.",
            """
            {"type":"object","properties":{
              "topic":{"type":"string","description":"Omit -> topic index. Examples: \"addressing\", \"selectors\", \"errors\", \"envelope\", \"edit-ops\", \"docx/paragraph\", \"xlsx/cell\", \"pptx/shape\""}}}
            """),
        Make(
            "office_schema",
            "Machine-readable JSON of the whole command surface: every verb, its args, error codes and examples.",
            """
            {"type":"object","properties":{
              "verb":{"type":"string","description":"Omit -> full surface. One of: create|read|query|get|edit|render|validate|audit|diff|template|convert|snapshot|doctor|schema|help|preview|mcp"}}}
            """),
        Make(
            "preview_open",
            "Live localhost preview (click=select, dblclick=edit, auto-reloads on change). Returns url; survives this session.",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},
              "port":{"type":"integer","description":"Fixed port (default: auto in 26500-26600)"}},
             "required":["file"]}
            """),
        Make(
            "preview_selection",
            "The canonical paths the human clicked; feed to office_get/office_edit.",
            """
            {"type":"object","properties":{"file":{"type":"string"}},"required":["file"]}
            """),
        Make(
            "preview_mark",
            "Highlight a preview element (advisory; no doc edit). path may be 'selected'.",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},"path":{"type":"string"},"note":{"type":"string"},"toFix":{"type":"boolean"}},
             "required":["file","path"]}
            """),
        Make(
            "preview_goto",
            "Scroll every preview viewer to an element and flash it.",
            """
            {"type":"object","properties":{"file":{"type":"string"},"path":{"type":"string"}},"required":["file","path"]}
            """),
    ];

    /// <summary>Tool names in spec order (for candidates lists and tests).</summary>
    public static IReadOnlyList<string> Names { get; } = [.. Tools.Select(t => t.Name)];

    private static Tool Make(string name, string description, string schemaJson) => new()
    {
        Name = name,
        Description = description,
        InputSchema = JsonSerializer.Deserialize<JsonElement>(schemaJson),
    };
}
