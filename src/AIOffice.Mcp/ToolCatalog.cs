using System.Text.Json;
using ModelContextProtocol.Protocol;

namespace AIOffice.Mcp;

/// <summary>
/// The 14 M1 MCP tools, mirroring the CLI verbs 1:1 (docs/MCP.md is the spec).
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
    /// <summary>All v0 tools, in spec order.</summary>
    public static IReadOnlyList<Tool> Tools { get; } =
    [
        Make(
            "office_create",
            "Create a blank .docx/.xlsx/.pptx inside the workspace; kind inferred from the extension.",
            """
            {"type":"object","properties":{
              "file":{"type":"string","description":"Target path ending in .docx/.xlsx/.pptx, inside workspace"},
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
              "view":{"type":"string","enum":["outline","text","stats","structure","revisions","comments","styles"],"default":"outline",
                "description":"outline: headings/slides/sheets skeleton with paths; text: plain text; stats: counters; structure: full element tree with paths+types; revisions/comments/styles: docx only"},
              "range":{"type":"string","description":"Scope limit 'a..b' (1-based): paragraphs for docx, slides for pptx, rows for xlsx"},
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
              "path":{"type":"string","description":"Canonical 1-based path, e.g. \"/body/p[3]\", \"/Sheet1/B2\", \"/'Q3 Data'/A1:C10\", \"/slide[2]/shape[3]\"; \"/\" for document-level properties. Grammar: office_help {topic:\"addressing\"}"}},
             "required":["file","path"]}
            """),
        Make(
            "office_edit",
            "ALL mutations: apply an atomic batch of set/add/remove/move/accept/reject ops — all-or-nothing single save, auto-snapshot, optional rev guard.",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},
              "ops":{"type":"array","minItems":1,
                "description":"Atomic batch, applied in order. Use office_help {topic:\"<fmt>/<element>\"} for exact prop names and element types — do NOT guess.",
                "items":{"type":"object","properties":{
                  "op":{"type":"string","enum":["set","add","remove","move","replace","accept","reject"],
                    "description":"accept/reject resolve docx tracked revisions (path: /revision[@id=N] or a scope like /body). replace = find/replace in scope: props {find,replace,regex?,matchCase?,wholeWord?}; path \"/\" = whole document (docx body+headers+footers, every sheet, every slide+notes); 0 matches -> ok + find_no_match warning"},
                  "path":{"type":"string","description":"set/remove/move: target element. add: PARENT element, e.g. \"/body\", \"/slide[2]\", \"/Sheet1\". replace: container scope or \"/\""},
                  "type":{"type":"string","description":"add only: element type, e.g. paragraph, run, table, row, col, cell, slide, shape, image, comment, note, style, header, footer, chart, pivot, conditionalFormat, toc, watermark, footnote, endnote, sectionBreak, animation"},
                  "props":{"type":"object","additionalProperties":{"type":"string"},
                    "description":"String-valued props, e.g. {\"text\":\"Hi\",\"bold\":\"true\",\"size\":\"12pt\",\"fill\":\"FF0000\"}. Sizes unit-qualified; colors hex/named. pptx add chart: {\"dataFrom\":\"book.xlsx!Sheet1/A1:B5\"} pulls categories+series from a workbook (first col = categories, header row = series names) instead of literals"},
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
            "office_template",
            "Fill {{key}} placeholders in text runs from a merge map (docx/xlsx/pptx; placeholders split across runs are matched).",
            """
            {"type":"object","properties":{
              "file":{"type":"string","description":"Template document containing {{key}} placeholders in text runs"},
              "data":{"type":"object","additionalProperties":{"type":"string"},
                "description":"Merge map, e.g. {\"client\":\"ACME Corp\",\"date\":\"2026-06-12\"}"},
              "output":{"type":"string","description":"Result path (recommended). Omit = merge in place; pre-image auto-snapshotted"},
              "overwrite":{"type":"boolean","default":false}},
             "required":["file","data"]}
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
              "verb":{"type":"string","description":"Omit -> full surface. One of: create|read|query|get|edit|render|validate|template|snapshot|doctor|schema|help|preview|mcp"}}}
            """),
        Make(
            "preview_open",
            "Open a live localhost preview where a human clicks elements to select them; returns the url. Survives this session.",
            """
            {"type":"object","properties":{
              "file":{"type":"string"},
              "port":{"type":"integer","description":"Fixed port (default: auto-pick in 26500-26600)"}},
             "required":["file"]}
            """),
        Make(
            "preview_selection",
            "Read the canonical paths the human clicked in the live preview; they feed office_get/office_edit directly.",
            """
            {"type":"object","properties":{"file":{"type":"string"}},"required":["file"]}
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
