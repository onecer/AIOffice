# AIOffice AI-Facing Contract (v1.0)

This document is the **stability promise** for the AIOffice surface that AI agents
and scripts depend on. Everything below is **frozen for v1.0**: the JSON envelope
shape, the error-code set, the addressing grammar, the exit codes, the op / view
vocabularies, and the MCP tool list. Within the `1.0` line these will only be
**added to**, never removed or renamed. A breaking change bumps `surfaceVersion`.

The contract version is reported as `surfaceVersion` (currently **`1.0`**) in
`office_schema` output and in `doctor` / `office_status` capabilities. It is
independent of the tool's package version (`meta.version`): the tool ships many
releases while the contract holds steady.

> 中文摘要：本文件是 AIOffice 面向 AI 的**稳定性契约**（v1.0）。以下内容在 1.0
> 版本线中**冻结**：JSON 信封结构、错误码集合、寻址语法、退出码、操作（op）与视图
> （view）词表、以及 MCP 工具清单。1.0 线内只增不删、不改名；任何破坏性变更都会提升
> `surfaceVersion`。契约版本通过 `office_schema` 与 `doctor` 能力中的 `surfaceVersion`
> 字段（当前为 `1.0`）声明，与工具包版本（`meta.version`）相互独立。

## Stability promise

These are stable for the entire **1.0 series**. Concretely:

- The **frozen lists** in §§2–7 (error codes, exit codes, addressing grammar, op
  kinds, view names, MCP tool names) are **additive-only** in 1.0: new entries may
  appear, but nothing is **removed or renamed** without bumping `surfaceVersion`.
- The **envelope shape** (§1) — `{ok, data, error, meta}`, the `error{code,
  message, suggestion, candidates?}` shape, and the camelCase, null-omitted wire
  form — does not change in 1.0.
- Existing **result fields keep their meaning**; new fields may be **added** to a
  result (parsers must ignore unknown keys). The per-command result *payloads* are
  not a frozen list — they stabilize at 1.0 but may gain fields. Known per-format
  payload differences that 1.0 does **not** unify are listed in §9.

---

## 1. The JSON envelope

Every command prints **exactly one** JSON object to stdout:

```json
{
  "ok": true,
  "data": { "...": "command result, or null on error" },
  "error": null,
  "meta": {
    "file": "report.docx",
    "rev": "a3f9c12be01d",
    "elapsedMs": 12,
    "version": "1.0.0",
    "warnings": [ { "code": "equation_partial", "message": "..." } ]
  }
}
```

- **Success**: `ok:true`, `data` present, `error` absent.
- **Failure**: `ok:false`, `error` present, `data` absent.
  (Null fields are **omitted** from the compact wire form, so on success the
  `error` key is absent rather than present-and-null, and vice versa.)
- `error` is always `{code, message, suggestion, candidates?}`. **`suggestion` is
  never empty** — every error names a way forward. `candidates` lists nearest-match
  paths and is **required** when an `invalid_path` resolution fails.
- `meta.warnings` (when present) is a list of `{code, message}` — non-fatal notes.
  A warning **never** changes the exit code; the command still succeeds (exit 0).
  The full warning vocabulary (FROZEN, additive-only) is:

  | code | meaning |
  |------|---------|
  | `convert_lossy` | a cross-format convert dropped content; see `data.dropped` |
  | `equation_partial` | a LaTeX construct was only partially rendered |
  | `find_no_match` | a replace matched nothing; the edit still succeeded |
  | `formula_not_evaluated` | a formula's cached value was returned (no headless recalc) |
  | `result_truncated` | a payload exceeded its byte cap; narrow scope or raise `--max-bytes` |
  | `diff_truncated` | a diff change list hit its cap and was trimmed |
  | `already_running` | a live preview server was already up; the existing one was reused |
  | `bibliography_cached` | (1.1) a docx bibliography's entries came from cached values; Word rebuilds on field refresh |
  | `caption_numbers_cached` | caption/cross-ref numbers came from cached SEQ values |
  | `csv_empty` | a csv import/export produced no rows |
  | `md_block_skipped` · `md_html_skipped` · `md_image_skipped` · `md_link_skipped` | a markdown source had content with no neutral equivalent |
  | `scope_defaulted` | a render scope was defaulted (e.g. a deck rendered slide 1) |
  | `stream_fallback` | a streaming view fell back to loading the whole workbook |
  | `template_unresolved` | one or more `{{placeholders}}` were left unfilled |
  | `toc_pages_unknown` · `toc_refreshed` | a table of contents could not be paginated / was refreshed |

  `office_schema` / `aioffice schema` lists the live set under `data.warningCodes`.
- `meta.rev` is a 12-hex content revision usable with `--expect-rev` for optimistic
  locking. `meta.version` is the tool package version.

Field names are **camelCase**. Null fields are omitted from the compact wire form.

## 2. The error codes (FROZEN)

The complete, closed set. Codes are stable snake_case strings:

| code | meaning |
|------|---------|
| `invalid_args` | malformed/missing arguments |
| `file_not_found` | the target file does not exist |
| `sandbox_denied` | a path escaped the workspace sandbox |
| `invalid_path` | an address did not resolve (carries `candidates`) |
| `unsupported_feature` | capability not implemented; `suggestion` names the workaround |
| `format_corrupt` | the file could not be opened as its declared kind |
| `stale_address` | `--expect-rev` mismatch; nothing was written |
| `file_too_large` | exceeds the size guard (`AIOFFICE_MAX_FILE_MB`) |
| `formula_not_evaluated` | warning-level: a formula's cached value was returned |
| `find_no_match` | warning-level: a replace matched nothing; the edit still succeeded |
| `internal_error` | an unexpected failure |
| `preview_not_running` | no live preview server for the file |

No code in this table will be removed or renamed in the 1.0 line.

## 3. Exit codes

| exit | when |
|------|------|
| `0` | success (`ok:true`) |
| `2` | user error (`invalid_args`, `file_not_found`, `invalid_path`, `stale_address`, `file_too_large`, `preview_not_running`) |
| `3` | `format_corrupt` or `internal_error` |
| `4` | `sandbox_denied` |
| `5` | `unsupported_feature` |

Errors never crash the process — they become a typed envelope plus the exit code above.

## 4. Addressing grammar

Paths are **1-based**; the `@id` / `@name` forms are stable across insertions and are
what `query` / `get` always return. `office_help {topic:"addressing"}` is the full grammar.

**docx** — `/body`, `/body/p[3]`, `/body/p[3]/run[2]`, `/body/p[3]/omath[1]` (equation),
`/body/table[1]/tr[2]/tc[3]`, `/section[1]`, `/embed[1]` (embedded object), `/properties`.

**xlsx** — `/Sheet1`, `/Sheet1/B2`, `/Sheet1/A1:C10`, `/Sheet1/table[@name=Sales]`,
`/Sheet1/image[1]`, `/Sheet1/embed[1]` (embedded object), `/Pivot/pivot[1]`, `/properties`.

**pptx** — `/` (presentation: slide size + sections), `/slide[2]`, `/slide[2]/notes`,
`/slide[2]/shape[@id=7]`, `/slide[2]/shape[@id=7]/p[1]`, `/slide[2]/shape[@id=7]/omath[1]`
(equation), `/slide[2]/chart[1]`, `/slide[2]/table[1]/tr[2]/tc[3]`, `/slide[2]/embed[@id=7]`
(embedded object), `/master[1]/layout[2]`, `/section[1]`, `/properties`.

## 5. The op-type vocabulary

`office_edit` op kinds (the `op` field):

`set` · `add` · `remove` · `move` · `replace` · `accept` · `reject` · `extract`

- `accept` / `reject` resolve docx tracked revisions.
- `extract` writes an embedded object's bytes to a sandbox destination — it is a
  **producing** op and does **not** modify the source document.

`add` element **types** (the `type` field) are format-specific; notable cross-format ones:
`embed` (any file as an OLE/package object, all formats) and `equation` (LaTeX → OMML
math; **docx + pptx**; xlsx returns `unsupported_feature` — Excel has cell formulas,
not math objects). Use `office_help {topic:"<fmt>/<element>"}` for exact prop names.

**1.1 added** `add` types (additive within the 1.0 contract line): `source`,
`citation`, `bibliography` (**docx** — bibliography sources, CITATION fields, and a
rendered bibliography; see `office_help {topic:"docx/citation"}`) and `media`
(**pptx** — embedded audio/video; see `office_help {topic:"pptx/media"}`).

## 6. The view vocabulary

`office_read --view` (the `view` field):

`outline` · `text` · `stats` · `structure` · `properties` · `embeds` · `comments`
— available on **all three** formats.

- `properties` returns core + custom document properties under the **unified shape**
  `data.properties.{core, custom}` — identical across docx/xlsx/pptx.
- `embeds` lists embedded OLE/package objects. Each entry is at minimum
  `{path, name, mediaType, size}`; docx/pptx add `container`, and **xlsx names its
  container field `sheet`** (the worksheet the object is anchored on) and adds a
  `source` field. Treat `container` / `sheet` as the per-format "where it lives" key.
- `comments` lists comment threads on every format (`{view:"comments", …}`).
- Every read view echoes the view name as `data.view` (xlsx also carries
  `data.kind = "xlsx"` for historical reasons; prefer `data.view`).

Format-specific views: docx adds `revisions`, `styles`, `fields`, `markdown`, and
(**1.1**) `sources` (the bibliography source store); xlsx adds `csv`, `styles`.

Applying a view a format does not support is `invalid_args` (exit 2): the
`error.suggestion` lists that format's valid views.

## 7. The MCP tool list (17 tools)

`office_create` · `office_read` · `office_query` · `office_get` · `office_edit` ·
`office_render` · `office_validate` · `office_audit` · `office_diff` ·
`office_template` · `office_convert` · `file_snapshot` · `office_status` ·
`office_help` · `office_schema` · `preview_open` · `preview_selection`

This is the exact, complete set `tools/list` returns. The tools mirror the CLI
verbs 1:1, with two naming details to note rather than infer:

- The **snapshot** verb's tool is **`file_snapshot`** (not `office_snapshot`) — it
  operates on the file's snapshot ring, not document content. Call it by this exact
  name; do not derive `office_<verb>` for snapshot.
- The `doctor` verb's tool is **`office_status`**, and the `preview` verb maps to
  **two** tools (`preview_open`, `preview_selection`). The CLI-only `version` verb
  and the `mcp` server-launch verb have no tools.

The embedded-object and equation surfaces ride on `office_edit` / `office_read` /
`office_get` — they **do not** add new tools; the count stays **17**.

## 7a. 1.1.0 additions (additive within the frozen 1.0 line)

Package **1.1.0** stays on **`surfaceVersion 1.0`** — every change below is purely
**additive** (new enum/prop values, one new view, one new warning), so nothing in
§§1–7 was removed or renamed. The 18 CLI verbs / 17 MCP tools are unchanged.

- **New `add` types** (§5): `source`, `citation`, `bibliography` (docx) and `media`
  (pptx).
- **New `read --view`** (§6): `sources` (docx bibliography store).
- **New `conditionalFormat` kind** (xlsx, prop-value enum): `iconSet`
  (3/4/5-icon glyph sets, e.g. `3TrafficLights1`), alongside the existing
  `cellIs` · `colorScale` · `dataBar` · `containsText`.
- **New chart `kind` values** (xlsx **and** pptx, prop-value enums): `doughnut`,
  `radar`, `bubble`, `stackedBar`, `percentStackedBar`, `stackedArea`, `combo` —
  added to the existing `bar` · `line` · `pie` · `scatter` · `area`. An unsupported
  kind still returns `unsupported_feature` listing the now-expanded supported set.
- **New text/shape effect props** (docx run-level, pptx shape-level set):
  `shadow`, `glow`, `reflection`, `outline`.
- **New pptx transition kinds**: `split`, `reveal`, `cut`, `zoom`, added to the
  existing `none` · `fade` · `push` · `wipe`.
- **New warning** (§1): `bibliography_cached`.

## 8. What is experimental (NOT frozen)

These are explicitly outside the frozen contract and may change within the 1.0 line:

- The **PNG/PDF render** path (Chromium screenshotting) and the `preview_*`
  live-preview protocol details.
- The exact **shape of the `structure` view** payload (the tree is stable in spirit,
  but new fields may appear and nesting may be refined).
- **Convert** drop-reporting detail (`data.dropped` entries) — the loss is reported,
  but the granularity may grow.
- Internal vendor attributes (e.g. how the original LaTeX is stored on an `m:oMath`,
  or the embed registry storage) — read them back through `get`, never by reaching
  into the XML.

## 9. Known per-format result differences (stable but not uniform)

These are intentional, stable quirks an agent learning the shape on one format
should know about. They are **not** bugs and 1.0 does not unify them (doing so would
be a breaking rename); a robust parser keys on what is present:

- **edit** result: the per-op array key is `ops` on docx/xlsx but `results` on pptx.
  All three carry `applied` (count) and `snapshot` (the pre-image number, for
  `snapshot restore`); a producing-only batch (all `extract` ops) writes nothing and
  omits `snapshot`.
- **query** result: docx/pptx carry `count` (= returned length); xlsx carries **both**
  `count` and `total` (the full pre-truncation match count, ≥ `count`).
- **validate** result: all three carry `valid` and `count` (issue count); xlsx adds
  `errors` / `warnings` sub-counts.
- **read views** echo `data.view`; xlsx additionally carries `data.kind = "xlsx"`.
- **embeds**: the container key is `container` on docx/pptx, `sheet` on xlsx (see §6).

## 10. Known limitations (what AIOffice does NOT do)

Honest scope boundaries for 1.0:

- **No xlsx equations.** Excel has cell formulas, not OMML math objects;
  `add type:equation` on xlsx returns `unsupported_feature`. (docx + pptx do support
  equations.)
- **No headless formula recalculation.** Formulas are written with a recalc flag so
  the host app recomputes on open; reads return the **cached** value with a
  `formula_not_evaluated` warning. AIOffice is not a spreadsheet engine.
- **Convert is lossy content-transfer**, not fidelity-preserving. Cross-format
  conversion moves *content* through a neutral model; charts, pivots, images,
  animations, transitions and embedded objects are named in `data.dropped` +
  `convert_lossy`, not carried.
- **Render is an HTML/SVG approximation**, not pixel-perfect Office layout. PNG/PDF
  are produced by screenshotting that HTML via a local Chromium; exact pagination,
  font metrics and effects will differ from Microsoft Office.
- **No real-time collaboration / co-authoring / live cloud sync.** AIOffice operates
  on local files in a sandbox; optimistic locking (`--expect-rev`) and the snapshot
  ring are the concurrency story, not multi-user editing.
- **No macros / VBA / scripting execution.** AIOffice neither runs nor authors macros.
- **xlsx audit edit-op suggestions use a compact shorthand** (`{op:set, path:…,
  props:{…}}` with unquoted keys) rather than strictly-quoted JSON; docx/pptx
  suggestions are valid JSON. Treat the xlsx form as a hint, not a paste target.
- **Single-maintainer, AI-built.** AIOffice is 100% self-built (C#/.NET, OpenXML SDK +
  ClosedXML) by one maintainer with AI assistance. It is not a Microsoft product and
  carries no enterprise support or certification.

Everything in §§1–7 is the stable v1.0 contract; §§8–10 describe what is flexible,
non-uniform, or out of scope.
