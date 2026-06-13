# AIOffice AI-Facing Contract (v1.0-rc)

This document is the **stability promise** for the AIOffice surface that AI agents
and scripts depend on. Everything below is **frozen for v1.0**: the JSON envelope
shape, the error-code set, the addressing grammar, the exit codes, the op / view
vocabularies, and the MCP tool list. Within the `1.0` line these will only be
**added to**, never removed or renamed. A breaking change bumps `surfaceVersion`.

The contract version is reported as `surfaceVersion` (currently **`1.0-rc`**) in
`office_schema` output and in `doctor` / `office_status` capabilities. It is
independent of the tool's package version (`meta.version`): the tool ships many
releases while the contract holds steady.

> 中文摘要：本文件是 AIOffice 面向 AI 的**稳定性契约**（v1.0-rc）。以下内容在 1.0
> 版本线中**冻结**：JSON 信封结构、错误码集合、寻址语法、退出码、操作（op）与视图
> （view）词表、以及 MCP 工具清单。1.0 线内只增不删、不改名；任何破坏性变更都会提升
> `surfaceVersion`。契约版本通过 `office_schema` 与 `doctor` 能力中的 `surfaceVersion`
> 字段（当前为 `1.0-rc`）声明，与工具包版本（`meta.version`）相互独立。

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
    "version": "0.11.0",
    "warnings": [ { "code": "equation_partial", "message": "..." } ]
  }
}
```

- **Success**: `ok:true`, `data` present, `error:null`.
- **Failure**: `ok:false`, `data:null`, `error` present.
- `error` is always `{code, message, suggestion, candidates?}`. **`suggestion` is
  never empty** — every error names a way forward. `candidates` lists nearest-match
  paths and is **required** when an `invalid_path` resolution fails.
- `meta.warnings` (when present) is a list of `{code, message}` — non-fatal notes
  (e.g. `equation_partial`, `find_no_match`, `formula_not_evaluated`).
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

## 6. The view vocabulary

`office_read --view` (the `view` field):

`outline` · `text` · `stats` · `structure` · `properties` · `embeds`
— available on **all three** formats.

- `properties` returns core + custom document properties under the **unified shape**
  `data.properties.{core, custom}` — identical across docx/xlsx/pptx.
- `embeds` lists embedded OLE/package objects: `{path, name, mediaType, size, container}`.

Format-specific views: docx adds `revisions`, `comments`, `styles`, `fields`,
`markdown`; xlsx adds `csv`, `styles`.

## 7. The MCP tool list (17 tools)

`office_create` · `office_read` · `office_query` · `office_get` · `office_edit` ·
`office_render` · `office_validate` · `office_audit` · `office_diff` ·
`office_template` · `office_convert` · `office_status` · `office_help` ·
`office_schema` · `preview_open` · `preview_selection`

The tools mirror the CLI verbs 1:1. The embedded-object and equation surfaces ride on
`office_edit` / `office_read` / `office_get` — they **do not** add new tools; the count
stays **17**.

## 8. What is experimental (NOT frozen)

These may still change before 1.0 final and are explicitly outside the frozen contract:

- The **PNG render** path (Chromium screenshotting) and the `preview_*` live-preview
  protocol details.
- The exact **shape of the `structure` view** payload (the tree is stable in spirit,
  but new fields may appear and nesting may be refined).
- **Convert** drop-reporting detail (`data.dropped` entries) — the loss is reported,
  but the granularity may grow.
- Internal vendor attributes (e.g. how the original LaTeX is stored on an `m:oMath`,
  or the embed registry storage) — read them back through `get`, never by reaching
  into the XML.

Everything in §§1–7 is the stable v1.0 contract.
