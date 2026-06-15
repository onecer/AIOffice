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
  | `figures_cached` | (1.2) a docx table of figures' entries came from cached captions; Word repaginates page numbers on open/refresh |
  | `index_cached` | (1.2) a docx index's entries were alphabetized from XE fields with cached page numbers; Word recomputes them on open/refresh |
  | `caption_numbers_cached` | caption/cross-ref numbers came from cached SEQ values |
  | `model3d_as_media` | (1.3) a pptx 3D model was embedded as a 3D part behind a poster picture fallback; PowerPoint 2019+ renders the model |
  | `linked_picture_static` | (1.7) an xlsx linked picture (camera tool) was embedded as a static snapshot of the source range, not a live link; re-add to refresh |
  | `equation_numbers_cached` | (1.7) a docx numbered display equation's number was written as a cached value; Word recomputes the sequence on open/refresh |
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
| `spill_blocked` | (1.4) a dynamic-array result would spill over non-empty cells; nothing was written |
| `find_no_match` | warning-level: a replace matched nothing; the edit still succeeded |
| `internal_error` | an unexpected failure |
| `preview_not_running` | no live preview server for the file |

No code in this table will be removed or renamed in the 1.0 line. (Codes added in a
minor release — e.g. `spill_blocked` in 1.4 — are additive; existing codes are stable.)

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
`/body/table[1]/tr[2]/tc[3]`, `/body/shape[1]` / `/body/textBox[1]` (1.3 body shapes),
`/formField[@name=status]` (1.3 form field), `/section[1]`, `/embed[1]` (embedded object),
`/theme` (1.3 theme color/font scheme), `/properties`.

**xlsx** — `/Sheet1`, `/Sheet1/B2`, `/Sheet1/A1:C10`, `/Sheet1/table[@name=Sales]`,
`/Sheet1/image[1]`, `/Sheet1/embed[1]` (embedded object), `/Pivot/pivot[1]`, `/properties`.

**pptx** — `/` (presentation: slide size + sections), `/slide[2]`, `/slide[2]/notes`,
`/slide[2]/shape[@id=7]`, `/slide[2]/shape[@id=7]/p[1]`, `/slide[2]/shape[@id=7]/omath[1]`
(equation), `/slide[2]/chart[1]`, `/slide[2]/model3d[@id=7]` (1.3 3D model),
`/slide[2]/table[1]/tr[2]/tc[3]`, `/slide[2]/embed[@id=7]`
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

**1.2 added** `add` types: `smartart` / `connector` / `group` / `ungroup` (**pptx**),
`tableOfFigures` / `indexEntry` / `index` / `mergeField` (**docx**), `formControl`
(**xlsx**). **1.3 added** `add` types: `shape` / `textBox` / `formField` (**docx**)
and `model3d` (**pptx**). All are `type` values, **not** new `op` kinds — §5's op
list is unchanged. Full additive lists: §§7a–7c.

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

## 7b. 1.2.0 additions (additive within the frozen 1.0 line)

Package **1.2.0** stays on **`surfaceVersion 1.0`** — every change below is purely
**additive** (new `add` type values, new prop keys, new warnings), so nothing in
§§1–7 was removed or renamed. The op **kinds** in §5 are unchanged
(`group`/`ungroup` are `add` **types**, not new op kinds). The 18 CLI verbs / 17 MCP
tools are unchanged.

- **New pptx `add` types** (§5): `smartart` (a list/process/hierarchy/orgChart/cycle
  diagram — **create + read**; editing nodes in place is `unsupported_feature`),
  `connector` (a straight/elbow/curved link between two shapes), `group` (wrap two or
  more shapes in a group) and `ungroup` (dissolve a group on its `/slide[i]/group[@id=N]`
  path). `group`/`ungroup` are `add` type values, **not** new `op` kinds. See
  `office_help {topic:"smartart"}` and `{topic:"connectors"}`.
- **New docx `add` types** (§5): `tableOfFigures` (a Figure/Table/Equation caption
  list), `indexEntry` (marks an `XE` field), `index` (builds the alphabetized index)
  and `mergeField` (a `MERGEFIELD` the `office_template` verb fills by name, alongside
  `{{key}}` placeholders). See `office_help {topic:"structural-fields"}`.
- **New xlsx `add` type** (§5): `formControl` (interactive `checkbox` / `optionButton`
  / `spinner` / `comboBox` / `listBox` / `button` on a sheet). See
  `office_help {topic:"properties-xlsx"}`.
- **New xlsx protection props** (set): per-cell `locked`; sheet-path `protected`
  (+ `password` and the `allow*` action flags); workbook-root `protectStructure`
  (+ `protectWindows`). This is Excel's light UI protection (not encryption); AIOffice
  always owns and can lift it. `office_get` reflects the state.
- **New xlsx `numberFormat` presets** (prop-value enum): named codes such as
  `accounting-usd`, `currency-usd/eur/gbp/jpy`, `percent`, `percent2`, `scientific`,
  `fraction`, `thousands`, `integer`, `date-iso`, `datetime-iso`, `time`, `duration`,
  `text`. A preset resolves to its Excel format code; any non-preset string is still
  accepted verbatim as a literal code (existing custom formats are unaffected). See
  `office_help {topic:"number-formats"}`.
- **New structure/fields surfacing** (§6, existing views — no new view names): docx
  `read --view structure` now lists `tablesOfFigures`, `indexes` and `mergeFields`;
  `read --view fields` lists merge fields alongside content controls.
- **New warnings** (§1): `figures_cached`, `index_cached`.

## 7c. 1.3.0 additions (additive within the frozen 1.0 line)

Package **1.3.0** stays on **`surfaceVersion 1.0`** — every change below is purely
**additive** (new prop keys, new prop-value enum members, new `add` type values, a
new `set` addressing form, one new warning), so nothing in §§1–7 was removed or
renamed. The op **kinds** in §5 are unchanged. The 18 CLI verbs / 17 MCP tools are
unchanged.

- **New chart-polish props** (xlsx **and** pptx, accepted both when adding a
  `chart` and when `set`-ting on an existing chart path — `/Sheet1/chart[i]` or
  `/slide[i]/chart[k]`): `dataLabels` (`true` or `{show, position?}`), `legend`
  (`none|right|left|top|bottom`), `axisTitles` (`{category?, value?}`), `trendline`
  (`none|linear|exponential|movingAverage`), `errorBars`
  (`none|stdErr|stdDev|percent`), `gridlines` (`{major?, minor?}`), `secondaryAxis`
  (a list of series names to move to a secondary value axis). `get` on the chart
  reports them. An unsupported sub-value returns `unsupported_feature` listing the
  supported set. See `office_help {topic:"chart-polish"}`.
- **New `conditionalFormat` kinds** (xlsx, prop-value enum): `formula` (an
  `=expression` rule relative to the range's top-left cell), `topBottom` (highest /
  lowest N or N%), `aboveBelowAverage` (above / below the range mean, with an
  optional `stdDev`) — added to the existing
  `cellIs` · `colorScale` · `dataBar` · `containsText` · `iconSet`. See
  `office_help {topic:"conditional-format"}`.
- **New pivot prop** (xlsx, `add type:"pivot"`): `calculatedFields` — a list of
  `{name, formula}` formula fields computed from source-column header names (e.g.
  `{name:"Margin", formula:"=Revenue-Cost"}`); validated at add time and surfaced
  under `get`'s `calculatedFields`.
- **New docx `add` types** (§5): `shape` (a floating DrawingML body shape with a
  preset geometry — `rect|roundRect|ellipse|line|arrow` — at `/body/shape[i]`),
  `textBox` (a body text box at `/body/textBox[i]`) and `formField` (a legacy
  `text|checkbox|dropdown` form field at `/formField[@name=…]`). See
  `office_help {topic:"form-fields"}` and the docx properties help.
- **New pptx `add` type** (§5): `model3d` (a `.glb`/`.gltf` 3D model embedded as a
  real 3DModel media part behind a poster picture fallback, at
  `/slide[i]/model3d[@id=N]`; `src`/`poster` are sandbox-resolved). See
  `office_help {topic:"3d-models"}`.
- **New pptx animation effect** (prop-value enum on `add type:"animation"`):
  `motionPath` (with `path` = `line|arc|circle|custom` and, for `custom`, a
  normalized `points` list) — added to the existing entrance/emphasis/exit
  effects. See `office_help {topic:"animations"}`.
- **New addressing form** (§4, additive): `/theme` on docx — `set /theme` edits the
  theme color scheme (`dk1/lt1/dk2/lt2`, `accent1`…`accent6`, `hlink`, `folHlink`)
  and font scheme (`majorFont`, `minorFont`); `get /theme` reports them. See
  `office_help {topic:"themes"}`.
- **New warning** (§1): `model3d_as_media` (a pptx 3D model was embedded as a 3D
  part behind a poster picture fallback; PowerPoint 2019+ renders the model).

## 7d. 1.4.0 additions (additive within the frozen 1.0 line)

Package **1.4.0** stays on **`surfaceVersion 1.0`** — every change below is purely
**additive** (new evaluation behavior on the existing `set value` verb, new `add`
type values, new prop keys, new addressing forms, extended `template` behavior, one
new error code), so nothing in §§1–7 was removed or renamed. The op **kinds** in §5
are unchanged. The 18 CLI verbs / 17 MCP tools are unchanged.

### xlsx — dynamic arrays, financial functions, what-if data tables

- **Dynamic-array formula evaluation + spill** (new behavior on `set value`):
  setting a single cell to a dynamic-array formula — `FILTER`, `UNIQUE`, `SORT`,
  `SORTBY`, `SEQUENCE`, `RANDARRAY` or `TRANSPOSE` — now EVALUATES it at write time,
  computes the result array, and SPILLS it into the rectangle anchored at the cell
  (the anchor carries the array formula + spill `ref`; every spilled cell carries a
  cached value). These functions no longer emit `formula_not_evaluated`. `get` on
  the anchor reports the new **`spillRange`** field. `RANDARRAY` is deterministic
  here (seeded from its arguments) so round-trips are stable; Excel reseeds on open.
- **Financial functions** (new behavior on `set value`): `RATE`, `IRR`, `XIRR`,
  `NPV`, `PV`, `FV`, `PMT` and `NPER` are now evaluated at write time (iterative ones
  by Newton's method with a bisection fallback) and the cached numeric result is
  written into the cell. These functions no longer emit `formula_not_evaluated`.
- **New `add` type** (§5, additive): `dataTable` — a one- or two-variable what-if
  data table over a rectangular range
  (`{op:add, type:dataTable, path:/Sheet1/A1:C10, props:{rowInput:"B1", colInput:"B2"}}`);
  the corner formula is recomputed across the row/column input axes into a cached
  body carrying the Excel `{=TABLE(…)}` construct. At least one of `rowInput` /
  `colInput` is required.
- **New addressing form** (§4, additive): `/Sheet1/dataTable[i]` (1-based per sheet,
  row-major anchor order) — `get` reports `{body, rowInput, colInput, twoDimensional}`;
  `remove` clears the table's construct and cached body.
- **New error code** (§2): `spill_blocked` — a dynamic-array formula's result would
  spill over non-empty cells; nothing was written (the suggestion names clearing the
  target range). Exit code `2` (user error).

### docx — mail-merge execution, IF fields, page borders

- **Mail-merge execution** (extended behavior on the existing `template` verb /
  `office_template` tool — NOT a new verb): when `--data` / `data` is a JSON **array**
  of record objects, `template` runs a mail merge. With the new **`--output`** flag
  (`office_template` gains an optional `output` param) given a path pattern — `{n}`
  = the 1-based record index, `{Field}` = that record's value — it writes one merged
  document per record; every expanded path is sandbox-resolved (an escaping pattern
  → `sandbox_denied`, a colliding pattern → `invalid_args`). Without `--output` it
  writes a single combined document (the body repeated per record, split by next-page
  sections) back to the source with an auto-snapshot. A single-object `--data` keeps
  the original one-document fill unchanged. Unresolved fields raise a single
  `template_unresolved` warning. Each record fills `{{key}}` markers, `MERGEFIELD`
  fields by name, and `ifField`s.
- **New docx `add` type** (§5, additive): `ifField` — a Word «IF» field
  (`props {field, operator, value, trueText, falseText}`; `operator` ∈
  `= <> > < >= <=`) resolved per record during `template`/mail-merge from the
  record's value for `field`. `get` reads the parsed parts back.
- **New docx `set` prop** (§4, additive): `pageBorder` on a `/section[i]` —
  `{style, color?, widthPt?, sides?}` (`style` ∈
  `single|double|thick|dashed|dotted|wave`, `sides` ∈ `all|top|bottom|left|right`),
  or the string `"none"` to clear it. `get /section[i]` reports it back.

### pptx — zoom navigation, click-trigger animations, table styles

- **New pptx `add` type** (§5, additive): `zoom` — a slide/section/summary zoom
  navigation object added on a slide path (`{op:add, path:/slide[i], type:zoom,
  props:{kind, target?, x?, y?, w?, h?}}`; `kind` ∈ `slide|section|summary`).
- **New addressing form** (§4, additive): `/slide[i]/zoom[k]` (1-based per slide) —
  `get` reports the zoom's kind and target; `remove` drops the frame. A zoom's
  kind/target are immutable in place (remove + re-add to retarget).
- **New pptx animation prop** (additive on `add type:"animation"`): `triggerOn:"@N"`
  plays the effect when another shape (stable id `N`) is clicked (it joins that
  shape's interactive `onClick` sequence). `N` must be a different shape on the
  slide; a bad id is `invalid_path`. `get`/`structure` report the trigger shape.
- **New pptx table props** (additive on `add type:"table"`): `style` (a built-in
  table style — `none|light1|light2|medium1|medium2|medium3|dark1|dark2`) plus the
  banding/edge flags `firstRow`, `lastRow`, `bandRow`, `firstCol`. `get`/`structure`
  report the applied style and flags.

### behavior improvement (backward-compatible)

The dynamic-array and financial-function evaluation above is a **behavior
improvement** within the frozen contract, not a breaking change: cells that used to
carry a `formula_not_evaluated` warning for `FILTER`/`UNIQUE`/`SORT`/`RATE`/`IRR`
(etc.) now carry a **cached value** and the warning no longer fires for them. The
`formula_not_evaluated` warning code itself is unchanged and still fires for other
unevaluated formulas, so an agent that handled the warning continues to work — it
simply stops seeing it for the now-evaluated functions.

## 7e. 1.5.0 additions (additive within the frozen 1.0 line)

Package **1.5.0** stays on **`surfaceVersion 1.0`** — every change below is purely
**additive** (new evaluation behavior on the existing `set value` verb, new `add`
type values, new `set` prop keys, new addressing forms, two new warning codes), so
nothing in §§1–7 was removed or renamed. The op **kinds** in §5 are unchanged. The
18 CLI verbs / 17 MCP tools are unchanged.

### xlsx — scalar function evaluation, scenarios, goal seek

- **Scalar function evaluation** (new behavior on `set value`): `XLOOKUP`, `IFS`,
  `SWITCH`, `LET`, `MAXIFS`, `MINIFS` and `AVERAGEIFS` — which the base engine
  returns `#NAME?` for — are now evaluated at write time and the cached scalar value
  is written. `TEXTSPLIT` joins the spilled dynamic-array family. These no longer emit
  `formula_not_evaluated`. (`TEXTJOIN`, `CONCAT`, `CONCATENATE`, `IFERROR`, `SUMIFS`,
  `COUNTIFS` were already evaluated and are unchanged.)
- **New `add` type** (§5, additive): `scenario` — a named what-if changing-cell set
  (`{op:add, path:/Sheet1, type:scenario, props:{name, cells:{addr:value,…}, comment?}}`)
  saved into the worksheet's scenarios part. `cells` values are constants.
- **New `set` prop** (§4, additive): `applyScenario` on a sheet path
  (`{op:set, path:/Sheet1, props:{applyScenario:"name"}}`) writes the stored values
  into the changing cells and recalculates dependents.
- **New `set` prop** (§4, additive): `goalSeek` on a cell path
  (`{op:set, path:/Sheet1/B1, props:{goalSeek:{targetCell, targetValue}}}`) solves for
  the changing cell's value that makes `targetCell` reach `targetValue` (Newton's
  method + bisection fallback), SETs the cell to the found value, and recalculates.
  The op result reports the found input and achieved target.
- **New addressing form** (§4, additive): `/Sheet1/scenario[@name=…]` (selected by
  name) — `get` reports `{name, comment, cells}`; `remove` drops it; `read
  --view structure` lists each sheet's scenarios.
- **New warning code** (§2): `goal_seek_no_solution` — goal seek did not converge;
  the changing cell is left unchanged and the command still succeeds (the warning
  rides `meta.warnings`).

### docx — table-cell formulas, building blocks, line numbering

- **New `set` prop** (§4, additive): `formula` on a table-cell path
  (`{op:set, path:/body/table[1]/row[3]/cell[2], props:{formula:"=SUM(ABOVE)"}}`) —
  directional aggregates (`SUM`/`AVERAGE`/`PRODUCT`/`COUNT`/`MIN`/`MAX` over
  `ABOVE`/`BELOW`/`LEFT`/`RIGHT`) or cell-ref arithmetic (`=A1*B2`) over the table's
  A1 grid. Computed headlessly and cached as a `w:fldSimple` formula-field result. An
  optional `numberFormat` (preset or a raw `\#` picture, valid only alongside
  `formula`) shapes the cached text.
- **New warning code** (§2): `table_formula_cached` — a table-formula input is itself
  a field; the cached value is still written but Word may refresh it on F9. Non-fatal.
- **New `add` types** (§5, additive): `buildingBlock` (`{op:add,
  path:/buildingBlocks, type:buildingBlock, props:{name, gallery:quickParts|autoText,
  category?, content}}`) stores reusable AutoText / Quick Parts content in the
  glossary part; `buildingBlockRef` (`{op:add, path:/body, type:buildingBlockRef,
  props:{name, position?}}`) inserts a stored block's content into the body.
- **New addressing form** (§4, additive): `/buildingBlock[@name=…]` — `get` reports
  `{name, gallery, category, content}`; `remove` drops it; `read --view structure`
  lists every stored block.
- **New `set` prop** (§4, additive): `lineNumbers` on a `/section[i]` —
  `{start, increment, restart:continuous|newPage|newSection, distance?}` (writes
  `w:lnNumType`), or the string `"none"` to clear it. `get /section[i]` reports it.

### pptx — embedded fonts, action buttons, custom layouts

- **New `add` type** (§5, additive): `font` on `/fonts` (`{op:add, path:/fonts,
  type:font, props:{src, name?, embedAll?, bold?, italic?, boldItalic?}}`) embeds a
  sandbox-resolved `.ttf`/`.otf` file as a font part and registers a `p:embeddedFont`
  in `p:embeddedFontLst` (regular slot by default; `embedAll` plus the per-style files
  fill all four slots). `src` is required and is sandbox-resolved (an escaping `src`
  → `sandbox_denied`, never read).
- **New addressing form** (§4, additive): `/fonts` and `/fonts/font[@name=…]` —
  `get /fonts` lists embedded fonts, `get /fonts/font[@name=…]` reports one, `remove`
  drops the registration and its parts.
- **New `add` type** (§5, additive): `actionButton` on a slide (`{op:add,
  path:/slide[i], type:actionButton, props:{action, target?, x?, y?, w?, h?, label?,
  fill?}}`; `action` ∈ `first|last|next|prev|home|end|slide|url`) — a navigation
  button building on M8 shape hyperlinks. `get` on its shape path reports the resolved
  action/target.
- **New `add type:"layout"` prop** (additive on the M6 layout add): `placeholders` —
  a list of `{type, x, y, w, h}` placeholder shapes that builds a fresh `slideLayout`
  part on a master; a slide then binds to it by name
  (`{op:add, path:/slide[i], type:slide, props:{layoutName:"Hero"}}`). `basedOn` (clone)
  and `placeholders` (build fresh) are mutually exclusive.

### behavior improvement (backward-compatible)

The scalar-function evaluation above is a **behavior improvement** within the frozen
contract, not a breaking change: cells that used to carry a `formula_not_evaluated`
warning for `XLOOKUP`/`IFS`/`SWITCH`/`LET`/`TEXTJOIN` (etc.) now carry a **cached
value** and the warning no longer fires for them. The `formula_not_evaluated` warning
code is unchanged and still fires for `LAMBDA` and the lambda-helpers
(`MAP`/`REDUCE`/`SCAN`/`BYROW`/`BYCOL`), which cannot be evaluated headlessly — so an
agent that handled the warning keeps working.

## 7f. 1.6.0 — distribution release (no surface change)

Package **1.6.0** is a **distribution & onboarding release**: the native binary is the
**same surface** as 1.5.0 — **`surfaceVersion 1.0`**, the **18 CLI verbs / 17 MCP tools**,
every envelope shape, error code, exit code, op/view/tool vocabulary and addressing form in
§§1–7 are unchanged. Nothing was added to or removed from the contract. The release adds
only packaging and onboarding artifacts *around* the binary (an npm package, a Homebrew
formula, one-line install scripts, and the `SKILL.md` / `docs/COOKBOOK.md` /
`docs/INSTALL.md` / `docs/MCP-SETUP.md` / `docs/SIGNING.md` guides). Agents that target the
1.5.0 surface are byte-for-byte compatible with 1.6.0.

## 7g. 1.7.0 — print readiness, camera tool, calc mode (additive, surface 1.0)

Package **1.7.0** keeps **`surfaceVersion 1.0`**, the **18 CLI verbs / 17 MCP tools**, and
every envelope shape, error code, exit code, op/view/tool vocabulary in §§1–7. The changes
are **additive** (new `set` props on an existing path, one new `add` type, one new
addressing form, one new warning, one new behavior on `get /`); nothing in §§1–7 was
removed or renamed.

### xlsx — print completeness (additive `set` props on `/SheetN`)

- **New `set` props** (§4, additive) on a sheet path, extending the M3/1.4 page setup:
  `printTitleRows` (`"1:1"` repeat header rows), `printTitleCols` (`"A:A"`),
  `pageBreaks` (`{rows:[20,40], cols:["F"]}`; `rowBreaks`/`colBreaks` aliases; `[]`
  clears an axis), `fitToPage` (`{fitToWidth, fitToHeight}` page counts **or**
  `{scale}` percent — mutually exclusive), `fitToHeight` (flat twin of the M3
  `fitToWidth`), `centerHorizontally`, `centerVertically`, `printGridlines`,
  `printHeadings`, and `printHeader`/`printFooter`
  (`{left, center, right}` with Excel field codes: `&P` page, `&N` pages, `&D` date,
  `&T` time, `&F` file, `&A` sheet; a section sent as `null`/`""` clears it). `get
  /SheetN` reflects everything under `pageSetup`.

### xlsx — linked picture / camera tool (additive `add` type)

- **New `add` type** (§5, additive): `linkedPicture` on a sheet (`{op:add,
  path:/Sheet1, type:linkedPicture, props:{sourceRange:"A1:D10", anchor:"G2",
  sheet?}}`) mirrors a cell range as a picture. It is an **honest static snapshot** of
  the range's values as of the edit (a true live link is validator-fragile), embedded
  as a real PNG picture; re-add it to refresh. The add fires the new
  `linked_picture_static` warning (§2 table additive).
- **New addressing form** (§4, additive): `/Sheet1/linkedPicture[i]` (1-based per
  sheet, ordered by picture name) — `get` reports `{sourceRange, anchor, …}`; `remove`
  drops it. A linked picture is a real picture, so it ALSO appears under `image[i]`.

### xlsx — workbook calculation mode (additive `set /` props + `get /`)

- **New `set /` props** (§4, additive on the workbook-root set that already carries
  structure protection): `calculationMode` (`auto|manual|autoExceptTables`),
  `iterativeCalc`, `maxIterations`, `maxChange`, `fullPrecision` — write the workbook
  `calcPr`. A single root set may carry both protection and calc props.
- **New behavior on `get /`** (additive): an xlsx `get /` previously returned
  `unsupported_feature`; it now returns `{calculation, workbookProtection}`. An agent
  that handled the old error keeps working (the call no longer errors — it succeeds).

### docx — drop caps, picture watermark, more fields, numbered/deeper equations

- **New `set` props** (§4, additive) on a body paragraph: `dropCap`
  (`drop|margin|none`), `dropCapLines` (height in lines, default 3), `dropCapFont`
  (font of the dropped letter) — restructure the paragraph into Word's drop-cap shape
  (`w:framePr w:dropCap`); `dropCap:"none"|false` removes it.
- **New `add` prop on the existing `watermark` type** (§5, additive): `props.image`
  (a sandbox-resolved PNG/JPEG, optional `washout`) makes a **picture** watermark
  instead of a text one; `get /watermark[1]` reports `{image, washout}`.
- **New `field` kinds** (§5, additive — the existing `field` add type gains kinds):
  `styleRef` (STYLEREF; `props.styleRef` = style name), `symbol` (SYMBOL;
  `props.charCode` decimal/`0x`-hex, optional `props.symbolFont`), `quote` (QUOTE;
  `props.quoteText`). All write real Word fields; Word refreshes on open / F9.
- **New `add type:equation` prop** (§5, additive): `props.number` on a **display**
  equation — `true` auto-increments the document equation counter, a string sets the
  label verbatim. Cached numbering fires the new `equation_numbers_cached` warning;
  a number on an inline equation is `invalid_args`.
- **New addressing form** (§4, additive): `/equation[@num=...]` addresses a numbered
  display equation by its number (`/equation[@num=1.1]` — the numeric core matches a
  `"(1.1)"` label; a whole-number label answers to `/equation[@num=1]`).
- **Deeper LaTeX→OMML** (backward-compatible — strictly *more* LaTeX now renders, so
  *fewer* `equation_partial` warnings; no previously-supported input changed meaning):
  equation arrays `\begin{aligned|gathered|cases}` (`m:eqArr`), `\binom`/`\dbinom`/
  `\tbinom`, `\overbrace`/`\underbrace`, multi-integrals `\iint`/`\iiint`, more accents
  (`\hat`/`\vec`/`\tilde`), and many more relations/arrows. The **same shared Core
  converter** drives docx and pptx, so pptx equations gain all of these automatically.

### pptx — notes/handout masters, animation timing, table-cell alignment

- **New `set` set-paths** (§4, additive — singletons, no index): `/notesMaster`
  (`{background, bodyFont}`) and `/handoutMaster` (`{background,
  headerFooter:{header, footer, date, pageNumber}, slidesPerPage}`,
  `slidesPerPage ∈ {1,2,3,4,6,9}`). Both parts are created on first edit and reported
  by `get` and `read --view structure`.
- **New `animation` timing props** (§4, additive on `set /slide[i]/animation[k]` — a
  retime, applied after the animation is added): `repeat`
  (`none | N | untilClick | untilNext`), `rewind` (bool), `autoReverse` (bool).
  `read --view structure` reports them.
- **New `set` props on a pptx table cell** (§4, additive): `valign`
  (`top|middle|bottom`), `marginLeft`/`marginRight`/`marginTop`/`marginBottom`
  (lengths), `textDirection` (`horizontal|vertical|vertical270|stacked`).

Agents that target the 1.6.0 surface are byte-for-byte compatible with 1.7.0.

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
