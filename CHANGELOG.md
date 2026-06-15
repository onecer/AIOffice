# Changelog

All notable changes to AIOffice are recorded here. The package **Version** follows
semantic versioning; the AI-facing **`surfaceVersion`** (the frozen contract in
[CONTRACT.md](CONTRACT.md)) moves independently and only bumps on a breaking change.

## 1.5.0 — fifth post-1.0 feature release (additive)

`surfaceVersion` stays **1.0** — every change is **additive** within the frozen 1.0
contract line (new write-time evaluation behavior on `set value`, new `add` type
values, new `set` prop keys, new addressing forms, two new warning codes). Nothing was
removed or renamed; the 18 CLI verbs / 17 MCP tools stand unchanged. **The modern
scalar-function gap is closed** (XLOOKUP / IFS / SWITCH / LET / MAXIFS / MINIFS /
AVERAGEIFS now evaluate), and the spreadsheet **what-if toolkit** gains the Scenario
Manager and Goal Seek. Word gains **table-cell formula fields**, **reusable building
blocks** and **section line numbering**; PowerPoint gains **embedded fonts**, **action
buttons** and **custom layouts with placeholders**.
**2125 tests** across 7 projects (Core 124 · Word 612 · Excel 572 · Pptx 675 ·
MCP 87 · Preview 24 · Render 31), green on macOS + Windows.

### Added

- **Scalar function evaluation** (xlsx): `=XLOOKUP`, `=IFS`, `=SWITCH`, `=LET`,
  `=MAXIFS`, `=MINIFS` and `=AVERAGEIFS` — which ClosedXML returns `#NAME?` for — are
  now evaluated at write time and the cached value is written, so headless readers see
  a real result with no `formula_not_evaluated` warning. XLOOKUP does exact (default)
  or approximate (`match_mode` -1/1) match and returns `if_not_found` on a miss; LET
  binds names left-to-right then evaluates the calculation; the conditional
  aggregates honour numeric/text criteria operators. `TEXTJOIN`, `CONCAT`,
  `CONCATENATE`, `IFERROR`, `SUMIFS` and `COUNTIFS` were already evaluated by the base
  engine and keep working. See `aioffice help formulas`.
- **`TEXTSPLIT` dynamic array** (xlsx): `=TEXTSPLIT(text, col_delim, [row_delim], …)`
  joins the FILTER/UNIQUE/SORT spill family — it splits text into a spilled array
  (1-D by the column delimiter, 2-D with a row delimiter); the text/delimiters may be
  string literals or cell references.
- **Scenario Manager** (xlsx): a new `add` type `scenario`
  (`props {name, cells:{addr:value,…}, comment?}`) saves a named set of changing cells
  and the constant values they take into the worksheet's scenarios part.
  `{op:set, props:{applyScenario:"name"}}` writes those values into the cells and
  recalculates dependents (real cached values). Addressed by
  `/Sheet1/scenario[@name=…]` for `get`/`remove`; `read --view structure` lists each
  sheet's scenarios. See `aioffice help scenarios`.
- **Goal Seek** (xlsx): `{op:set, path:/Sheet1/B1, props:{goalSeek:{targetCell,
  targetValue}}}` solves for the value of the changing cell that makes the target
  formula cell equal `targetValue` (Newton's method with a bisection fallback), SETs
  the changing cell to the found value and recalculates; the op result reports the
  found input and achieved target. A compute action that persists its result —
  additive behavior on the frozen `set` verb, no new verb. See `aioffice help
  goal-seek`.
- **New warning code** `goal_seek_no_solution`: when goal seek does not converge the
  changing cell is left unchanged and this non-fatal warning rides `meta.warnings`
  (the command still succeeds — correctness over coverage; never a wrong value).
- **Word table-cell formulas** (docx): a table cell gains a `formula` prop —
  `=SUM(ABOVE)` / `=AVERAGE(LEFT)` / `=PRODUCT(ABOVE)` directional aggregates or
  cell-ref arithmetic (`=A1*B2`) over the table's A1 grid. It becomes a `w:fldSimple`
  formula field with the value computed headlessly and cached, so `read --view text`,
  `get` and Word (before F9) agree. An optional `numberFormat` (preset or a raw `\#`
  picture) shapes the cached text. See `aioffice help table-formulas`.
- **New warning code** `table_formula_cached`: when a table-formula input is itself a
  field the cached value is still written but this non-fatal warning flags that Word
  may refresh to a different number on F9.
- **Word building blocks** (docx): a new `add` type `buildingBlock`
  (`props {name, gallery:quickParts|autoText, category?, content}`) stores reusable
  AutoText / Quick Parts content in the glossary part; `buildingBlockRef`
  (`props {name, position?}`) inserts a stored block's content into the body.
  `get /buildingBlock[@name=…]`, `remove`, and `read --view structure` round-trip
  them. See `aioffice help building-blocks`.
- **Word section line numbering** (docx): a section gains a `lineNumbers` prop
  (`{start, increment, restart:continuous|newPage|newSection, distance?}` or `"none"`)
  writing `w:lnNumType`. `get /section[i]` reads it back. See
  `aioffice help line-numbers`.
- **Embedded fonts** (pptx): a new `add` type `font` on `/fonts`
  (`props {src, name?, embedAll?, bold?, italic?, boldItalic?}`) embeds a
  sandbox-resolved `.ttf`/`.otf` file as a font part and registers a `p:embeddedFont`
  in `p:embeddedFontLst` (regular slot by default; `embedAll` plus the per-style files
  fill all four slots). `src` is required and must live inside the workspace —
  aioffice cannot pull a system font; an escaping `src` is `sandbox_denied`.
  `get /fonts`, `get /fonts/font[@name=…]` and `remove` round-trip them. See
  `aioffice help embedded-fonts`.
- **Action buttons** (pptx): a new `add` type `actionButton` on a slide
  (`props {action, target?, x?, y?, w?, h?, label?, fill?}`) creates a navigation
  button — `first/last/next/prev/home/end` show-jumps, `slide` (slide-jump with a real
  relationship) or `url` (external link) — building on M8 shape hyperlinks; the SVG
  render draws the glyph. See `aioffice help action-buttons`.
- **Custom layouts with placeholders** (pptx): the M6 `add` type `layout` is extended
  with a `placeholders` prop (a list of `{type, x, y, w, h}` placeholder shapes) that
  builds a fresh `slideLayout` part on a master; a slide then binds to it by name
  (`add /slide[i] type:slide props:{layoutName:"Hero"}`). `basedOn` (clone) and
  `placeholders` (build fresh) are mutually exclusive. See `aioffice help layouts`.

### Honest non-support (unchanged, documented)

- Stored `LAMBDA` and the lambda-helpers (`MAP`/`REDUCE`/`SCAN`/`BYROW`/`BYCOL`) still
  emit `formula_not_evaluated` — they cannot be evaluated headlessly and Excel
  computes them on open. The formula text is saved with no stale cached value.

## 1.4.0 — fourth post-1.0 feature release (additive)

`surfaceVersion` stays **1.0** — every change is **additive** within the frozen 1.0
contract line (new evaluation behavior on `set value`, new `add` type values, new
prop keys, new addressing forms, extended `template` behavior, one new error code).
Nothing was removed or renamed; the op **kinds** are unchanged and the 18 CLI verbs /
17 MCP tools stand. **The long-standing dynamic-array gap is closed**: FILTER /
UNIQUE / SORT (and friends) now evaluate and spill instead of only being recognised.
**2016 tests** across 7 projects (Core 124 · Word 580 · Excel 545 · Pptx 625 ·
MCP 87 · Preview 24 · Render 31), green on macOS + Windows.

### Added

- **Dynamic-array evaluation + spill** (xlsx): setting a cell to `=FILTER`,
  `=UNIQUE`, `=SORT`, `=SORTBY`, `=SEQUENCE`, `=RANDARRAY` or `=TRANSPOSE` now
  EVALUATES the formula and SPILLS the result array into the rectangle anchored at the
  cell (anchor keeps the array formula; every spilled cell carries a cached value).
  These no longer emit `formula_not_evaluated`; `get` on the anchor reports a
  `spillRange`. `RANDARRAY` is seeded deterministically for stable round-trips. See
  `aioffice help formulas`.
- **Financial functions** (xlsx): `RATE`, `IRR`, `XIRR`, `NPV`, `PV`, `FV`, `PMT` and
  `NPER` are now evaluated at write time (iterative ones by Newton's method with a
  bisection fallback) and the cached numeric value is written — no more
  `formula_not_evaluated` for them.
- **What-if data tables** (xlsx): a new `add` type `dataTable` builds a one- or
  two-variable data table over a range (`props {rowInput?, colInput?}`); the corner
  formula is recomputed across the input axes into a cached body carrying the Excel
  `{=TABLE(…)}` construct. Addressed by `/Sheet1/dataTable[i]`. See
  `aioffice help data-tables`.
- **New error code** `spill_blocked` (exit 2): a dynamic array whose result would
  overwrite non-empty cells writes nothing; the suggestion names the range to clear.
- **Mail-merge execution** (docx): the existing `template` verb now runs a mail merge
  when `--data` is a JSON **array** of records — one merged document per record with
  the new **`--output`** flag (`{n}` = record index, `{Field}` = a record value; every
  expanded path sandbox-resolved), or a single combined document (a section per
  record) without it. `office_template` gains an optional `output` param. A single
  object `--data` still fills one document unchanged. See `aioffice help mail-merge`.
- **IF fields** (docx): a new `add` type `ifField` adds a Word «IF» field
  (`props {field, operator, value, trueText, falseText}`) resolved per record during a
  merge. See `aioffice help page-borders`.
- **Page borders** (docx): a `pageBorder` `set` prop on `/section[i]`
  (`{style, color?, widthPt?, sides?}`, or `"none"`). See `aioffice help page-borders`.
- **Slide zoom** (pptx): a new `add` type `zoom` adds a slide/section/summary zoom
  navigation object on a slide (`props {kind, target?, x?, y?, w?, h?}`), addressed by
  `/slide[i]/zoom[k]`. See `aioffice help zoom`.
- **Click-trigger animations** (pptx): a new `triggerOn:"@N"` prop on `add
  type:"animation"` plays the effect when another shape (stable id `N`) is clicked.
  See `aioffice help animations`.
- **Table styles** (pptx): `add type:"table"` accepts a built-in `style`
  (`none|light1|light2|medium1|medium2|medium3|dark1|dark2`) plus banding/edge flags
  `firstRow`, `lastRow`, `bandRow`, `firstCol`. See `aioffice help table-styles`.
- **New help topics**: `formulas`, `data-tables`, `mail-merge`, `page-borders`,
  `zoom`, `table-styles` (CLI + `office_help`); `animations` extended with `triggerOn`.

### Notes

- **The dynamic-array / financial evaluation is backward-compatible.** Cells that used
  to carry a `formula_not_evaluated` warning for these functions now carry a cached
  value, and the warning no longer fires for them. The warning code is unchanged and
  still fires for other unevaluated formulas — an agent that handled it keeps working.
- `surfaceVersion` stays **1.0**; the 18 CLI verbs and 17 MCP tools are unchanged
  (the new things are `add` types, prop keys, addressing forms and extended
  `template` behavior — no new verb or tool). The MCP tool surface stays inside its
  token budget.

## 1.3.0 — third post-1.0 feature release (additive)

`surfaceVersion` stays **1.0** — every change is **additive** within the frozen 1.0
contract line (new prop keys, new prop-value enum members, new `add` type values, a
new `/theme` addressing form, one new warning). Nothing was removed or renamed; the
op **kinds** are unchanged and the 18 CLI verbs / 17 MCP tools stand. **1924 tests**
across 7 projects (Core 124 · Word 556 · Excel 522 · Pptx 580 · MCP 87 · Preview 24 ·
Render 31), green on macOS + Windows.

### Added

- **Chart polish** (xlsx **and** pptx): additive presentation props accepted both
  when adding a `chart` and when `set`-ting on an existing chart path
  (`/Sheet1/chart[i]`, `/slide[i]/chart[k]`) — `dataLabels` (`true` or
  `{show, position?}`), `legend` (`none|right|left|top|bottom`), `axisTitles`
  (`{category?, value?}`), `trendline` (`none|linear|exponential|movingAverage`),
  `errorBars` (`none|stdErr|stdDev|percent`), `gridlines` (`{major?, minor?}`) and
  `secondaryAxis` (named series → a secondary value axis). `get` reports them; every
  combination is OpenXmlValidator-clean. See `aioffice help chart-polish`.
- **Advanced conditional formatting** (xlsx): three new `conditionalFormat` kinds —
  `formula` (an `=expression` rule), `topBottom` (top/bottom N or N%) and
  `aboveBelowAverage` (above/below the range mean, optional `stdDev`). See
  `aioffice help conditional-format`.
- **Pivot calculated fields** (xlsx): a new `calculatedFields` prop on `add
  type:"pivot"` — `[{name, formula}]` formula fields computed from source-column
  headers (e.g. `Margin = Revenue - Cost`), validated at add time and surfaced by
  `get`.
- **Body shapes & text boxes** (docx): new `add` types `shape` (a floating DrawingML
  shape — `rect|roundRect|ellipse|line|arrow`, at `/body/shape[i]`) and `textBox`
  (at `/body/textBox[i]`), each with fill/line/inline-text, edited by `get`/`set`/
  `remove`.
- **Legacy form fields** (docx): a new `add` type `formField` — `text|checkbox|
  dropdown` fields addressed by `/formField[@name=…]`, valued by `set`, and listed by
  `read --view fields`. See `aioffice help form-fields`.
- **Theme editing** (docx): `set /theme` edits the document theme color scheme
  (`dk1/lt1/dk2/lt2`, `accent1`…`accent6`, `hlink`, `folHlink`) and fonts
  (`majorFont`, `minorFont`); `get /theme` reports them. See `aioffice help themes`.
- **3D models** (pptx): a new `add` type `model3d` embeds a `.glb`/`.gltf` as a real
  3DModel media part behind a poster picture fallback (PowerPoint 2019+ renders it);
  `src`/`poster` are sandbox-resolved and the add carries a `model3d_as_media`
  warning. Addressed by `/slide[i]/model3d[@id=N]`. See `aioffice help 3d-models`.
- **Motion-path animations** (pptx): a new `motionPath` animation effect with `path`
  `line|arc|circle|custom` (custom takes a normalized `points` list); `read --view
  structure` lists it. See `aioffice help animations`.
- **New help topics**: `chart-polish`, `conditional-format`, `themes`, `3d-models`,
  `form-fields`, `animations` (CLI + `office_help`).

### Notes

- New warning: `model3d_as_media`. Contract additions catalogued in
  [CONTRACT.md](CONTRACT.md) §7c. No existing behavior, prop, type, view, op, verb or
  tool changed.

## 1.2.0 — second post-1.0 feature release (additive)

`surfaceVersion` stays **1.0** — every change is **additive** within the frozen 1.0
contract line (new `add` type values, new prop keys, two new warnings). Nothing was
removed or renamed; the op **kinds** are unchanged (`group`/`ungroup` are `add`
**types**, not new op kinds) and the 18 CLI verbs / 17 MCP tools stand. **1807
tests** across 7 projects (Core 124 · Word 529 · Excel 477 · Pptx 535 · MCP 87 ·
Preview 24 · Render 31), green on macOS + Windows.

### Added

- **SmartArt** (pptx, create + read): a new `add` type `smartart` builds a real
  diagram — layouts `list` · `process` · `hierarchy` · `orgChart` · `cycle`, each
  mapping to a built-in PowerPoint layout that regenerates on open. Nodes are a flat
  `{text, level}` list (0-based level builds the hierarchy). `office_get
  /slide[i]/smartart[k]` returns the layout and node tree; `read --view structure`
  lists each diagram. Editing nodes in place is `unsupported_feature` (rebuild).
- **Connectors** (pptx): a new `add` type `connector` wires a `p:cxnSp` between two
  shapes — `kind` straight/elbow/curved, `startArrow`/`endArrow` none/arrow/triangle,
  `color`, `width`, `name`. Returns `/slide[i]/shape[@id=N]`.
- **Grouping** (pptx): new `add` types `group` (wrap two or more shapes →
  `/slide[i]/group[@id=N]`, children addressable as `…/group[@id=N]/shape[@id=M]`) and
  `ungroup` (dissolve a group, promoting children with absolute coordinates). These
  are `add` **types**, not new op kinds.
- **Table of figures** (docx): a new `add` type `tableOfFigures` collects
  Figure/Table/Equation captions into a navigable list; entries come from cached
  captions with a `figures_cached` warning (Word repaginates on open).
- **Index** (docx): new `add` types `indexEntry` (marks an `XE` field) and `index`
  (builds the alphabetized index, `columns` configurable); page numbers cached with an
  `index_cached` warning.
- **Mail-merge fields** (docx): a new `add` type `mergeField` inserts a `MERGEFIELD`
  the `template` verb fills by name — from the **same** `--data` map that fills
  `{{key}}` placeholders.
- **Form controls** (xlsx): a new `add` type `formControl` authors interactive legacy
  controls — `checkbox`, `optionButton`, `spinner`, `comboBox`, `listBox`, `button` —
  with `linkedCell`, `items`/`listFillRange`, `min`/`max`/`increment`. `read --view
  structure` lists per-sheet controls.
- **Protection** (xlsx): per-cell `locked`; sheet-path `protected` (+ `password` and
  `allow*` action flags); workbook-root `protectStructure` (+ `protectWindows`).
  Excel's light UI protection (not encryption); AIOffice always owns and can lift it.
  `office_get` reflects the state.
- **numberFormat presets** (xlsx): named codes for the `numberFormat` prop —
  `accounting-usd`, `currency-usd/eur/gbp/jpy`, `percent`, `percent2`, `scientific`,
  `fraction`, `thousands`, `integer`, `date-iso`, `datetime-iso`, `time`, `duration`,
  `text`. A preset resolves to its Excel code; any non-preset string stays a literal.
- **Structure / fields surfacing** (docx, existing views): `read --view structure`
  now lists `tablesOfFigures`, `indexes` and `mergeFields`; `read --view fields` lists
  merge fields alongside content controls.
- New `office_help` topics: `smartart`, `connectors`, `number-formats`,
  `structural-fields`; the `properties-pptx` and `properties-xlsx` topics document the
  new vocabulary.
- New warnings: `figures_cached`, `index_cached`.

### Changed

- Package `Version` → `1.2.0`. `surfaceVersion` unchanged at `1.0`.
- CLI `schema` / MCP `office_schema` and `office_help` surface the additive
  vocabulary; the `SchemaConsistencyTests` and token-budget guards stay green (the MCP
  tool surface grew by enum text only and stays inside its 3500-token ceiling).

## 1.1.0 — first post-1.0 feature release (additive)

`surfaceVersion` stays **1.0** — every change is **additive** within the frozen 1.0
contract line (new enum/prop values, one new view, one new warning). Nothing was
removed or renamed; the 18 CLI verbs / 17 MCP tools are unchanged. **1692 tests**
across 7 projects (Core 124 · Word 507 · Excel 430 · Pptx 489 · MCP 87 · Preview 24
· Render 31), green on macOS + Windows.

### Added

- **Expanded chart kinds** (xlsx **and** pptx): `doughnut`, `radar`, `bubble`,
  `stackedBar`, `percentStackedBar`, `stackedArea`, `combo` — alongside the existing
  `bar` · `line` · `pie` · `scatter` · `area`. `bubble` takes an x/y/size triple per
  point; `combo` draws the first series as columns plus the rest as a line (≥2
  series). Unsupported kinds still return `unsupported_feature` listing the expanded
  supported set.
- **iconSet conditional formatting** (xlsx): a new `conditionalFormat` kind painting
  3/4/5-icon glyph sets (e.g. `3TrafficLights1`, `3Arrows`, `4Rating`, `5Quarters`),
  with `reverse` / `showValue` options.
- **Citations & bibliography** (docx): new `add` types `source` (path `/sources`),
  `citation` (a `CITATION` field), and `bibliography` (renders every cited source,
  styles `APA` · `MLA` · `Chicago`). A new `read --view sources` lists the store, and
  a new `bibliography_cached` warning flags the cached entries (Word rebuilds on `F9`).
- **Embedded media** (pptx): a new `add` type `media` embedding audio/video
  (mp4/mov/m4a/mp3/wav) as a `p:pic` with an `a:videoFile`/`a:audioFile`. `src` (and
  the optional `poster` image) is sandbox-resolved before any byte is read.
- **Text & shape effects**: new `set` props `shadow`, `glow`, `reflection`,
  `outline` — run-level on docx (Word 2010 `w14:` effects) and shape-level on pptx
  (`a:effectLst`).
- **New slide transitions** (pptx): `split`, `reveal`, `cut`, `zoom` — added to the
  existing `none` · `fade` · `push` · `wipe`.
- New `office_help` topics: `docx/citation`, `docx/effect`, `pptx/media`,
  `pptx/effect`; the `charts` / `conditional-format` / `transition` topics document
  the expanded vocabulary.

### Changed

- Package `Version` → `1.1.0`. `surfaceVersion` unchanged at `1.0`.
- CLI `schema` / MCP `office_schema` and `office_help` surface the additive
  vocabulary; the `SchemaConsistencyTests` guard keeps the two surfaces in lock-step.

## 1.0.0 — the API-stability release

**Stabilization, not new features.** `surfaceVersion` promoted to **`1.0`**;
[CONTRACT.md](CONTRACT.md) finalized as the frozen v1.0 surface with a Stability
promise and a Known-limitations section. The CLI and MCP introspection surfaces were
locked in lock-step on the shared vocabularies (a `SchemaConsistencyTests` guard
reddens CI on drift); the `read --view` enum + `edit` op list expose `embeds` /
`extract` on **both** surfaces; xlsx read views echo `data.view`; xlsx convert reports
chart/pivot/image drops via `data.dropped` + `convert_lossy` (no more silent loss);
the full warning-code vocabulary was centralized, documented, and surfaced in `schema`
as `data.warningCodes`. 18 verbs / 17 MCP tools, green on macOS + Windows.

## 0.11.0 (M10) — document intelligence II

Embedded objects across all three formats (`add type:embed`, `read --view embeds`,
the new `extract` op writing the payload back byte-identical); pptx equations through
a shared LaTeX → OMML converter in `AIOffice.Core.Equations`; 1.0 contract prep
(unified `properties` envelope, `surfaceVersion` declared, CONTRACT.md drafted). Still
17 MCP tools.

## 0.10.0 (M9) — cross-format convert

The cross-format `convert` verb plus 1.0 hardening (standard core properties).

## 0.9.0 (M8) — semantic diff

The semantic `diff` verb, captions / cross-references, slicers, shape
hyperlinks/actions.
