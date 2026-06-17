# Changelog

All notable changes to AIOffice are recorded here. The package **Version** follows
semantic versioning; the AI-facing **`surfaceVersion`** (the frozen contract in
[CONTRACT.md](CONTRACT.md)) moves independently and only bumps on a breaking change.

## 1.9.0 — fidelity: text autofit + LibreOffice render engine (additive)

`surfaceVersion` stays **1.0**; **18 CLI verbs / 17 MCP tools** unchanged. Additive only —
one new optional pptx shape prop, one new optional `render --engine` option (default behavior
byte-for-byte unchanged), an additive `renderers` field on `doctor`/`office_status`, one new
warning (`engine_fallback`), one new help topic (`render-engines`). Agents that target the
1.8.0 surface are byte-for-byte compatible with 1.9.0.

### Added — pptx

- **Text autofit** (`set`/`add` shape prop `autofit`): `"shrink"` → `a:normAutofit`
  (PowerPoint shrinks text to fit its box — the fix for agent text overflow), `"resize"` →
  `a:spAutoFit` (shape grows to fit), `"none"` → `a:noAutofit`; object form
  `{mode:"shrink", fontScale, lineSpaceReduction}` writes an explicit scale. `get` reports it.

### Added — render / cross-cutting

- **Optional LibreOffice render engine** (`render --engine chromium|soffice|auto`, default
  `chromium`): `soffice` produces true-fidelity `--to pdf` (whole document) and `--to png`
  (per page via `pdftoppm`, page from `--scope`); `svg|html|text` fall back to the native
  engine with an `engine_fallback` warning. `auto` uses LibreOffice when present, else chromium.
- **`doctor` / `office_status` `renderers` probe**: `{chromium, libreoffice, poppler}` each
  `{found, path}` (`AIOFFICE_SOFFICE` / `AIOFFICE_PDFTOPPM` overrides).
- **New help topic** `render-engines`.

## 1.8.0 — design depth: fills, fonts, brand colors (additive)

`surfaceVersion` stays **1.0**; **18 CLI verbs / 17 MCP tools** unchanged. Additive visual
primitives so an agent can derive a per-brand look instead of default-white. Agents that
target the 1.7.0 surface are byte-for-byte compatible with 1.8.0.

### Added — pptx

- **Gradient fill** (`gradient` prop) on shapes and slide/master/layout backgrounds; **image
  (picture) fill** (`image` prop) on shapes and backgrounds; **master theme fonts**
  (`majorFont`/`minorFont` on `set /master[m]`).

### Added — xlsx

- **`seriesColors`** chart prop (brand-match chart palette, on `add`/`set` chart); expanded
  **numberFormat presets** (`usd-millions`, `usd-thousands`, `millions`, `thousands-k`,
  `percent-0|1|2`, `accounting-eur|gbp`); **`gradientFill`** cell/range prop.

### Added — docx

- **Paragraph visuals** on `/body/p[i]`: `shading`, `border`, `spacingBefore`/`spacingAfter`,
  `indentLeft`/`indentRight` (callout & lead blocks); **custom-style** `font`/`next`;
  **run-level `font`** (implements the previously documented-but-missing prop).

## 1.7.0 — print readiness, camera tool, calc mode, deeper equations (additive)

`surfaceVersion` stays **1.0**; **18 CLI verbs / 17 MCP tools** unchanged. Every change is
**additive** — new `set` props on existing paths, one new `add` type (`linkedPicture`), two
new `set`-paths (`/notesMaster`, `/handoutMaster`), one new addressing form
(`/equation[@num=…]`), two new warnings, deeper LaTeX support — nothing removed or renamed.
Agents that target the 1.6.0 surface are byte-for-byte compatible with 1.7.0.
**2331 tests** across 7 projects (Core 180 · Word 681 · Excel 619 · Pptx 709 · MCP 87 ·
Preview 24 · Render 31), green on macOS + Windows.

### Added — xlsx

- **Print completeness** (`set` props on `/SheetN`): `printTitleRows`/`printTitleCols`
  (repeat bands), `pageBreaks` (`{rows,cols}`), `fitToPage` (`{fitToWidth,fitToHeight}` or
  `{scale}`), `fitToHeight`, `centerHorizontally`/`centerVertically`,
  `printGridlines`/`printHeadings`, and `printHeader`/`printFooter`
  (`{left,center,right}` field-code strings: `&P &N &D &T &F &A`). `get /SheetN` reflects them.
- **Linked picture / camera tool** (`add type:linkedPicture` on a sheet): mirrors a cell
  range as a picture — an honest **static snapshot** (a true live link is validator-fragile),
  fired with the new `linked_picture_static` warning. Addressed as
  `/SheetN/linkedPicture[i]` (`get`/`remove`).
- **Workbook calculation settings** (`set /`): `calculationMode`
  (`auto|manual|autoExceptTables`), `iterativeCalc`, `maxIterations`, `maxChange`,
  `fullPrecision` — one root set can carry these alongside structure protection. `get /`
  now returns `{calculation, workbookProtection}` (was `unsupported_feature`).

### Added — docx

- **Drop caps** (`set` props on a paragraph): `dropCap` (`drop|margin|none`),
  `dropCapLines`, `dropCapFont`.
- **Picture watermark** (`add type:watermark` with `props.image` + optional `washout`)
  alongside the existing text watermark.
- **New field kinds** (`add type:field`): `styleRef` (STYLEREF), `symbol` (SYMBOL,
  `charCode` + optional `symbolFont`), `quote` (QUOTE).
- **Numbered display equations**: `add type:equation` `props.number` (`true` auto-increments,
  or a verbatim label); cached numbering fires the new `equation_numbers_cached` warning;
  addressed by `/equation[@num=…]`.

### Added — equations (shared Core converter — docx + pptx)

- **Deeper LaTeX→OMML**: equation arrays `\begin{aligned|gathered|cases}` (`m:eqArr`),
  `\binom`/`\dbinom`/`\tbinom`, `\overbrace`/`\underbrace`, multi-integrals
  `\iint`/`\iiint`, more accents (`\hat`/`\vec`/`\tilde`), and many more relations/arrows.
  Backward-compatible — strictly *more* LaTeX renders, so *fewer* `equation_partial`
  warnings. pptx equations inherit all of this automatically (same Core converter).

### Added — pptx

- **Notes & handout masters** (`set` set-paths): `/notesMaster` (`{background, bodyFont}`)
  and `/handoutMaster` (`{background, headerFooter:{header,footer,date,pageNumber},
  slidesPerPage∈{1,2,3,4,6,9}}`). Created on first edit; reported by `get` + structure.
- **Animation timing** (`set /slide[i]/animation[k]` retime, applied after the add):
  `repeat` (`none|N|untilClick|untilNext`), `rewind`, `autoReverse`.
- **Table-cell alignment & spacing** (`set` on a cell): `valign` (`top|middle|bottom`),
  `marginLeft`/`marginRight`/`marginTop`/`marginBottom`, `textDirection`.

### Changed

- `Directory.Build.props` `Version` → **1.7.0**; npm `package.json`, the Homebrew formula
  `version`, and the `install.sh`/`install.ps1` fallback versions all pin to **1.7.0**.
- `office_edit` schema + `aioffice schema`/help surfaces and `office_help`/`aioffice help`
  topics extended additively (new `print-setup` and `masters` topics; equations, animations,
  fields, watermark, drop-cap, table-cell, calc docs deepened). MCP tool surface stays inside
  its 3500-token budget.

## 1.6.0 — distribution & onboarding release (no capability change)

`surfaceVersion` stays **1.0** and the native binary is **byte-for-byte the same surface**
as 1.5.0 — **18 CLI verbs / 17 MCP tools**, every envelope, error code and addressing form
unchanged. This release adds the packaging and onboarding layer *around* the binary, almost
entirely **outside the C#/.NET source tree**, so the build and tests are unaffected.
**2125 tests** across 7 projects (Core 124 · Word 612 · Excel 572 · Pptx 675 · MCP 87 ·
Preview 24 · Render 31), green on macOS + Windows.

### Added

- **npm package** (`npm/`): `npm i -g aioffice` (or `npx aioffice …`). A dependency-free
  wrapper that, on install, downloads the matching native binary from the GitHub release
  and **SHA256-verifies** it against the release `SHA256SUMS`; the binary is not shipped in
  the tarball. A thin bin shim spawns the binary with full stdio passthrough (so
  `aioffice mcp` stays transparent) and lazily installs on first run if `--ignore-scripts`
  skipped the postinstall. Env overrides `AIOFFICE_DOWNLOAD_VERSION` /
  `AIOFFICE_DOWNLOAD_BASEURL` for pinning and private mirrors.
- **Homebrew formula** (`dist/Formula/aioffice.rb`): `brew install onecer/tap/aioffice`
  installs the prebuilt per-platform binary; the formula's `test` block runs
  `aioffice version`.
- **One-line install scripts**: `dist/install.sh` (POSIX sh, macOS/Linux) and
  `dist/install.ps1` (Windows PowerShell) — detect platform → download → SHA256-verify →
  install → PATH hint; `install.sh` strips the macOS Gatekeeper quarantine attribute.
- **Onboarding docs**: `SKILL.md` (AI-facing skill guide), `docs/COOKBOOK.md` (10
  copy-paste recipes), `docs/INSTALL.md` (all four install paths), `docs/MCP-SETUP.md`
  (Claude Desktop / Claude Code / Cursor / generic stdio / TonoBraid configs), and
  `docs/SIGNING.md` (code-signing / notarization roadmap). README (EN + 简体中文) gains an
  INSTALL section (npm → brew → curl → direct download) and a "Use it from your agent (MCP)"
  section.

### Changed

- `Directory.Build.props` `Version` → **1.6.0**, so the tagged release binaries report
  `1.6.0` and the npm `package.json` / Homebrew formula / install-script defaults all pin
  to **1.6.0** consistently.
- Homebrew formula `license` corrected to **Apache-2.0** to match the repository `LICENSE`.

### Publishing (run by a maintainer — not done in this release)

- npm: `cd npm && npm publish --access public` (after `npm login`); optionally add the
  `NPM_TOKEN` secret to enable the `.github/workflows/npm-publish.yml` auto-publish on tag.
- Homebrew: create `onecer/homebrew-tap` and commit `dist/Formula/aioffice.rb` as
  `Formula/aioffice.rb`, filling the four `sha256` values from the v1.6.0 `SHA256SUMS`.
- Install scripts are served straight from `main` — no publish step.

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
