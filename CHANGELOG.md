# Changelog

All notable changes to AIOffice are recorded here. The package **Version** follows
semantic versioning; the AI-facing **`surfaceVersion`** (the frozen contract in
[CONTRACT.md](CONTRACT.md)) moves independently and only bumps on a breaking change.

## Unreleased — `preview` watch parity: marks, goto, incremental reload, browser edit (additive)

`surfaceVersion` stays **1.0**. The `preview` verb gains four watch-style actions and the live page
gains an interactive layer; **two MCP tools** are added (`preview_mark`, `preview_goto`), so the tool
count goes **17 → 19** (the `preview` verb now carries open/selection/mark/goto; the MCP **verb** set
stays 17). The whole tool surface still fits the 3,500-token budget.

### Added — preview

- **Marks** — `aioffice preview mark <file> <path> [--color --note --find --tofix]`, `preview unmark <file>
  <path>|--all`, `preview marks <file>` (and the `preview_mark` MCP tool). Advisory, in-memory highlights
  pushed to every live viewer over SSE — the "agent highlights → human reviews" loop. `path` may be the
  pseudo-path `selected` to mark everything the human clicked. Marks never touch the document.
- **Goto** — `aioffice preview goto <file> <path>` (and `preview_goto`) scrolls every live viewer to an
  element and flashes it. Server route `POST /goto` → SSE `scroll`.
- **Incremental reload** — on a disk change the page now soft-patches: it fetches `GET /content` and swaps
  only the changed `data-aio-path` nodes (full `<main>` swap only on structural change) instead of
  `location.reload()`, so scroll, selection, marks and the SSE stream survive.
- **Browser editing** — box-drag rubber-band multi-select, and double-click inline edit of an xlsx cell or
  docx paragraph that POSTs `/api/edit` (runs the format handler in-process, then live-reloads). Chart-drag
  repositioning is intentionally deferred.
- New server routes: `GET /content`, `GET|POST|DELETE /marks`, `POST /goto`, `POST /api/edit`; new SSE
  events `marks` and `scroll`. New CLI flags `--color/--note/--find/--tofix/--all`.

This closes the gap with OfficeCLI's `watch` command (marks/goto/incremental reload/in-browser editing);
aioffice's filesystem-watch auto-reload already covered — and exceeded — `watch`'s command-driven refresh.

## Unreleased — `aioffice plugin install`: self-install into AI coding hosts (additive)

`surfaceVersion` stays **1.0** and the **MCP surface is unchanged — still 17 tools**. This adds one
CLI-only verb, **`plugin`** (now **19 CLI verbs**), a host-install utility that edits local host config
files; it has no MCP tool and is outside the AI/document contract, so `office_schema` / `tools/list` are
byte-identical (guarded by `SchemaConsistencyTests`).

### Added — distribution

- **`aioffice plugin <install|uninstall|list|status>`** registers aioffice into the AI coding hosts on the
  machine without hand-editing any config:
  - **Claude Code** — `mcpServers.aioffice` in `~/.claude.json` (or `<workspace>/.mcp.json` at
    `--scope project`), the skill at `~/.claude/skills/aioffice/SKILL.md`, and the `/aioffice` command.
  - **Codex** — a self-headed `[mcp_servers.aioffice]` block appended to `~/.codex/config.toml` (backed up
    first), a skill under `~/.codex/skills/`, and a delimited section in `~/.codex/AGENTS.md`.
  - **opencode** — `mcp.aioffice` (`type:"local"`, command array) in `opencode.json`, an `agents/aioffice.md`
    subagent, a `command/aioffice.md`, and an `AGENTS.md` section.
  - **TonoBraid** — a `cc`-format plugin at `~/.tonoagent/plugins/aioffice/` (manifest + skill + command +
    bundled `.mcp.json`), registered in `~/.tonoagent/plugins.json` with `trust:"trusted"` and a sha256 tree
    digest computed byte-for-byte the way TonoBraid recomputes it on load (so trust never silently downgrades).
  - **`--host claude|codex|opencode|tonobraid|all`** (default: all *detected* hosts; comma-separate for
    several), **`--scope user|project`**, **`--dry-run`** (prints the exact files/keys it would change and
    writes nothing), **`--force`** (overwrite instead of skip). Every write is idempotent and only ever
    touches the `aioffice` key; `uninstall` removes only aioffice and leaves your other servers intact.
  - The embedded payload is the design / anti-homogenization **agent guide** (promoted into the source tree
    at `src/AIOffice.Cli/Plugin/`), version-stamped at install time and adapted per host (Claude/Codex skill
    frontmatter, opencode subagent, TonoBraid plugin skill, AGENTS.md pointer sections).
  - The guide is upgraded with methodology borrowed from the design literature: a **spec-lock + re-read-per-slide**
    discipline (so a long deck doesn't drift), a **page-rhythm** rule (anchor/dense/breathing — kills the
    "every slide is a card grid" default), presentation **modes** (pyramid/narrative/instructional/showcase/briefing),
    a **diagram-archetype → native-primitive** map (SmartArt/connectors/shapes instead of fake card grids), and a
    "reach for the full surface" section that names aioffice's underused power (SmartArt, connectors, conditional-format
    kinds, table styles, animations, embedded fonts, equations, gradient/image fills, `seriesColors`). Skill-directory
    hosts (Claude, Codex, TonoBraid) also receive two reference files — **`palette-library.md`** (10 palettes as
    usage-contracts with hex + temperament) and **`diagram-archetypes.md`** (copy-paste op skeletons for
    process/cycle/matrix/funnel/pyramid/timeline/comparison, using only real shape presets).

See [docs/MCP-SETUP.md](docs/MCP-SETUP.md#the-easy-way--aioffice-plugin-install) for the full recipe.

## 1.24.0 — complete the families: pptx inner-shadow · xlsx tab color · docx underline styles (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — each pick fills a family/vocabulary gap, byte-stable
on legacy paths. Agents that target 1.23.0 are byte-for-byte compatible with 1.24.0. (With pptx
inner-shadow the shape `a:effectLst` effect family — shadow/glow/reflection/outline/softEdge/innerShadow —
is now complete.)

### Added — pptx

- **Shape inner-shadow effect**: `innerShadow` on a shape `set`/`add` writes `a:innerShdw` (a hex/`true`,
  or `{color?, blur?, dist?, dir?}`; `false` clears) — the last slot in the effect family. Read back on
  `get`, discriminated (string vs object) like `outline`.

### Added — xlsx

- **Worksheet tab color**: `tabColor` on the sheet-level `set` writes `sheetPr/tabColor` (a hex; `''`
  clears), reported on `get` when set.

### Added — docx

- **Underline styles**: the `underline` prop widens from `bool` to also accept a style name
  (`double`/`thick`/`dotted`/`dash`/`wave`/…) on run and style-definition set. The bool form stays
  byte-identical; `get` reads a bool for single/none and the style string otherwise.

## 1.23.1 — patch: wire the CSV export delimiter through the CLI

`surfaceVersion` stays **1.0**; byte-for-byte compatible with 1.23.0.

### Fixed

- **`read --view csv --delimiter` now works on the CLI**: 1.23.0 shipped the CSV export delimiter on the
  handler + MCP surface (`office_read {view:"csv", delimiter:"tab"}`), but the CLI `read` verb never
  forwarded `--delimiter` to the handler context, so `aioffice read … --view csv --delimiter tab` silently
  emitted comma. The CLI now threads `--delimiter` through (and declares the option). Added an end-to-end
  `FileVerbs`-level regression test — the handler-level tests could not catch a missing CLI forward.

## 1.23.0 — complete the half-shipped: docx hyperlink ScreenTip · xlsx CSV export delimiter · pptx shape soft-edge (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — each format gains the one missing facet of an
already-shipped feature, byte-stable on legacy paths. Agents that target 1.22.0 are byte-for-byte
compatible with 1.23.0.

### Added — docx

- **Hyperlink ScreenTip**: `add type:link` accepts an optional `tooltip` → `w:hyperlink/@w:tooltip`
  (external + internal links), reported on `get` when present. Mirrors the shipped xlsx cell tooltip.

### Added — xlsx

- **CSV export delimiter**: `read --view csv` accepts an optional `delimiter` (`comma|semicolon|tab`,
  the same vocabulary CSV import already parses) — export was previously comma-only. Delimiter-aware
  quoting; omitting it is byte-identical.

### Added — pptx

- **Shape soft-edge effect**: `softEdge` on a shape `set`/`add` writes `a:softEdge` (`true` = 2.5pt,
  a size like `"5pt"`, `false` clears) — the last slot in the shadow/glow/reflection/outline roster,
  read back on `get`.

### Fixed

- **Honesty**: the CSV help topics advertised a `|` pipe delimiter the parser never accepted; the docs
  now match the real `comma|semicolon|tab` vocabulary.

## 1.22.0 — finish the vocabulary: xlsx CF bundle · docx content-control lock · pptx adjust handles (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — three vocabulary completions on existing ops,
byte-stable on legacy paths. Agents that target 1.21.0 are byte-for-byte compatible with 1.22.0.

### Added — xlsx

- **CF vocabulary bundle**: the `cellIs` operator `notBetween` (two formulas, `value2` required) plus
  three text-match kinds `notContainsText` / `startsWith` / `endsWith` (`{text, fill?, color?, bold?}`,
  siblings of `containsText`). Surface tokens stay stable (`startsWith` serializes as OOXML
  `beginsWith`); all four round-trip natively.

### Added — docx

- **Content-control `lock`**: `sdtLocked|contentLocked|sdtContentLocked|unlocked` →
  `w:sdtPr/w:lock`, reported on `get` when present. Honest semantics: the lock affects Word's UI only —
  AIOffice can still edit/remove a locked control (pinned by tests).

### Added — pptx

- **Preset-shape adjust handles**: `adjust` on `roundRect` (corner radius), `arrow` (`{adj1?, adj2?}`),
  and `triangle` (apex) writes `a:avLst/a:gd` guides (raw ECMA units 0–100000), read back on `get`
  (incl. group children and foreign decks). Non-adjustable presets reject with candidates. The SVG
  render approximation does not reflect adjusts (OOXML is authoritative).

## 1.21.0 — completing the pairs: pptx cell gradient/image fill · docx section vAlign · xlsx duplicate/unique CF (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — each format gains the missing counterpart of a
surface it already ships, byte-stable on its legacy path. Agents that target 1.20.0 are byte-for-byte
compatible with 1.21.0.

### Added — pptx

- **Table-cell gradient & image fills**: the cell `fill` prop widens from a bare hex to
  `hex | {gradient:{…}} | {image:{…}}` (the same vocabulary/builders as shapes and slide backgrounds),
  with the object read back on `get` and rendered as its start colour in SVG. Border-edge/fill child
  order holds; replacing a fill prunes orphaned image parts.

### Added — docx

- **Section vertical page alignment**: `verticalAlign` (`top|center|justify|bottom`) →
  `w:sectPr/w:vAlign` (`justify` ↔ OOXML `both`, symmetric). Reported on `get` only when present.

### Added — xlsx

- **duplicateValues / uniqueValues conditional formats**: two new kinds with the shared
  `fill`/`color`/`bold` styling, full round-trip. Also corrects the read-side spelling for
  Excel-authored rules of these types (previously projected as `isDuplicate`/`isUnique` via the
  lowercase fallback; now the OOXML-aligned `duplicateValues`/`uniqueValues`).

## 1.20.0 — read/write parity & customization: xlsx iconSet thresholds · docx cell textDirection · pptx cell formatting read-back (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — each closes a customization or read/write-parity gap
on an existing cell-or-rule op, byte-stable on its legacy path. Agents that target 1.19.0 are byte-for-byte
compatible with 1.20.0.

### Added — xlsx

- **iconSet custom per-icon thresholds**: an `iconSet` conditional format now accepts `thresholds`
  (`[{type:percent|num|percentile|formula, value}]`, N entries for an N-icon set), replacing the hardcoded
  even split. Omitting it is byte-identical to 1.19; `get` reports them only for a non-default set.

### Added — docx

- **Table-cell `textDirection`**: a table cell `set`/`get` accepts `textDirection` (`lrTb|tbRl|btLr`, +
  rotated variants) → `w:tcPr/w:textDirection`, at parity with pptx. `get` reports it (null when absent).

### Added — pptx

- **Table-cell run formatting on `get`**: `get` now reports the `bold`/`color`/`fontSize`/`align` a cell
  already writes (previously write-only) — a read/write-parity fix. `color` is bare hex (matching `fill`);
  inherited/theme values and absent props project `null`. Bare cells emit no new keys (byte-identical).

## 1.19.0 — round-trip the writable, evaluate the proven gap: pptx shape fill object · xlsx AVERAGEIF · docx content-control data binding (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — each closes a write/read asymmetry or one
probe-confirmed function gap, byte-stable on its legacy path. Agents that target 1.18.0 are byte-for-byte
compatible with 1.19.0.

### Added — pptx

- **Full gradient/image shape fill object on `get`**: a shape's `fill` now reads back as the gradient
  (`{type, angle?, stops}`) or image (`{src, mode, tint?}`) object — the read-side inverse of the shape
  fill the contract already writes (mirrors v1.18's background read-back), for body shapes, group children,
  and master/layout shapes. Solid stays a bare hex; no fill stays `null`. Read-only change.

### Added — xlsx

- **AVERAGEIF write-time evaluation**: `AVERAGEIF` (the one criteria-aggregate the base engine left as
  `#NAME?`; its twin `AVERAGEIFS` was already evaluated, and `SUMIF`/`COUNTIF` evaluate natively) now
  computes and caches its value. Supports the 2-arg (`range, criteria`) and 3-arg (`crit_range, criteria,
  avg_range`) forms; no match → `#DIV/0!`.

### Added — docx

- **Content-control external XML data binding**: `add type:contentControl` accepts an optional
  `dataBinding` (`{xpath, storeItemId?, prefixMappings?}`) writing a `w:dataBinding` under `sdtPr`, on all
  control kinds. `get` reports it when present; a control without it is byte-identical. Missing/empty
  `xpath` → `invalid_args`.

## 1.18.0 — round-trip & prop completions: pptx bg object · xlsx colorScale midpoint · docx tracked merge/if fields (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — each closes a write/read asymmetry, byte-stable on
its legacy path. Agents that target 1.17.0 are byte-for-byte compatible with 1.18.0.

### Added — pptx

- **Full gradient/image background object on `get /slide`**: `get` now projects the gradient
  (`{type, angle?, stops}`) or image (`{src, mode, tint?}`) object — completing the v1.14 write round-trip
  (previously `get` only reported a solid hex / the `backgroundKind` string). Solid stays a bare hex, none
  stays `null`. Read-only change; no write path touched.

### Added — xlsx

- **colorScale custom 3-color midpoint**: a 3-color `colorScale` now accepts `midType`
  (`num`/`percent`/`percentile`) + `midValue`, replacing the hardcoded percentile-50. Omitting both is
  byte-identical to 1.17; `get` reports them only for a custom midpoint. Invalid combos (midType without
  midColor, midValue without midType, unknown midType, percent/percentile out of 0–100) → `invalid_args`.

### Added — docx

- **Tracked mergeField / ifField insertion (append path)**: `add type:mergeField`/`type:ifField` now work
  under `track:true` on the append form, wrapping the complex-field runs in `w:ins` (`kind:"insert"`;
  accept keeps, reject removes). The `find` mid-paragraph form, `type:field` (w:fldSimple element), and
  tracked link/toc/tableOfFigures/index/indexEntry stay unsupported.

## 1.17.0 — table & tracked depth: pptx cell borders · docx tracked cross-refs · xlsx totals editing (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — each adds a prop or relaxes a guard on an existing
op, byte-stable on its legacy path. Agents that target 1.16.0 are byte-for-byte compatible with 1.17.0.

### Added — pptx

- **Table-cell per-edge borders**: `set` on a table cell accepts `borders` (`{top/bottom/left/right/all}`,
  each `{color?, widthPt?, style?}`, or `"none"` to clear), writing `a:lnL/lnR/lnT/lnB` at the front of
  `a:tcPr` (style ∈ single/double/dotted/dashed/none; widthPt → EMU at 12700/pt). The preset light/medium/
  dark look is byte-stable; `get` reports the per-edge borders.

### Added — docx

- **Tracked cross-reference insertion**: `add type:crossRef` now works under `track:true` (previously
  `unsupported_feature`), wrapping only the new complex-field runs in `w:ins`. `read --view revisions`
  reports `kind:"insert"`; accept keeps the REF field, reject removes it. Tracked link/field/caption stay
  unsupported (CT_Ins constraints).

### Added — xlsx

- **Table totals-row editing**: `set` on a table accepts `totals` (column → `{function?, label?}`),
  relaxing the table-set guard for totals-only sets. Function and label are mutually exclusive per column
  (Excel's model; label wins when both given); `function:"none"` / empty label clear. `get` reports
  `totalsFunction`/`totalsLabel`. Non-string or empty settings → `invalid_args`.

## 1.16.0 — completion & parity: docx tracked notes · xlsx AGGREGATE EXC/mode · pptx text-frame anchoring (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — three explicitly-deferred threads finished, one
per format, each byte-stable on its legacy path. Agents that target 1.15.0 are byte-for-byte compatible
with 1.16.0.

### Added — docx

- **Tracked footnote / endnote insertion**: `add type:footnote`/`type:endnote` now works under
  `track:true` (previously `unsupported_feature`), completing v1.15's tracked authoring. Only the
  inserted reference run is wrapped in `w:ins`; `read --view revisions` reports `kind:"insert"`; accept
  keeps the note, reject removes the reference. Untracked add is unchanged; tracked link/field inserts
  stay unsupported.

### Added — xlsx

- **AGGREGATE 13 / 18 / 19**: `MODE.SNGL`, `PERCENTILE.EXC`, `QUARTILE.EXC` now evaluate at write time
  (previously `formula_not_evaluated`), completing v1.11's headless eval. PERCENTILE.EXC endpoints are
  inclusive like Excel; QUARTILE.EXC accepts `q∈{1,2,3}` (0/4 → `#NUM!`). The `options` error-handling
  applies; only the array form stays deferred.

### Added — pptx

- **Text-frame anchoring on shapes**: a text shape's `set`/`add` accept `vAlign` (top/middle/bottom →
  `a:bodyPr @anchor`), `textDirection` (horizontal/vertical/vertical270 → `@vert`), and
  `marginLeft/Right/Top/Bottom` (→ EMU insets), at parity with table cells. `get` projects them as
  nullable fields (incl. group children). These bodyPr attributes don't disturb autofit; invalid
  tokens/over-range margins → `invalid_args`.

## 1.15.0 — styling depth + tracked authoring: pptx outline styling · xlsx data-bar thresholds · docx tracked formatting (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — three existing props are widened in the value
shapes they accept, each legacy form byte-stable. Agents that target 1.14.0 are byte-for-byte
compatible with 1.15.0.

### Added — pptx

- **Shape outline styling depth**: the `outline` prop now also accepts `{color?, width?, dash?,
  compound?}` (writing `a:ln @w` + `a:prstDash @val` + `@cmpd`) beyond the bare hex string that always
  drew a 1pt solid line. `dash` ∈ solid/dash/dot/dashDot/dashDotDot/lgDash/lgDashDot/lgDashDotDot;
  `compound` ∈ single/double/thickThin/thinThick/triple. The bare-string/`false` form is unchanged;
  `get` reports the object when a non-default width/dash/compound is present, a bare hex otherwise.

### Added — xlsx

- **Data-bar thresholds + show-value**: the `dataBar` conditional format now accepts `minType`/`maxType`
  (auto/fixed/percent/percentile/formula), `minValue`/`maxValue`, and `showValue`, replacing the
  always-auto lowest→highest scaling. Omitting them is byte-identical to 1.14. Thresholds are authored
  into the rule's `cfvo` (+ x14 twin) via a post-save pass and survive the workbook fix-up; `get`
  reports them.

### Added — docx

- **Tracked formatting-change authoring**: a tracked `set` (`track:true`) on a body paragraph/run now
  produces `w:rPrChange` (run formatting: bold/italic/underline/color/fontSize/…) and `w:pPrChange`
  (paragraph style/props) instead of returning `unsupported_feature`. `text`+formatting in one op makes
  `w:del`+`w:ins` plus a `w:rPrChange` on the inserted run. The changes read back as `kind:"format"` and
  the existing accept/reject resolve them. Text-only tracked edits, header/table-cell scope, and
  untracked sets are unchanged.

## 1.14.0 — default → hand-made: docx cell borders · xlsx pivot grand-totals · pptx slide backgrounds (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — three "looks default vs hand-made" gaps closed,
one per format, each a new value on an existing op. Agents that target 1.13.0 are byte-for-byte
compatible with 1.14.0.

### Added — docx

- **Per-cell table borders** via `set` on a `tc`: `borders` takes any subset of `top`/`bottom`/`left`/
  `right`/`insideH`/`insideV`/`all`, each `{color?, widthPt?, style?}` (or `"none"` to clear an edge).
  `style` ∈ `single|double|thick|dashed|dotted|wave|none`. Writes `w:tcBorders` as a cell-level override
  that coexists with table-level `w:tblBorders`; `get` reports it. Previously any unknown cell prop
  errored — beyond the table-wide `all`/`outer`/`none` borders there was no per-edge control.

### Added — xlsx

- **Pivot grand-total visibility** via `add` (pivot): `grandTotals` accepts `"both"`/`"rows"`/
  `"columns"`/`"none"` or `{rows, columns}` booleans, written to `rowGrandTotals`/`colGrandTotals`.
  Omitting it keeps the always-both default byte-identical. `get`/describe report `grandTotals:{rows,
  columns}`; the flags survive the post-save raw pass.

### Added — pptx

- **Per-slide gradient & image backgrounds**: the slide `background` prop now also accepts
  `{gradient:{type?, angle?, stops:[{color, at}]}}` or `{image:{src}}` (deck-wide via `set /` /
  `set /master[1]` / `set /layout[i]` too), reusing the shape-fill builders. Replacing a background
  prunes the prior image's media part (no orphan). `get` reports `backgroundKind`. The solid-hex path
  is byte-stable; a gradient/image *string* or an array background still returns `unsupported_feature`.

## 1.13.0 — document finishing: docx protection · xlsx pivot show-values-as · pptx footers (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — three half-implemented capabilities
finished with additive props. Agents that target 1.12.0 are byte-for-byte compatible with 1.13.0.

### Fixed

- **xlsx pivot `showAs` is now applied** — it was previously accepted but silently ignored
  (the pivot still showed raw sums). An unknown `showAs` is now rejected instead of ignored.

### Added — docx

- **Document protection** via `set /`: `protection {edit: readOnly|comments|trackedChanges|
  forms|none, enforce?}` and `readOnlyRecommended` (enforcement-flag protection; password /
  strong encryption stays out of scope). `get /` reports it. (`set /` on a docx previously
  errored; a `set /` with non-protection props now returns `invalid_args` instead of
  `unsupported_feature`.)

### Added — xlsx

- **Pivot show-values-as**: `showAs` on a value field — `percentOfTotal`/`percentOfColumn`/
  `percentOfRow`, `runningTotal`, `differenceFrom`/`percentDifferenceFrom`/`percentOf`
  (`baseField`/`baseItem`), `index`. Excel computes the displayed % on open from the written
  `showDataAs`; `percentOfParentTotal`/`rank*` are rejected (not expressible) rather than ignored.

### Added — pptx

- **Footer / slide-number / date placeholders**: `set /slide[i] {footer, slideNumber, date}`
  per slide, or a deck-wide `set /` (optional `skipTitle`) to number a whole deck in one op.

## 1.12.0 — completion depth: edit-chart-in-place · more docx fields · xlsx autofilter criteria (additive)

`surfaceVersion` stays **1.0**; no new op/verb/tool — three half-implemented capabilities
completed with additive props/kinds. Agents that target 1.11.0 are byte-for-byte compatible
with 1.12.0.

### Added — pptx

- **Edit an existing chart in place**: `set /slide[i]/chart[k]` now accepts `title` (or `false`
  to remove), `categories`, and `series` (`{name?, values}`) — rewriting both the chart caches
  and the embedded Edit-Data workbook, instead of requiring remove-and-re-add.

### Added — docx

- **More `add type:field` kinds**: `fileName`, `numWords`, `numChars`, `author`, `createDate`,
  `saveDate`, `printDate`, `ref` (to a bookmark; text/page/aboveBelow), `hyperlink`, `fillIn` —
  each with a headless cached value (e.g. NUMWORDS = real body word count).

### Added — xlsx

- **AutoFilter criteria**: the sheet/table `autoFilter` prop now takes `{column, values:[…]}`
  or `{column, criteria:">100"|"*text*"}` (or an array per column), applying a real filter that
  hides non-matching rows; `get` reports the active filters. The bool form is unchanged.

## 1.11.0 — xlsx headless function evaluation (additive)

`surfaceVersion` stays **1.0**; no new op/prop/verb/tool — the write-time formula evaluator
now computes + caches more functions instead of firing `formula_not_evaluated`. Agents that
target 1.10.0 are byte-for-byte compatible with 1.11.0.

### Fixed

- **`SMALL` now evaluates** — it returned `#NAME?` while its twin `LARGE` worked. The
  reported bug.

### Added — xlsx write-time evaluation

- `RANK`/`RANK.EQ`, `PERCENTILE`/`PERCENTILE.INC`, `QUARTILE`/`QUARTILE.INC`, `CHOOSE`,
  `OFFSET`, `INDIRECT`, and `AGGREGATE` (function numbers 1–12 and 14–17, with the `options`
  argument honored — 2/3/6/7 ignore error cells, 0/1/4/5 propagate them) now compute and
  cache a real value. The `.EXC` AGGREGATE variants (18/19), `MODE.SNGL` (13), and
  `OFFSET`/`INDIRECT` used as a multi-cell range argument stay honestly unevaluated.
  `HLOOKUP` was already evaluated natively and is unchanged.

## 1.10.0 — docx typography primitives (additive)

`surfaceVersion` stays **1.0**; **18 CLI verbs / 17 MCP tools** unchanged. Additive only —
new optional docx run + paragraph formatting props flowing through the existing set/get op
dispatch (no MCP tool-schema change). Agents that target the 1.9.0 surface are byte-for-byte
compatible with 1.10.0.

### Added — docx

- **Run typography** (`set /body/p[i]/run[j]`, also fans out from `set /body/p[i]`):
  `highlight` (named Word highlight color), `strike`/`doubleStrike`, `smallCaps`/`allCaps`,
  `superscript`/`subscript`, `characterSpacing` (points).
- **Paragraph typography** (`set /body/p[i]`): `lineSpacing` (multiple, or
  `{atLeast|exactly: points}`), `keepNext`, `keepLines`, `pageBreakBefore`, `widowControl`,
  `outlineLevel` (0–9), `tabStops` (array of `{pos:cm, align?, leader?}`). An agent can now
  produce double-spaced reports, highlighted/struck text, super/subscripts, keep-with-next
  headings, page-break-before sections, and dot-leader tabs.

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
