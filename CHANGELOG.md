# Changelog

All notable changes to AIOffice are recorded here. The package **Version** follows
semantic versioning; the AI-facing **`surfaceVersion`** (the frozen contract in
[CONTRACT.md](CONTRACT.md)) moves independently and only bumps on a breaking change.

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
