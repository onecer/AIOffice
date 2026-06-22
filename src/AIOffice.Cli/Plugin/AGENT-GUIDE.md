# AIOffice — make Office files that look *designed* (and varied), not default-white

> Why agent output comes out crude: a fresh aioffice file is **blank and unstyled**,
> and the design props are **not** in the terse MCP tool schema. Title + bullets + a
> bare chart = the plain "简陋" result. Everything that makes it look good is opt-in.
> This guide teaches the **method**, and gives a **menu** of directions so output is
> polished *and* different each time — not default-white, and not one house style.

## Driving aioffice
One self-contained binary reads/edits/renders real `.docx/.xlsx/.pptx`. Two equivalent ways:
- **MCP tools** (if registered): `office_create`, `office_edit`, `office_read`, `office_get`, `office_query`, `office_render`, `office_validate`, `office_audit`, `office_help`, `office_schema`, `office_status`.
- **CLI** via Bash: `aioffice <verb> <file> …` (run `aioffice` directly; it is on PATH after install).

Every call prints one JSON envelope `{ok, data, error{code,message,suggestion,candidates?}, meta}`; on failure read `suggestion`/`candidates`. Use aioffice for any `.docx/.xlsx/.pptx`; don't hand-roll OOXML or use python-pptx/pptxgenjs/SheetJS.

## RULE 0 — never ship default-white
A bare `create` + title + bullets + `kind`-only chart = crude. You MUST set a theme/master, position shapes with fills, set `fontSize`+`color` on every text box, and give every chart `dataLabels`+`legend`+`axisTitles`. Polish is opt-in.

## RULE 1 — don't ship the *same look* every time either
Choose a visual **direction per document** from the brief, brand, audience and topic — a
finance board deck, a dev-tool launch, and a wedding invite should look nothing alike.
The palettes below are a **menu and starting points, not a house style**. If the user gave
brand colors / a logo / a website, **derive from those first**. Never reuse the last
document's palette by reflex.

## The design props are behind `help` — load them first
The MCP schema is terse (token budget); the per-element design props live in help topics:
```
office_help chart-polish      # dataLabels / legend / axisTitles / trendlines
office_help themes  masters   # master background + accent scheme + fonts
office_help properties-pptx  properties-xlsx  properties-docx
office_help number-formats  conditional-format  table-styles
```
(CLI: `aioffice help <topic>`; `aioffice help` lists all ~40; `aioffice schema` dumps the surface.)

## Workflow (every visual doc)
**brief → help → choose a direction → design (theme + layout) → edit (fill) → render → LOOK → fix → validate(0) → audit.**
- **brief**: who is the audience, what's the brand, what's the ONE message? Derive the direction from this.
- aioffice's default `render --to png` is an approximation (shape outlines, no auto-wrap). For TRUE fidelity pass **`--engine soffice`** (or `--engine auto`): `aioffice render deck.pptx --to pdf --engine soffice` (whole doc) or `aioffice render deck.pptx --to png --engine soffice --scope /slide[N]` (per slide). Needs LibreOffice (+ poppler/`pdftoppm` for png) — `aioffice doctor` shows availability under `renderers`. Then look at the output.
- **No autofit/auto-wrap** — oversized text overflows/overlaps. Render, look, shrink/reposition. Even a good design needs one fix pass.
- Gate: `validate` = 0; `audit` clean (alt text, titles, font ≥12pt, contrast).

## Choose a direction (do NOT default)
**Palette** — pick ONE that fits the audience/topic, or derive from the user's brand. Each = bg · panel · ink · muted · accents. Use 1 bg + 1 panel + 1 ink + 1 muted + ≤3 accents. No rainbow.

| Direction | bg | panel | ink | muted | accents |
|---|---|---|---|---|---|
| Midnight tech (dark) | `0B1220` | `16213A` | `F8FAFC` | `94A3B8` | `38BDF8` `818CF8` `34D399` |
| Warm editorial (light) | `FBF7EF` | `FFFFFF` | `1C1917` | `78716C` | `C2410C` `B45309` `1D4ED8` |
| Clean corporate (light) | `FFFFFF` | `F4F8FD` | `0F2742` | `64748B` | `2563EB` `0EA5E9` `059669` |
| Bold mono (high-contrast) | `111111` | `1C1C1C` | `FFFFFF` | `A3A3A3` | `FACC15` (single) |
| Soft product (light) | `F8FAFC` | `FFFFFF` | `0F172A` | `64748B` | `6366F1` `EC4899` `14B8A6` |
| Deep forest (dark) | `0F1F1A` | `15302A` | `ECFDF5` | `9CA3AF` | `34D399` `FBBF24` `60A5FA` |
| **Brand-derived** | *a dark from the brand* | *a tint of it* | *near-black or white* | *desaturated* | *1–3 brand accents* |

**Typography** — vary the pairing to match the tone (master theme on pptx, or `/theme {majorFont,minorFont}` on docx): neutral sans (Inter/Inter), editorial (serif display + sans body), modern (geometric display like Space Grotesk/Sora + sans body), technical (a mono accent for labels/code).

**Layout archetype** — vary across slides; do NOT make everything a 3-card grid:
card grid · big-stat hero (one giant number) · split (text | visual) · full-bleed chart/image + overlay title · timeline/process · quote/statement · comparison (A | B).

## Mechanics — apply with the direction you chose
`{bg}` `{panel}` `{ink}` `{muted}` `{accent1..3}` = the hex from your chosen palette. Substitute them; don't paste one canned palette everywhere.

### PPTX (16:9 · lengths are CENTIMETERS · canvas ≈ 33.87 × 19.05 cm)
1. `edit /  {slideSize:"16:9"}` → `set /master[1] {background:"{bg}", accent1:"{accent1}", accent2:"{accent2}", accent3:"{accent3}"}` → set each slide `{background:"{bg}"}`.
2. **Title** (or another archetype): accent bar `add type:shape {shape:"rect", h:0.15, fill:"{accent1}"}` + wordmark `{fontSize:78, bold:true, color:"{ink}"}` + tagline `{fontSize:26, color:"{muted}"}` + an accent block.
3. **Content**: accent bar + section title (`fontSize:34, bold, color:"{ink}"`) + 15pt `{muted}` subtitle, then the chosen layout — e.g. cards `add type:shape {shape:"roundRect", fill:"{panel}", x, y:6.6, w:9.4, h:9.4}` with a color-coded number/eyebrow + bold title + `{muted}` body.
4. **Charts**: add a `roundRect` card (a light `{panel}`/white) **then** the chart on top **then** `move <card> position:"back"`. Chart must carry `dataLabels:{show:"value",position:"outEnd"}, legend:"bottom", axisTitles:{category,value}, gridlines:{major:true,minor:false}`.
5. **Gotchas**: bare number = cm; insert manual `\n` (no auto-wrap); `title` is an `add type:slide` prop, not `set`; set `fontSize`+`color` on every text box.

### XLSX dashboard
- Canvas in `{bg}`/white; title row bold `{ink}` + `{muted}` subtitle.
- **KPI band**: label row (`{accent}`-tint fill, bold) over a value row (lighter tint, bold `{ink}`); each KPI a real formula (`=SUM`/`=AVERAGE`/`=XLOOKUP(MAX(rng),rng,labels)`), formatted.
- **Table**: header row `{ink}`/dark-band fill + white bold; data rows zebra-banded (two near tints); first column bold; a Total band with `=SUM`.
- `numberFormat` on EVERY numeric cell: `"$"#,##0` money · `0.0%` ratio · `#,##0` counts.
- `add type:conditionalFormat {kind:"dataBar", color:"{accent1}"}` on the % column; `{kind:"colorScale", minColor, maxColor}` on revenue.
- Native charts (`bar` + `line`) with `dataLabels`/`legend`/`axisTitles`; put 2-col chart-source helpers off to the side (e.g. `K:L`).
- `set /Sheet {printArea:"B2:J54", fitToPage:{fitToWidth:1}}`; set column widths + row heights.

### DOCX report
- `set /theme {accent1:"{accent1}", dk1:"{ink}", lt2, majorFont, minorFont}`; `set /section[1] {pageSize:"Letter", margins ≈2.2cm}`.
- Custom paragraph styles for a masthead (Kicker eyebrow / Subtitle / Lead).
- **Tables**: GEOMETRY props: `set /body/table[1] {headerRow:true, borders:"all", borderColor, borderWidthPt:0.75, columnWidths:[…], cellPaddingCm:0.18}`; shade header cells `{ink}`/dark + white run `color`; shade a totals band; total cell `{formula:"=SUM(ABOVE)", numberFormat:"integer"}`.
- Running header `add /header[1] type:header {text}` + `Page X of Y` footer via `add /footer[1]/p[1] type:field {kind:"pageNumber", leadingText:"Page "}` then `{kind:"numPages", leadingText:" of "}`. Set `/properties`.

## More design range (use these for variety)
- **pptx**: `gradient` (`{type:linear|radial, angle?, stops:[{color, at:0..100}]}`) and `image` (a path, or `{src, mode:stretch|tile, tint?}`) fill on **any shape OR the slide/master/layout background** — reach for a gradient hero/panel instead of a flat fill when a direction wants depth; `majorFont`/`minorFont` on `set /master[1]` apply your type pairing deck-wide.
- **xlsx**: `seriesColors:["RRGGBB", …]` on a chart (`add` or `set`) so dashboard charts render in **your** brand colors, not Excel's default office palette; numberFormat presets `usd-millions` `usd-thousands` `millions` `thousands-k` `percent-0|1|2` `accounting-eur|gbp`; `gradientFill` (`{type, angle?, stops:[{color, pos}]}`) on a cell/range for KPI bands.
- **docx**: paragraph `shading` / `border` (`{style,color?,widthPt?,sides?}`) / `spacingBefore` / `spacingAfter` / `indentLeft` / `indentRight` on `/body/p[i]` for colored callout & left-accent **lead blocks**; `font` on a run, a paragraph, or a custom style for a display face.

## Worked method ≠ template
Copy the **method** (theme → positioned cards → framed+labeled chart → render→look→fix), **not** any single palette — choose your own from the brief. `aioffice help scenarios` walks a deck/dashboard/report end to end.

## Version
Installed from **aioffice {{AIOFFICE_VERSION}}**. Requires **aioffice ≥ 1.8** for the design primitives above (`aioffice version` / `office_status`). Older binaries silently drop gradient/image fills, master fonts, `seriesColors`, the new number-format presets, paragraph callouts, chart-polish, theme and table-style props.
