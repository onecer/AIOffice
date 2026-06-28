# AIOffice — make Office files that look *designed* (and varied), not default-white

> A fresh aioffice file is **blank and unstyled**, and the design props are **not** in the
> terse MCP schema. Title + bullets + a bare chart = the plain "简陋" result. Everything that
> makes it look good is opt-in. This guide teaches the **method**, a **lock-it-once contract**
> so a long deck stays one deck, and a **menu** of directions so output is polished *and*
> different each time — not default-white, not one house style.

## Driving aioffice
One self-contained binary reads/edits/renders real `.docx/.xlsx/.pptx`. Two equivalent ways:
- **MCP tools**: `office_create office_edit office_read office_get office_query office_render office_validate office_audit office_help office_schema office_status`. Some hosts **namespace** MCP tools — they may appear as `mcp__aioffice__office_create` or `mcp__aioffice.aioffice__office_create`. Call whichever `*office_create` / `*office_edit` the host exposes; match by the `office_*` suffix, not the exact bare name.
- **CLI** via Bash: `aioffice <verb> <file> …` (on PATH after install).

Every call prints one JSON envelope `{ok, data, error{code,message,suggestion,candidates?}, meta}`; on failure read `suggestion`/`candidates`.

**Producing a `.pptx/.docx/.xlsx` means producing that actual file with aioffice.** Do NOT satisfy the request by writing an `.html` "slides" page, a Markdown outline, or hand-rolled OOXML, and do NOT silently substitute python-pptx/pptxgenjs/SheetJS. If the aioffice MCP tools aren't reachable in your host, fall back to the `aioffice` **CLI via Bash**; if neither is available, say so plainly — never hand back a webpage and call it a deck.

## RULE 0 — never ship default-white
A bare `create` + title + bullets + `kind`-only chart = crude. You MUST set a theme/master, position shapes with fills, set `fontSize`+`color` on every text box, and give every chart `dataLabels`+`legend`+`axisTitles`. Polish is opt-in.

## RULE 1 — don't ship the *same look* every time
A finance board deck, a dev-tool launch, and a wedding invite should look nothing alike. Choose a **direction per document** from the brief/brand/audience/topic. The menus below are starting points, **not** a house style. If the user gave brand colors/logo/website, **derive from those first**. Never reuse the last document's palette by reflex.

## RULE 2 — lock it once, then don't drift
On a deck >~5 slides the palette/fonts decay mid-deck (you forget slide 2's accent by slide 20). Cure: after the brief, write a tiny **`aioffice-spec.md`** in the workspace —
```
palette: bg #.. panel #.. ink #.. muted #.. accents #.. #.. #..
fonts:   major "..", minor ".."
mode:    pyramid | narrative | instructional | showcase | briefing
treatment: flat-vector | swiss | glass | isometric | editorial | dark-cinematic
rhythm:  S1 anchor · S2 dense · S3 breathing · S4 dense · …
```
**Re-read this file before editing each slide.** All colors/fonts come from it — none from memory. This is what makes 30 slides feel like one deck.

## RULE 3 — vary density per slide (kill the card-grid default)
Tag every slide one of three; the spec's `rhythm` line is the record:
- **anchor** — cover / section / closing. Deliver impact, never a default centered title+subtitle or a generic "Thank you."
- **dense** — info-heavy. Card grids, KPI dashboards, tables, charts. The baseline.
- **breathing** — low-density. **No multi-card grid** (no 3-card row, no 4-KPI grid, no 2×2 of cards). Use one giant number + one line, a hero quote, a full-bleed image with a floating caption, or naked text. Rhythm follows the narrative, not a quota — don't pad with filler "Thank you" slides.

## The design props are behind `help` — load them first
The MCP schema is terse; the per-element design vocabulary lives in help topics. Load the ones your job needs:
```
office_help chart-polish smartart connectors animations table-styles
office_help themes masters conditional-format number-formats embedded-fonts equations layouts
office_help properties-pptx properties-xlsx properties-docx scenarios
```
(CLI: `aioffice help <topic>`; `aioffice help` lists ~42; `aioffice schema` dumps the surface.) Deeper menus ship alongside this guide as **`palette-library.md`** and **`diagram-archetypes.md`** (read them when present).

## Workflow (every visual doc)
**brief → confirm → spec-lock → help → design → edit → render → LOOK → fix → validate(0) → audit.**
- **brief/confirm**: audience · brand · the ONE message · canvas · **mode** · palette · type pairing · asset plan. Present ≥2-3 real options for palette/mode (with a real-world analogy — "like an Economist spread" / "like a Stripe keynote"), not one silent default. Confirm, then write `aioffice-spec.md`.
- aioffice's default `render --to png` is an approximation (no auto-wrap). For TRUE fidelity pass **`--engine soffice`** (or `--engine auto`): `aioffice render deck.pptx --to pdf --engine soffice` or `--to png --engine soffice --scope /slide[N]`. Needs LibreOffice (+poppler for png); `aioffice doctor` shows availability. Then **actually look** at the output.
- **No autofit/auto-wrap by default** — oversized text overflows. Render, look, then shrink/reposition (pptx: `autofit:"shrink"` on a text box is the direct cure). Check each slide against its declared intent: does the most-prominent element match the slide's point? Does a `breathing` slide stay un-gridded? Fix, don't oscillate.
- Gate: `validate` = 0; `audit` clean (alt text, titles, font ≥12pt, contrast ≥4.5).

## Pick a MODE — how the deck argues (lock once, ⊥ to look)
Mode shapes title voice + slide sequencing, not pixels. Any mode pairs with any treatment.
- **pyramid** — conclusion-first, MECE. **Assertion titles** ("Domestic market grows 23% YoY", not "Market Overview"). Board / exec / investor / strategy.
- **narrative** — situation→tension→resolution, ≥1 turn. Titles read as beats. Pitches / case studies / fundraising.
- **instructional** — one concept per slide, show-then-tell. "How X is computed." Training / explainers / onboarding.
- **showcase** — one dominant visual per slide, words support. Short evocative titles. Launches / keynotes / 发布会.
- **briefing** — no thesis, even weight, sectioned. **Topic titles** ("Q3 headcount by team"). Status / reference / 周报.

## Pick a PALETTE — a usage *contract*, not just colors
Pick ONE (or derive from brand). A palette is **proportions + roles**: 1 bg (~60% breathing) + 1 panel + 1 ink + 1 muted + ≤3 accents (accent <8% — one bar + one dot). No rainbow. Full menu + temperaments in `palette-library.md`.

| Direction | bg | panel | ink | muted | accents |
|---|---|---|---|---|---|
| Cool corporate (light) | `FFFFFF` | `F4F8FD` | `0F2742` | `64748B` | `2563EB` `0EA5E9` `059669` |
| Warm editorial (light) | `FBF7EF` | `FFFFFF` | `1C1917` | `78716C` | `C2410C` `B45309` `1D4ED8` |
| Midnight tech (dark) | `0B1220` | `16213A` | `F8FAFC` | `94A3B8` | `38BDF8` `818CF8` `34D399` |
| Dark-cinematic (premium) | `0A0E27` | `1E293B` | `F1F5F9` | `94A3B8` | `14B8A6` `D4AF37` |
| Bold mono (high-contrast) | `111111` | `1C1C1C` | `FFFFFF` | `A3A3A3` | `FACC15` |
| Soft product (light) | `F8FAFC` | `FFFFFF` | `0F172A` | `64748B` | `6366F1` `EC4899` `14B8A6` |
| Jewel-tone (luxury) | `FFFDF7` | `FFFFFF` | `1A1A1A` | `78716C` | `047857` `D4AF37` `7C3AED` |
| Frost-ice (medical/SaaS) | `F8FBFD` | `FFFFFF` | `0F2742` | `64748B` | `3B82F6` `0EA5E9` |
| Nature-organic (wellness) | `F7F4EC` | `FFFFFF` | `14310F` | `6B7280` | `166534` `D4AF37` `84CC16` |
| **Brand-derived** | *a dark/light from brand* | *a tint of it* | *near-black/white* | *desaturated* | *1–3 brand accents* |

**Typography** — vary the pairing to the tone (master `majorFont`/`minorFont` on pptx, `/theme {majorFont,minorFont}` on docx): neutral sans · editorial (serif display + sans body) · modern geometric (Space Grotesk/Sora display + sans body) · technical (a mono accent for labels/code). For CJK, pair a CJK family with a Latin family explicitly.

## Pick a TREATMENT — native visual style (the depth/surface axis)
How shapes/fills carry depth. Lock one; it decides your fill + shadow vocabulary.
- **flat-vector** (default safe) — flat fills, 1.5px strokes as accent, ≤8% soft shadow on 2-3 floating elements only.
- **swiss** — aggressive whitespace, hairline rules, **no shadow**, color = zone. Consulting / luxury / type-forward.
- **glass** — translucent light-tint panels over a `gradient` bg, thin top-highlight stroke, soft `shadow`. SaaS / fintech.
- **isometric** — `gradient` faces (a ~15% darker shadowed face), shapes from preset/freeform geometry. Architecture / product structure.
- **editorial** — thin precise divider rules, paper-tint panels, intentional layering, no shadow. Finance / journalism.
- **dark-cinematic** — deep dark bg + one bright lit subject + `glow`, a gold accent line. Premium / film / launch.

> RESTRAINT (the "designed" feel comes from *absence*): pick **one** weight-tool per container — shadow OR border OR gradient OR strong tint. Stacking them = instant template look. Shadow suits ≤2-3 genuinely-floating elements per slide; keep peer-grid cards flat. Reach for typography weight, spacing, and accent bars before shadow.

## Pick LAYOUT archetypes — vary across slides (do NOT make everything a 3-card grid)
big-stat hero (one giant number) · split (text | visual) · full-bleed image/chart + overlay title · timeline/process · quote/statement · comparison (A | B) · card grid (use sparingly).

## DIAGRAM archetypes → native aioffice primitives (the underused power)
Don't hand-build a card grid for what is really a diagram. Map the shape of the idea to the right native object (op skeletons in `diagram-archetypes.md`):

| The idea is a… | Use this native aioffice primitive |
|---|---|
| **process / pipeline / steps** | `add type:smartart {layout:"process", nodes}` — or `add type:connector {kind:"elbow", endArrow}` wiring `roundRect` cards |
| **cycle / loop / flywheel (PDCA)** | `add type:smartart {layout:"cycle", nodes}` |
| **hierarchy / org chart** | `add type:smartart {layout:"hierarchy"\|"orgChart", nodes}` |
| **framework / hub + satellites** | a center shape + `connector`s to satellite shapes, or `smartart` list |
| **matrix / 2×2 (SWOT/BCG)** | 4 `roundRect` quadrants + 2 axis `line`s + axis labels |
| **funnel / pyramid (value stack)** | stacked trapezoid/triangle shapes (preset/freeform), color base→apex |
| **timeline / roadmap** | a horizontal `line` axis + milestone `ellipse` markers + dated labels (or `connector`s) |
| **comparison / before-after** | symmetric split: two panels + a divider `line`, each its own anchor |
| **equation / formula** | `add type:equation {latex, display}` (pptx + docx; LaTeX→OMML) |

## Reach for the FULL surface (these ship today and are usually skipped)
- **pptx**: **SmartArt** (`smartart`) and **connectors** (`connectors`) for real diagrams (SmartArt may render blank in some LibreOffice versions — fall back to connected shapes+connectors if so); **table styles** `add type:table {style:"medium2", firstRow:true, bandRow:true}` (fastest polished table) + per-edge cell `borders`; `gradient`/`image` fill on any shape OR slide/master/layout **background** (pptx chart series take their color from the master theme accents — set `accent1..3` on the master; `seriesColors` is xlsx-only); **animations** (entrance/emphasis/exit/motionPath, `triggerOn:"@N"`) + **transitions** (`fade` — keep restrained; an auto-firing per-element build is an "AI deck" tell, leave OFF by default); **embedded fonts** so the deck renders right off-machine; **equations**; reusable **layouts** then bind by name; `autofit:"shrink"`; deck **footer/slideNumber/date**.
- **xlsx**: all conditional-format kinds — `dataBar` · `colorScale` · `iconSet` (traffic-lights/arrows) · `topBottom` · `formula` · `aboveBelowAverage`; `gradientFill` KPI bands; `seriesColors`; numberFormat presets (`usd-millions usd-thousands percent-0|1|2 accounting-eur|gbp`); **sparkline**, **slicer**, **pivot**; print setup (`printArea fitToPage printTitles`).
- **docx**: `set /theme {accent1..6, dk1, lt2, majorFont, minorFont}`; paragraph `shading`/`border`/`spacing`/`indent` for colored & left-accent callout lead blocks; custom `style {font, next, basedOn}`; per-cell table `borders`; **captions / cross-refs**; `field` kinds (`pageNumber numPages styleRef`); **dropCap**; **page border**; **watermark** (text or picture); text effects (`shadow glow reflection outline`).

## Mechanics — apply with the direction you chose
`{bg}{panel}{ink}{muted}{accent1..3}` = the hex from your locked palette. Substitute them; don't paste one canned palette.

### PPTX (16:9 · lengths are CENTIMETERS · canvas ≈ 33.87 × 19.05 cm)
1. `edit / {slideSize:"16:9"}` → `set /master[1] {background:"{bg}", accent1:"{accent1}", accent2:"{accent2}", accent3:"{accent3}", majorFont, minorFont}` → set each slide `{background:"{bg}"}`.
2. **anchor slide**: accent bar `add type:shape {shape:"rect", h:0.15, fill:"{accent1}"}` + wordmark `{fontSize:78, bold:true, color:"{ink}"}` + tagline `{fontSize:26, color:"{muted}"}` + an accent block or full-bleed `image`/`gradient` background.
3. **dense slide**: accent bar + assertion title (`fontSize:34, bold, color:"{ink}"`) + 15pt `{muted}` deck, then the archetype — a real SmartArt diagram, or cards `add type:shape {shape:"roundRect", fill:"{panel}", x, y:6.6, w:9.4, h:9.4}` with eyebrow + bold title + `{muted}` body. `autofit:"shrink"` on text boxes.
4. **breathing slide**: one giant number `{fontSize:160, bold, color:"{accent1}"}` + one interpreting line, OR a hero quote, OR full-bleed image + floating caption. No card grid.
5. **Charts**: on a dark deck, put the chart on a light `{panel}`/white `roundRect` card (chart text has no color prop, so it needs a light backing) → chart on top → `move <card> position:"back"`. Chart carries `dataLabels:{show:"value",position:"outEnd"}, legend:"bottom", axisTitles:{category,value}, gridlines:{major:true,minor:false}`. Series colors come from the master theme accents (set `accent1..3` on `/master[1]`) — pptx charts don't take `seriesColors` (that's xlsx).
6. **Gotchas**: bare number = cm; insert manual `\n` (no auto-wrap); `title` is an `add type:slide` prop, not `set`; set `fontSize`+`color` on every text box.

### XLSX dashboard
- Canvas `{bg}`/white; title row bold `{ink}` + `{muted}` subtitle.
- **KPI band**: label row (`{accent}`-tint fill, bold) over a value row (lighter tint, bold `{ink}`); each KPI a real formula (`=SUM`/`=AVERAGE`/`=XLOOKUP(MAX(rng),rng,labels)`), formatted. `gradientFill` for a premium band.
- **Table**: header row `{ink}`/dark fill + white bold; zebra data rows; bold first column; Total band `=SUM`.
- `numberFormat` on EVERY numeric cell (`"$"#,##0` · `0.0%` · `#,##0`, or a preset).
- `add type:conditionalFormat {kind:"dataBar", color:"{accent1}"}` on % cols; `colorScale` on revenue; `iconSet` for status.
- Native charts (`bar`+`line`) with `dataLabels`/`legend`/`axisTitles`/`seriesColors`; chart-source helpers off to the side.
- `set /Sheet {printArea, fitToPage:{fitToWidth:1}}`; set column widths + row heights.

### DOCX report
- `set /theme {accent1:"{accent1}", dk1:"{ink}", lt2, majorFont, minorFont}`; `set /section[1] {pageSize:"Letter", margins ≈2.2cm}`.
- Custom paragraph styles for a masthead (Kicker / Subtitle / Lead); left-accent callouts via paragraph `border`+`shading`.
- **Tables**: GEOMETRY props — `set /body/table[1] {headerRow:true, borders:"all", borderColor, borderWidthPt:0.75, columnWidths:[…], cellPaddingCm:0.18}`; shade header `{ink}`/dark + white run; shade a totals band; total cell `{formula:"=SUM(ABOVE)"}`.
- Running header + `Page X of Y` footer via `field` (`pageNumber` + `numPages`). Set `/properties`. Optional `dropCap`, page border, `watermark`.

## ASSET BOUNDARY — what aioffice will NOT do for you
aioffice is a deterministic OOXML author. It brings **no pixels and no vectors of its own**:
- **No SVG import** (SVG is render-output only). To use an SVG icon/figure, rasterize it externally (`sips -s format png in.svg out.png`, `rsvg-convert`, ImageMagick) then `add type:image`.
- **No AI image generation, no web image search.** *You* source the file (generate/fetch/screenshot), drop it in the workspace, then embed.
- **No icon library.** Approximate icons with preset shape geometry (`rect roundRect ellipse triangle diamond arrow line`), **SmartArt**, a docx Wingdings symbol field, or an icon **PNG** you supply.
- Embeds **PNG/JPEG** raster (header-sniffed). Fonts: `.ttf`. 3D: `.glb`.
- Web/stock images that require credit get a visible caption text box.

## Worked method ≠ template
Copy the **method** (brief → spec-lock → mode+palette+treatment → diagrams-not-just-cards → render→look→fix), **not** any single palette. `aioffice help scenarios` walks a deck/dashboard/report end to end.

## Version
Installed from **aioffice {{AIOFFICE_VERSION}}**. Requires **aioffice ≥ 1.8** for the design primitives above (`aioffice version` / `office_status`). Older binaries silently drop gradient/image fills, master fonts, `seriesColors`, the number-format presets, SmartArt/connector/animation/equation/table-style props, paragraph callouts and chart-polish.
