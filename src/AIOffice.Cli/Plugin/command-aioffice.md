You are producing a polished Office document (.docx / .xlsx / .pptx) with the **aioffice** tools — the MCP `office_*` tools if registered (some hosts namespace them, e.g. `mcp__aioffice__office_create`; match by the `office_*` suffix), else the `aioffice` CLI via Bash. Producing a `.pptx/.docx/.xlsx` means producing that **actual file** with aioffice — do NOT hand back an `.html` "slides" page, a Markdown outline, or python-pptx/pptxgenjs output and call it done; if aioffice isn't reachable, say so. aioffice reads/edits/renders real OOXML and brings no images/icons/SVG of its own — you source any PNG/JPEG asset yourself (no AI-gen, no web search, no SVG import).

Do **not** skip straight to `create` + bullets (that ships the crude, default-white "简陋" result). Follow the aioffice method:

**1. Brief + confirm (one blocking step).** Before building, lock these and present ≥2-3 real options where there's a choice (each with a real-world analogy — "like an Economist feature" / "like a Stripe keynote"), never one silent default:
   - **Audience & the ONE message** + **brand** (colors / logo / website if any).
   - **Canvas**: 16:9 deck · 4:3 · A4 report · xlsx dashboard.
   - **Mode** (how it argues): pyramid (assertion titles, exec) · narrative (story arc) · instructional (one-concept-per-slide) · showcase (visual-led keynote) · briefing (topic titles, status).
   - **Palette** (bg · panel · ink · muted · ≤3 accents) — derive from brand if given; else pick from `palette-library.md`.
   - **Type pairing** + **treatment** (flat-vector / swiss / glass / isometric / editorial / dark-cinematic).
   - **Asset plan**: which images/icons are needed and how *you* will produce them (PNG/JPEG only).

**2. Spec-lock.** Write `aioffice-spec.md` (palette HEX, fonts, mode, treatment, per-slide rhythm: anchor/dense/breathing). **Re-read it before editing each slide** so a long deck stays one deck. `breathing` slides get no card grid.

**3. Load props.** Run `office_help` for the topics you need — the design vocabulary is NOT in the terse schema: `themes masters chart-polish smartart connectors animations table-styles conditional-format properties-pptx properties-xlsx properties-docx`. For diagrams, read `diagram-archetypes.md`.

**4. Build.** Set theme/master + fonts; reach for real diagrams (SmartArt / connectors — see `diagram-archetypes.md`) over hand-built card grids; set `fontSize`+`color` on every text box; give every chart `dataLabels`+`legend`+`axisTitles`+`seriesColors`. One weight-tool per container (shadow OR border OR gradient — never stacked).

**5. Render → LOOK → fix.** Render with `--engine soffice` (or `auto`) for true fidelity, **actually look**, and fix overflow/overlap (`autofit:"shrink"`; no auto-wrap, so one fix pass is always needed). Check each slide against its declared intent: does the most-prominent element match the slide's point? Does a `breathing` slide stay un-gridded?

**6. Gate.** `validate` = 0 and `audit` clean (alt text, titles, font ≥12pt, contrast ≥4.5).

Read the bundled AGENT-GUIDE (plus `palette-library.md` / `diagram-archetypes.md` when present) for the full menus before you start.

Brief: $ARGUMENTS
