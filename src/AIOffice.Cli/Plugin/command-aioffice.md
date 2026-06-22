You are producing a polished Office document (.docx / .xlsx / .pptx) with the **aioffice** tools — the MCP `office_*` tools if registered, else the `aioffice` CLI via Bash. aioffice reads/edits/renders real OOXML; never hand-roll OOXML or reach for python-pptx / pptxgenjs / SheetJS.

Follow the aioffice method — do **not** skip straight to `create` + bullets (that ships the crude, default-white "简陋" result):

1. **Brief.** Confirm the audience, the brand (colors / logo / website if any), and the ONE message. Derive everything from this.
2. **Direction.** Pick a non-default visual direction — a palette (bg · panel · ink · muted · ≤3 accents), a type pairing, and varied layout archetypes. If the user gave brand assets, derive from those; never reuse the last document's look by reflex. Run `office_help themes masters chart-polish properties-pptx` (or `aioffice help <topic>`) to load the design props — they are NOT in the terse tool schema.
3. **Build.** Set theme/master, position shapes with fills, set `fontSize`+`color` on every text box, give every chart `dataLabels`+`legend`+`axisTitles`.
4. **Render → LOOK → fix.** Render with `--engine soffice` (or `auto`) for true fidelity, actually look at the output, and fix overflow/overlap — there is no auto-wrap, so one fix pass is always needed.
5. **Gate.** `validate` must be 0 and `audit` clean (alt text, titles, font ≥12pt, contrast).

Read the bundled aioffice skill / AGENT-GUIDE for the palette menu and per-format mechanics before you start.

Brief: $ARGUMENTS
