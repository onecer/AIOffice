using AIOffice.Core;

namespace AIOffice.Mcp;

/// <summary>
/// The progressive documentation behind <c>office_help</c> / <c>aioffice help</c>.
/// Everything too big for tool schemas (property tables, selector grammar,
/// addressing details) lives here so the always-resident schema surface stays
/// inside its token budget.
/// </summary>
public static class HelpTopics
{
    private const string IndexTopic = "index";

    private static readonly IReadOnlyDictionary<string, (string Doc, string[] Related)> Topics =
        new Dictionary<string, (string, string[])>(StringComparer.OrdinalIgnoreCase)
        {
            [IndexTopic] = (
                """
                ## aioffice help — topics
                - addressing      path grammar for docx/xlsx/pptx (1-based)
                - selectors       CSS-like query grammar for office_query
                - edit-ops        op kinds, position forms, atomicity, expect_rev
                - envelope        the {ok,data,error,meta} result shape
                - errors          all error codes and how to recover
                - audit           accessibility + quality lint codes + --fix semantics (M7)
                - diff            semantic compare two files / a snapshot: change kinds + the workflow (M8)
                - bridges         markdown↔docx and csv↔xlsx import/export (M5)
                - equations       LaTeX → Office Math (docx + pptx): the supported subset (M6/M10)
                - embeds          embed/list/extract files as OLE objects in docx/xlsx/pptx (M10)
                - rtl             right-to-left paragraph/run/table (M6)
                - docx/paragraph  paragraph element: settable props (incl. rtl)
                - docx/table      deep tables: merges, borders, shading, widths, rtl (M5/M6)
                - docx/field      PAGE/NUMPAGES/DATE/TITLE fields (M5)
                - docx/header     default/firstPage/even header+footer variants (M5)
                - docx/revisions  tracked changes: track:true, accept/reject ops
                - docx/comment    comments + threaded replies: add/read/remove
                - docx/style      style definitions: add/set/remove, apply
                - docx/image      inline pictures: add type "image"
                - docx/caption    Figure/Table/Equation captions + cross-references (M8)
                - docx/list       numbered/bulleted/nested list items (M3)
                - docx/link       hyperlinks, bookmarks, footnotes (M3)
                - docx/section    page size/orientation/margins + columns (M3/M6)
                - docx/equation   inline/display LaTeX equations (M6)
                - pptx/equation   LaTeX equations in a slide text box (M10)
                - xlsx/cell       cell element: settable props incl. hyperlink (M5)
                - xlsx/table      Excel Tables (ListObjects) + totals + structured refs (M6)
                - xlsx/group      row/column outline grouping (M6)
                - xlsx/dataValidation  list dropdowns + wholeNumber/decimal/date/textLength rules (M5)
                - xlsx/sparkline  line/column/winLoss sparklines (M5)
                - xlsx/comment    threaded comments + replies (M5)
                - xlsx/pivot      pivot tables: add type "pivot"
                - xlsx/conditionalFormat  cellIs/colorScale/dataBar/containsText
                - xlsx/slicer     table-column / pivot-field slicers (M8)
                - xlsx/image      anchored pictures: add type "image"
                - xlsx/sheet      freeze/autoFilter/print setup, names, streaming read+write (M3/M6)
                - pptx/shape      shape element: settable props incl. preset geometry/z-order, hyperlink/action (M8)
                - pptx/table      native tables: merges, looks, columnWidths (M5)
                - pptx/animation  entrance/emphasis/exit effects + timeline reorder (M4/M5/M6)
                - pptx/comment    slide comments + threaded replies (M5)
                - pptx/slide      slide props incl. background + transition
                - pptx/section    slide sections grouping the outline (M6)
                - pptx/master     editable masters/layouts + cloned layouts (M6)
                - pptx/slideSize  slide dimensions / aspect ratio (M6)
                - pptx/notes      speaker notes: /slide[i]/notes
                - pptx/image      pictures: add type "image"
                - pptx/chart      native charts + cross-doc dataFrom (M3)
                Call office_help {topic:"<name>"} (CLI: aioffice help <name>).
                """,
                ["addressing", "selectors", "edit-ops", "bridges"]),

            ["audit"] = (
                """
                ## audit (M7) — office_audit {file, category?, severity?, fix?}
                Findings are DATA, not errors: ok:true / exit 0 even with error-severity findings (like office_validate).
                category: accessibility | quality | all (default all). severity: error|warning|info = the MINIMUM to report (default info).
                Result: {findings:[{id:"code#path", severity, category, code, path?, message, suggestion, autofixable}], summary:{errors,warnings,infos}}.
                fix:true applies only the SAFE autofixes, then re-audits — result adds {fixed:N, remaining:[ids]}. Never destructive.
                accessibility codes: a11y_no_alt_text(err,fix) a11y_no_table_header(err,fix) a11y_no_doc_title(warn,fix)
                  a11y_no_slide_title(warn,fix) a11y_heading_skip(warn) a11y_low_contrast(warn) a11y_tiny_font(warn)
                  a11y_merged_data_cells(warn) a11y_reading_order(info).
                quality codes: quality_broken_ref(err) quality_formula_error(err) quality_broken_link(err)
                  quality_empty_heading(warn) quality_off_canvas(warn) quality_empty_placeholder(warn)
                  quality_orphan_bookmark(info,fix) quality_duplicate_id(warn).
                Autofixes: placeholder alt text "(describe this image)", mark a table header row, set a doc/slide title
                  (first heading > file name > placeholder), drop an orphan bookmark. Everything else is report-only.
                """,
                ["errors", "edit-ops"]),

            ["diff"] = (
                """
                ## diff (M8) — office_diff {file, other? | snapshot?, view?}  (CLI: aioffice diff <file> [<other>] [--snapshot N] [--view …])
                Semantically compares the CURRENT document (file) against a baseline and returns a SORTED, platform-stable
                change list. Changes are DATA: ok:true / exit 0 even for a big diff (like office_validate/office_audit).
                Baseline = exactly one of:
                  other     another SAME-format file (the OLD document); a format mismatch is invalid_args naming it.
                  snapshot  snapshot N of file's own auto pre-edit ring (restored to a throwaway temp baseline); a missing
                            index is invalid_args naming the available numbers. file_snapshot {action:"list"} shows the ring.
                Result: {changes:[{kind, path, before?, after?, detail?}], summary:{added,removed,modified,moved}, baseline, view}.
                  kind     added | removed | modified | moved.  path = canonical path in the CURRENT file (the BASELINE path for removed).
                  modified before/after carry concise old/new values (text, "Normal"->"Heading1", a cell value, …); detail names what changed.
                  moved    detail = "moved from <old path>" (matched content, new position — distinguished via LCS/content-hash).
                view: detailed (default) keeps full before/after; summary trims each change to {kind, path} (counts still in summary).
                Order is canonical: sorted by (path, kind) ordinal, so the same two documents diff identically on every platform.
                Snapshot-diff workflow: edit a file (auto-snapshots the pre-image) -> office_diff {file, snapshot:1} shows exactly
                what that edit changed; great for a "what did I just do" review before committing.
                """,
                ["edit-ops", "errors"]),

            ["docx/caption"] = (
                """
                ## docx captions + cross-references (M8)
                caption:  {"op":"add","path":"/body/p[3]","type":"caption","props":{"label":"Figure","text":"Quarterly trend","position":"after"}}
                          label = Figure | Table | Equation (each has its own SEQ counter). position before|after the anchor block (default after).
                          A Caption-styled paragraph: "Figure " + a SEQ field + ": text", wrapped in a bookmark. Address it /caption[@label=Figure][1]
                          (1-based per label). The number is CACHED — Word renumbers all captions on open or field refresh (F9); a
                          caption_numbers_cached warning says so.
                crossRef: {"op":"add","path":"/body/p[5]","type":"crossRef","props":{"to":"/caption[@label=Figure][1]","show":"labelAndNumber"}}
                          appends a REF field at the paragraph pointing at the caption's bookmark, with cached display text ("Figure 1").
                          show = labelAndNumber (default) | numberOnly | text | page (page uses PAGEREF). leadingText prefixes literal text.
                get /caption[@label=Figure][1] returns label/number/text; read {view:"structure"} lists captions.
                """,
                ["docx/paragraph", "edit-ops"]),

            ["xlsx/slicer"] = (
                """
                ## xlsx slicers (M8) — add type "slicer" (a dashboard filter on a table column or a pivot field)
                table-column: {"op":"add","path":"/Sheet1/table[@name=Sales]","type":"slicer","props":{"column":"Region",
                              "caption"?,"x"?:"H2","widthCells"?,"heightCells"?}}
                pivot-field:  {"op":"add","path":"/Sheet1/pivot[1]","type":"slicer","props":{"field":"Region","caption"?,"x"?,…}}
                column/field is matched case-insensitively; a wrong name is invalid_args with the available names as candidates.
                x is the top-left anchor cell (default: two columns right of the used range); box defaults to 2x7 cells.
                Authored on raw OpenXml (ClosedXML can't create slicer parts) — validates clean. Address slicers /Sheet1/slicer[i]
                (1-based per sheet) or /Sheet1/slicer[@name=X]. get reports kind(table|pivot)/column/caption; read {view:"structure"}
                lists per-sheet slicers. remove by either path.
                """,
                ["xlsx/table", "xlsx/pivot"]),

            ["bridges"] = (
                """
                ## markdown / csv bridges (M5)
                Import: office_create {file:"report.docx", from:"notes.md"} — Markdig-parsed GFM: headings (styles),
                        bold/italic/strike/code runs, nested bullet+number lists, pipe tables (header row bolded),
                        links, images (sandbox-resolved), blockquotes, code blocks, horizontal rules. Raw HTML and
                        missing images degrade to meta.warnings, never failures.
                        office_create {file:"orders.xlsx", from:"orders.csv"} — RFC 4180 (quotes, embedded commas/newlines),
                        delimiter sniffed (, ; tab |) or forced via delimiter:",". Cells are typed: numbers/dates/booleans
                        convert, leading-zero codes like "007" STAY TEXT; >50k cells stream through the SAX writer.
                Matrix: .md/.markdown → .docx, .csv/.tsv → .xlsx — any other pair is invalid_args naming this matrix.
                Export: office_read {view:"markdown"} (docx) — GFM of the body; structure round-trips with the importer.
                        office_read {view:"csv", sheet?:"Name", range?:"A1:C10"} (xlsx) — one sheet as RFC 4180 csv.
                CLI:    aioffice create report.docx --from notes.md;  aioffice read report.docx --view markdown
                        aioffice create orders.xlsx --from orders.csv; aioffice read orders.xlsx --view csv --sheet Sheet1
                """,
                ["edit-ops", "errors"]),

            ["docx/table"] = (
                """
                ## docx deep tables (M5)
                table set (path /body/table[1]): borders all|outer|none (+ borderColor hex, borderWidthPt),
                  shading hex, headerRow "true" (bolds row 1 + repeats it on every page), width ("12cm"|"100%"),
                  columnWidths ["3cm","auto","2.5cm"] (one entry per column), alignment left|center|right, cellPaddingCm.
                cell set (path /body/table[1]/tr[1]/tc[1]): text, mergeRight N (gridSpan = the TOTAL span; 1 unmerges),
                  mergeDown N (vMerge chain, total rows covered; 1 unmerges), shading hex, valign top|center|bottom.
                Merged-away cells disappear from get/render; render {to:"html"} emits real colspan/rowspan.
                Add rows: {"op":"add","path":"/body/table[1]","type":"tr"}; add a table: type "table" + {rows, columns}.
                """,
                ["docx/paragraph", "edit-ops"]),

            ["docx/field"] = (
                """
                ## docx fields (M5)
                Add: {"op":"add","path":"/footer[1]/p[1]","type":"field","props":{"kind":"pageNumber"}}
                kinds: pageNumber (PAGE), numPages (NUMPAGES), date (DATE, optional format:"yyyy"), docTitle (TITLE).
                leadingText prefixes literal text: 'Page X of Y' = a pageNumber field, then a numPages field with
                {"leadingText":" of "}. Fields live in body or header/footer paragraphs; get/render show kind + instruction.
                """,
                ["docx/header", "edit-ops"]),

            ["docx/header"] = (
                """
                ## docx header/footer variants (M5)
                Paths: /header[1] (= /header[default]), /header[firstPage], /header[even] — same for /footer[…].
                Add: {"op":"add","path":"/header[firstPage]","type":"header","props":{"text":"Cover"}} — wires
                w:titlePg (firstPage) / w:evenAndOddHeaders (even) automatically. All three variants coexist;
                named paths address content inside: {"op":"set","path":"/header[firstPage]/p[1]","props":{…}}.
                get reports the variant; combine with docx/field for 'Page X of Y' footers.
                """,
                ["docx/field", "addressing"]),

            ["xlsx/dataValidation"] = (
                """
                ## xlsx data validation (M5) — add type "dataValidation" (path = the range)
                list:  {"op":"add","path":"/Sheet1/B2:B20","type":"dataValidation","props":{"kind":"list",
                       "values":["Open","Closed"]}}  — or "sourceRange":"/Lists/A1:A3" (not both); Excel shows a dropdown.
                rules: kind wholeNumber|decimal|date|textLength + operator (between needs value+value2; > >= < <= == != need value),
                       e.g. {"kind":"wholeNumber","operator":"between","value":1,"value2":10}; dates as "2026-01-01".
                extras: errorTitle, errorMessage, errorStyle stop|warning|information, allowBlank.
                Read back: get /Sheet1/dataValidation[1] (dvKind, operator, values…); structure lists per-sheet rules;
                remove by the same path.
                """,
                ["xlsx/cell", "edit-ops"]),

            ["xlsx/sparkline"] = (
                """
                ## xlsx sparklines (M5) — add type "sparkline" (path = the host CELL)
                {"op":"add","path":"/Sheet1/E2","type":"sparkline","props":{"dataRange":"A2:D2","kind":"line",
                 "color":"376092","markers":true}}
                kinds: line | column | winLoss (markers: line only). One sparkline per op; rows of them = one op per cell.
                Read back: get /Sheet1/sparkline[1] (sparklineKind, cell, dataRange); structure lists sparklineGroups;
                remove by the same path.
                """,
                ["xlsx/cell", "edit-ops"]),

            ["xlsx/comment"] = (
                """
                ## xlsx threaded comments (M5)
                Add:   {"op":"add","path":"/Sheet1/B2","type":"comment","props":{"text":"looks wrong","author":"Reviewer"?}}
                Reply: {"op":"add","path":"/Sheet1/comment[@id=GUID]","type":"reply","props":{"text":"fixed","author"?}}
                       (the GUID id comes back from the add / from read {view:"comments"} / structure)
                Real xl/threadedComments parts (modern Excel threads) plus a legacy note fallback for old clients.
                Read:  office_read {view:"comments"} lists threads; get /Sheet1/B2 shows the anchored thread;
                       get /Sheet1/comment[@id=GUID] returns the whole thread. Remove the root to drop the thread.
                """,
                ["xlsx/cell", "edit-ops"]),

            ["pptx/table"] = (
                """
                ## pptx native tables (M5) — add type "table" on /slide[i]
                {"op":"add","path":"/slide[1]","type":"table","props":{"rows":3,"cols":4,"headerRow":true,
                 "style":"medium","x":"2cm","y":"5cm","w":"28cm","columnWidths":[…]?}}
                styles: light | medium | dark — direct paint (banded rows + header fill), no theme dependency.
                Cells: {"op":"set","path":"/slide[1]/table[1]/tr[1]/tc[1]","props":{"text":"Q3","mergeRight":2,"mergeDown"?}}
                — pptx mergeRight/mergeDown = how many cells to ABSORB (mergeRight:1 -> span 2; docx counts total span instead).
                get /slide[1]/table[1] reports rows/cols/headerRow/rowsDetail with per-cell paths + merge shape.
                render {to:"svg"} draws the real grid. Remove by the table path.
                """,
                ["pptx/slide", "edit-ops"]),

            ["pptx/animation"] = (
                """
                ## pptx animations (M4 entrance, M5 emphasis+exit) — add type "animation" (path = the shape)
                {"op":"add","path":"/slide[1]/shape[@id=4]","type":"animation","props":{"effect":"pulse","trigger":"click"?,
                 "duration":"0.5s"?,"delay"?,"direction"?,"color"?}}
                entrance: appear, fade, flyIn, wipe   emphasis: pulse, grow, spin, colorPulse (takes color)
                exit:     fadeOut, flyOut, wipeOut    direction: flyIn/wipe/flyOut/wipeOut only (left|right|top|bottom)
                triggers: click | withPrevious | afterPrevious. read {view:"structure"} lists per-slide animation order.
                M6 timeline: move /slide[i]/animation[2] before/after /slide[i]/animation[1] reorders; set retunes props;
                remove /slide[i]/animation[k] drops one.
                """,
                ["pptx/shape", "pptx/slide"]),

            ["pptx/section"] = (
                """
                ## pptx slide sections (M6) — add type "section" on the "/" root
                {"op":"add","path":"/","type":"section","props":{"name":"Intro","afterSlide":0}}
                afterSlide is 0-based: 0 = before slide 1, N = after slide N; the section claims the still-unsectioned
                slides from there up to the next section. set /section[i] {name} renames; remove /section[i] drops it
                (slides survive, just unsectioned). read {view:"outline"} groups slides by section; get /section[i]
                returns its name + slides. Sections track slides by sldId, so they survive reordering.
                """,
                ["pptx/slide", "pptx/slideSize"]),

            ["pptx/master"] = (
                """
                ## pptx master/layout editing (M6) — the M1 read-only debt is paid
                set /master[m]            {background (hex), accent1..accent6 (theme color scheme)}
                set /master[m]/layout[l]  {background}
                add /master[m] {type:layout, props:{name, basedOn?}}  clones an existing layout (basedOn 1-based)
                set/add/remove /master[m]/shape[i], /master[m]/layout[l]/shape[i]  reuse the slide shape ops
                Use a cloned layout on a new slide: add type:slide props:{layout:N} (1-based). remove a layout only
                when no slide references it. get /master[m] lists layouts (name, type, usedBySlides).
                """,
                ["pptx/slide", "pptx/shape"]),

            ["pptx/slideSize"] = (
                """
                ## pptx slide size (M6) — set on the "/" root
                {"op":"set","path":"/","props":{"slideSize":"16:9|4:3|16:10|A4|letter"}}  (a named preset)
                or explicit {"width":"33.87cm","height":"19.05cm"} (cm/in/emu) — not both. Rewrites p:sldSz; existing
                shapes keep their coordinates. get / reports slideSize, widthCm, heightCm, slideCount, sectionCount.
                """,
                ["pptx/section", "pptx/slide"]),

            ["pptx/comment"] = (
                """
                ## pptx comments + replies (M2 comments, M5 threads)
                Add:   {"op":"add","path":"/slide[2]","type":"comment","props":{"text":"tighten this","author"?}}
                Reply: {"op":"add","path":"/slide[2]/comment[@id=1]","type":"reply","props":{"text":"done","author"?}}
                Replies are p15 threadingInfo on classic comment parts — PowerPoint 2013+ shows real threads.
                Read: office_read {view:"comments"} lists threads per slide; get /slide[2]/comment[@id=1] returns the
                thread; remove the root comment to drop it (replies go with it).
                """,
                ["pptx/slide", "edit-ops"]),

            ["addressing"] = (
                """
                ## Addressing grammar (1-based indices everywhere)
                docx:  /body/p[3]   /body/table[1]/tr[2]/tc[1]   /body/p[3]/run[2]   /header[1]/p[1]   /footer[1]/p[1]
                xlsx:  /Sheet1/A1   /Sheet1/A1:C10   /Sheet1/row[3]   /Sheet1/chart[1]   sheet names with spaces or specials are quoted: /'Q3 Data'/B2 (escape ' as '')
                pptx:  /slide[2]   /slide[2]/shape[3]   /slide[2]/shape[@id=7] (stable-id form, canonical in results)   /slide[2]/shape[3]/p[1]
                       /master[1]   /master[1]/layout[2]   /master[1]/shape[1]   (M6: editable — background, accents, shapes, cloned layouts)
                       /section[i] (slide section)   /slide[i]/animation[k] (timeline)
                docx M6:  /body/p[i]/omath[j] (inline equation)   /section[i] (page setup + columns)
                docx M8:  /caption[@label=Figure][1] (caption, 1-based per label; label Figure|Table|Equation)
                xlsx M6:  /Sheet1/table[@name=X] (Excel Table)   /Sheet1/row[a]:row[b], /Sheet1/col[a]:col[b] (outline group spans)
                xlsx M8:  /Sheet1/slicer[i] or /Sheet1/slicer[@name=X] (table-column / pivot-field slicer)
                "/" alone is the document root: pptx slide size + sections, and document-level office_get.
                office_query returns canonical paths; office_get / office_edit accept them verbatim.
                Positional indices DRIFT after inserts/removes — re-run office_query instead of reusing old indices.
                """,
                ["selectors", "edit-ops"]),

            ["selectors"] = (
                """
                ## Selector grammar (office_query)
                form:       element predicates*        element = p | run | table | cell | row | sheet | slide | shape | image | * (any)
                predicates: [attr OP value]            OP = "=", "!=", ">", ">=", "<", "<=", "*=" (text contains)
                            :contains('text')          escape ' inside the quotes as ''
                examples:   p[style=Heading1]          cell[value>100]          shape:contains('Q3')
                            run[bold=true]             p:contains('budget')     *[style=Caption]
                Numeric comparisons apply when the value parses as a number; otherwise ordinal text comparison.
                xlsx note: value predicates on formula cells without a cached value add a formula_not_evaluated warning to meta.warnings.
                Results carry canonical paths (max 50; data.total is the full hit count, data.truncated says if the list was cut).
                """,
                ["addressing", "errors"]),

            ["edit-ops"] = (
                """
                ## office_edit ops — the only mutation surface
                kinds: set (change props), add (insert child; path = PARENT; needs type), remove (delete), move (reposition),
                       replace (find/replace: props {find, replace, regex?, matchCase?, wholeWord?}; scope = any container path,
                       or "/" for the whole document — docx body+headers+footers, every sheet, every slide incl. notes).
                replace results carry {replacements, locations (max 20)}; 0 matches is ok:true + a find_no_match warning.
                replace with track:true (docx) records w:del+w:ins revision pairs; tracked replace is body-scoped.
                regex uses .NET syntax with a 2s match budget (timeout -> invalid_args). CLI sugar: --find X --replace Y [--regex --match-case --whole-word].
                position (add/move): 1-based integer within parent | "before:<path>" | "after:<path>" | omit = append.
                props are string-valued: {"text":"Hi","bold":"true","size":"12pt","fill":"FF0000"}; sizes unit-qualified (12pt, 2cm); colors hex or named.
                Atomicity: the whole ops[] batch is one save — any failure means NOTHING is written.
                Ops resolve at execution time: an add shifts the indices later ops see. Put set/remove targets BEFORE inserts, or split into two edits and re-query.
                expect_rev: pass the last meta.rev; mismatch -> stale_address before any write. dry_run:true validates without writing.
                Every successful edit auto-snapshots the pre-image (file_snapshot lists/restores).
                """,
                ["docx/paragraph", "xlsx/cell", "pptx/shape", "addressing"]),

            ["envelope"] = (
                """
                ## Result envelope — every tool, every time
                { "ok": bool, "data": {...}|null, "error": {code,message,suggestion,candidates?}|null,
                  "meta": { file?, rev?, elapsedMs, version, warnings?: [{code,message}] } }
                error.suggestion is never empty — it is the next action to take.
                rev = first 12 hex of SHA-256 of the file bytes AFTER the call; feed it to office_edit expect_rev.
                meta.warnings carries non-fatal issues (e.g. formula_not_evaluated); check it even when ok=true.
                """,
                ["errors"]),

            ["errors"] = (
                """
                ## Error codes
                invalid_args         bad parameter/selector/topic/snapshot number; suggestion shows the correct shape, candidates list near matches
                file_not_found       check spelling relative to the workspace root, or office_create it
                sandbox_denied       path escapes the workspace (symlinks count); widen with --workspace / AIOFFICE_WORKSPACE
                invalid_path         element path does not resolve; candidates[] holds the nearest canonical paths — pick one
                stale_address        expect_rev mismatch: file changed since you read it; re-read, then retry (nothing was written)
                unsupported_feature  capability not in this milestone; suggestion names the workaround
                file_too_large       file exceeds the opt-in AIOFFICE_MAX_FILE_MB cap (default unlimited); raise/unset it or split the file
                format_corrupt       not valid OOXML; try office_validate and file_snapshot restore
                internal_error       our bug; report with office_status output
                formula_not_evaluated  WARNING in meta.warnings, not an error: formula text returned instead of a value
                """,
                ["envelope"]),

            ["docx/revisions"] = (
                """
                ## docx tracked changes (M2)
                Record: office_edit {track:true, author:"Reviewer", ops:[…]} — text set/add/remove ops become w:ins/w:del.
                Only TEXT changes can be tracked ({"props":{"text":…}}); tracked formatting -> invalid_args naming the workaround.
                Author resolution: op props.author > tool arg author > AIOFFICE_AUTHOR env > "AIOffice". Date = now UTC.
                Read:    office_read {view:"revisions"} -> [{id, kind:insert|delete, author, date, text, path}].
                Resolve: {"op":"accept","path":"/revision[@id=3]"} applies one; path "/body" resolves every revision in scope.
                         {"op":"reject",…} undoes instead. Revisions are never "removed" — only accepted or rejected.
                """,
                ["docx/comment", "edit-ops"]),

            ["docx/comment"] = (
                """
                ## docx comments (M2)
                Add:    {"op":"add","path":"/body/p[2]","type":"comment","props":{"text":"…","author":"Reviewer"?}}
                        (path = anchored content: a paragraph or a run; result reports /comment[@id=N] + anchor).
                Read:   office_read {view:"comments"} -> [{id, author, date, text, anchorPath, anchorText}].
                Get:    office_get {path:"/comment[@id=2]"}. Remove: {"op":"remove","path":"/comment[@id=2]"}.
                """,
                ["docx/revisions", "edit-ops"]),

            ["docx/style"] = (
                """
                ## docx style definitions (M2)
                Add:    {"op":"add","path":"/styles","type":"style","props":{"id":"Callout","kind":"paragraph"?,
                        "name"?,"basedOn"?,"bold"?,"italic"?,"underline"?,"color"?,"fontSize"?,"alignment"?,
                        "spacingBefore"?,"spacingAfter"?}}
                Modify: {"op":"set","path":"/style[@id=Callout]","props":{…}} (id/kind are fixed).
                Apply:  {"op":"set","path":"/body/p[2]","props":{"style":"Callout"}}.
                Read:   office_read {view:"styles"}; office_get {path:"/style[@id=Callout]"}.
                Remove: {"op":"remove","path":"/style[@id=Callout]"} (custom styles only; built-ins are modified, not removed).
                """,
                ["docx/paragraph", "edit-ops"]),

            ["docx/image"] = (
                """
                ## docx inline pictures (M2)
                Add: {"op":"add","path":"/body","type":"image","props":{"src":"logo.png","width":"10cm"?,"height":"3cm"?}}
                     position "before:<path>"/"after:<path>" places it; omit = append. PNG/JPEG; src resolves through the
                     workspace sandbox (escaping it -> sandbox_denied). Omitted width/height keep the natural aspect ratio.
                """,
                ["edit-ops", "addressing"]),

            ["xlsx/pivot"] = (
                """
                ## xlsx pivot tables (M2)
                Add: {"op":"add","path":"/Sheet1","type":"pivot","props":{
                       "sourceRange":"A1:E7","targetSheet":"Pivot","name"?,"targetAnchor"?,
                       "rows":["Region"],"columns"?,"filters"?,
                       "values":[{"field":"Sales","agg":"sum|average|count|min|max"}]}}
                     path = SOURCE sheet; sourceRange needs a header row. targetSheet is created when absent.
                Get: office_get {path:"/Pivot/pivot[1]"} or {path:"/Pivot/pivot[@name=SalesPivot]"} (the TARGET sheet).
                Remove: {"op":"remove","path":"/Pivot/pivot[@name=SalesPivot]"}. Excel recomputes on open (refreshOnLoad).
                """,
                ["xlsx/cell", "edit-ops"]),

            ["xlsx/conditionalFormat"] = (
                """
                ## xlsx conditional formatting (M2)
                Add (path = the range): {"op":"add","path":"/Sheet1/A1:C10","type":"conditionalFormat","props":{…}}
                kinds: cellIs      {operator:"> >= < <= == != between", value, value2 (between only), fill?, color?, bold?}
                       colorScale  {minColor, maxColor, midColor?}   dataBar {color}   containsText {text, fill?, color?, bold?}
                Get: office_get {path:"/Sheet1/conditionalFormat[1]"}; remove by the same path (later indices shift down).
                """,
                ["xlsx/cell", "edit-ops"]),

            ["xlsx/image"] = (
                """
                ## xlsx anchored pictures (M2)
                Add: {"op":"add","path":"/Sheet1","type":"image","props":{"src":"logo.png","anchor":"E2","name"?,
                     "widthPx"?,"heightPx"?}} — PNG/JPEG; src resolves through the workspace sandbox.
                Get: office_get {path:"/Sheet1/image[1]"} -> name, format, anchor, size. Remove by the same path.
                """,
                ["xlsx/cell", "edit-ops"]),

            ["pptx/slide"] = (
                """
                ## pptx slide (incl. M2 background, M3 transition)
                Add:        {"op":"add","path":"/slide[3]","type":"slide","props":{"title"?,"background"?}} — becomes slide 3.
                Background: {"op":"set","path":"/slide[1]","props":{"background":"0F172A"}} — a real p:bg solid fill (hex).
                Transition: {"op":"set","path":"/slide[1]","props":{"transition":"fade","transitionDuration":"0.5s"}} (see pptx/transition).
                Move/remove a slide by its /slide[i] path. render {to:"svg"|"png", scope:"/slide[N]"} to look after edits;
                render {to:"pdf"} prints the whole deck, one page per slide.
                """,
                ["pptx/shape", "pptx/notes", "pptx/transition"]),

            ["pptx/notes"] = (
                """
                ## pptx speaker notes (M2)
                Path /slide[i]/notes addresses the whole notes body (no /p[j] beneath it).
                Set:    {"op":"set","path":"/slide[2]/notes","props":{"text":"line1\nline2"}} (replaces; \n = new paragraph)
                Append: {"op":"add","path":"/slide[2]/notes","props":{"text":"follow-up"}}
                Remove: {"op":"remove","path":"/slide[2]/notes"}. office_read outline + office_get /slide[i]/notes read them back.
                """,
                ["pptx/slide", "edit-ops"]),

            ["pptx/image"] = (
                """
                ## pptx pictures (M2)
                Add: {"op":"add","path":"/slide[1]","type":"image","props":{"src":"logo.png","x"?,"y"?,"w"?,"h"?,"name"?}}
                     sizes unit-qualified ("6cm","1.5in") or EMU; omit w/h to keep natural size/aspect. PNG/JPEG;
                     src resolves through the workspace sandbox. Result path: /slide[1]/shape[@id=N] (remove by that path).
                """,
                ["pptx/slide", "edit-ops"]),

            ["docx/paragraph"] = (
                """
                ## docx paragraph — props (office_edit set/add, type "paragraph")
                text     replaces the paragraph's text content
                style    named style, e.g. Heading1, Heading2, Normal, Quote
                bold / italic / underline   "true" | "false"
                size     unit-qualified, e.g. "12pt"
                color    hex "FF0000" or named
                align    left | center | right | justify
                add child runs with type "run" for mixed formatting inside one paragraph.
                """,
                ["edit-ops", "addressing"]),

            ["xlsx/cell"] = (
                """
                ## xlsx cell — props (office_edit set; address cells as /Sheet1/B2)
                value         cell value; numbers/dates auto-typed, anything else stays text
                formula       e.g. "=SUM(A1:A10)" (calculated on save where supported)
                numberFormat  e.g. "0.00", "yyyy-mm-dd"
                bold / italic "true" | "false"
                fill          background color, hex "FFFF00"
                office_get returns value, valueType, formula, numberFormat. Reading a formula cell without
                a cached value adds a formula_not_evaluated warning and returns the formula text.
                """,
                ["edit-ops", "addressing"]),

            ["pptx/shape"] = (
                """
                ## pptx shape — props (office_edit set/add, type "shape")
                text          replaces the shape's text content
                x / y / w / h position and size, unit-qualified ("2cm", "1.5in") or EMU integers
                fill          hex "4472C4" or named
                fontSize      unit-qualified, e.g. "18pt"
                bold / italic "true" | "false"
                shape (M3)    preset geometry on add: rect|roundRect|ellipse|triangle|diamond|arrow|line ("line" builds a connector; fill = stroke)
                flipH / flipV (M3) "true" mirrors the geometry
                hyperlink (M8) the shape's click action: a url ("https://…"), a slide jump ("#slide:4"), or a show action
                              ("#first" | "#last" | "#next" | "#prev" | "#end"); "" clears it. get reports the canonical form back.
                linkText (M8) wraps the shape's runs in a text hyperlink (same target grammar as hyperlink).
                z-order (M3): {"op":"move","path":"/slide[1]/shape[@id=5]","position":"front|back|forward|backward"} (front = painted last/topmost)
                add slides with type "slide" on path "/"; render→look→fix: office_render {to:"svg", scope:"/slide[N]"} after each visual change.
                """,
                ["edit-ops", "addressing"]),

            ["docx/list"] = (
                """
                ## docx lists (M3) — add type "p" with list props
                list         bullet | number
                level        0-based nesting (default 0)
                listRestart  "true" restarts a number list at this item
                {"op":"add","path":"/body","type":"p","props":{"text":"Step one","list":"number"}}
                Nested: {"props":{"text":"Detail","list":"number","level":"1"}}.
                read {view:"text"} shows "1." / "•" markers; render {to:"html"} emits real <ol>/<ul>.
                """,
                ["docx/paragraph", "edit-ops"]),

            ["docx/link"] = (
                """
                ## docx hyperlinks / bookmarks / footnotes (M3)
                link:     {"op":"add","path":"/body/p[3]","type":"link","props":{"text":"site","url":"https://…"}}
                          internal jump: {"props":{"text":"see intro","anchor":"Intro"}} (a bookmark name)
                bookmark: {"op":"add","path":"/body/p[1]","type":"bookmark","props":{"name":"Intro"}} (letter/_ start, ≤40 chars)
                footnote: {"op":"add","path":"/body/p[3]","type":"footnote","props":{"text":"the note"}} — reference lands at the paragraph end.
                """,
                ["docx/paragraph", "edit-ops"]),

            ["docx/section"] = (
                """
                ## docx page setup + columns (M3/M6) — set on /section[1]
                pageSize     A4 | Letter | Legal | A3 | A5
                orientation  portrait | landscape (width/height swap automatically)
                marginTop / marginBottom / marginLeft / marginRight   unit-qualified, e.g. "2cm"
                columns (M6) newspaper column count (1 clears)   columnGap (M6) gap between columns, e.g. "1.25cm"
                {"op":"set","path":"/section[1]","props":{"pageSize":"A4","orientation":"landscape"}}
                two columns: {"op":"set","path":"/section[1]","props":{"columns":2}}
                office_get /section[1] reflects everything (cm, columns, columnGapCm). New section breaks:
                {"op":"add","type":"sectionBreak"}; push text to the next column: {"op":"add","type":"columnBreak"}.
                """,
                ["docx/paragraph", "addressing"]),

            ["equations"] = (
                """
                ## equations — LaTeX in, Office Math out (docx + pptx)
                docx add: {"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"E = mc^2","display":false}}
                     display:true -> centered block (m:oMathPara); false (default) -> inline at /body/p[i]/omath[j].
                pptx add: {"op":"add","path":"/slide[1]","type":"equation","props":{"latex":"x = \\frac{1}{2}"}}
                     -> a text box holding the OMML; or target /slide[i]/shape[@id=N] to append to that box.
                     pptx path: /slide[i]/shape[@id=N]/omath[k].
                Read: office_get <omath path> -> {latex, ...} (your source, stored verbatim).
                      docx office_read {view:"text"} shows $…$ (inline) / $$…$$ (display) markers.
                Subset (the SAME shared converter for both formats): x^2 a_i ; \\frac \\dfrac \\sqrt ;
                        \\sum \\prod \\int \\lim with _/^ ;
                        \\begin{pmatrix|bmatrix|Bmatrix|vmatrix|matrix}…&…\\\\…\\end{…} ; \\left( \\right) ;
                        \\bar \\overline ; \\text \\mathrm \\mathbf ; Greek \\alpha..\\omega ;
                        \\pm \\times \\cdot \\leq \\geq \\neq \\infty \\partial \\nabla \\rightarrow ; spacing \\, \\quad.
                Unknown commands degrade to literal runs + an equation_partial warning (file still validates).
                xlsx has NO equation object (Excel carries cell formulas, not OMML math): add type "equation"
                on xlsx returns unsupported_feature — put the math in a cell formula or embed a rendered image.
                Tracked-changes (docx) does not record equation adds.
                """,
                ["docx/equation", "pptx/equation", "rtl"]),

            ["docx/equation"] = (
                """
                ## docx equations (M6) — add type "equation"
                props: latex (required, the LaTeX source) ; display (bool: true = centered block, false = inline).
                Inline path: /body/p[i]/omath[j]. See office_help {topic:"equations"} for the supported LaTeX subset.
                """,
                ["equations", "docx/paragraph"]),

            ["pptx/equation"] = (
                """
                ## pptx equations (M10) — add type "equation"
                props: latex (required, the LaTeX source) ; x,y,w,h (place a new text box) ; fontSize (the fallback run).
                Target /slide[i] to create a text box for the equation, or /slide[i]/shape[@id=N] to append to one.
                Canonical path: /slide[i]/shape[@id=N]/omath[k]. office_get returns {latex}. remove drops the equation.
                Stored as native OMML inside an mc:AlternateContent / a14:m (the form PowerPoint writes), validator-clean.
                Same LaTeX subset and equation_partial behaviour as docx — see office_help {topic:"equations"}.
                """,
                ["equations", "pptx/shape"]),

            ["embeds"] = (
                """
                ## embedded objects (M10) — embed/list/extract files in docx/xlsx/pptx
                Embed any file (a source .xlsx attached to a report, a .pdf, a .zip…) as an OLE/package object.
                Add: {"op":"add","path":"CONTAINER","type":"embed","props":{"src":"data.xlsx","name":"optional","icon":"optional png/jpeg"}}
                     docx CONTAINER = /body (or a tc/header/footer); xlsx = /Sheet1 (anchored); pptx = /slide[i].
                     The media type is sniffed from the file; the op returns the canonical embed path.
                List: office_read {view:"embeds"} -> every embed's {path, name, mediaType, size, container}.
                      structure view also includes embeds. office_get on an embed path returns metadata, NOT the bytes.
                Extract: {"op":"extract","path":"/embed[1]","props":{"to":"out/data.xlsx"}}  (a producing op; does NOT
                     modify the source document). Paths: docx /embed[i]; xlsx /Sheet1/embed[i]; pptx /slide[i]/embed[j]
                     or /slide[i]/embed[@id=N]. The extracted bytes are byte-identical to what was embedded.
                Remove: the normal remove op on the embed path. The src and the extract dest are both sandbox-resolved.
                """,
                ["addressing", "edit-ops"]),

            ["rtl"] = (
                """
                ## right-to-left / bidi (M6) — docx, prop "rtl" (bool)
                paragraph: {"op":"set","path":"/body/p[5]","props":{"rtl":true}} (w:bidi; also right-aligns).
                run:       {"op":"set","path":"/body/p[5]/run[2]","props":{"rtl":true}} (w:rtl, mixed-direction span).
                table:     {"op":"set","path":"/body/table[1]","props":{"rtl":true}} (w:bidiVisual; mirrors columns).
                office_get reports rtl on each. aioffice never reshapes glyphs; it sets the OOXML direction only.
                """,
                ["docx/paragraph", "docx/table"]),

            ["xlsx/sheet"] = (
                """
                ## xlsx sheet-level features (M3)
                freeze:     {"op":"set","path":"/Sheet1","props":{"freezeRows":"1","freezeCols":"2"}} (0 clears an axis)
                autoFilter: {"op":"set","path":"/Sheet1/A1:D20","props":{"autoFilter":"true"}} (one per sheet; false clears)
                print:      {"op":"set","path":"/Sheet1","props":{"orientation":"landscape","paperSize":"A4","fitToWidth":"1","printArea":"A1:F40"}}
                names:      {"op":"add","path":"/Sheet1/B2:C5","type":"name","props":{"name":"SalesData","scope"?:"workbook|sheet"}}
                            then formulas just use it: {"op":"set","path":"/Sheet1/D6","props":{"value":"=SUM(SalesData)"}} (evaluates).
                            get /name[@name=X] or /Sheet1/name[@name=X].
                Streaming reads: files over 20 MB (or args.stream:true) answer read stats/text and cell/range get
                via a SAX scan (no DOM). M6 streaming WRITES: with stream:true (or a >20 MB file) an all-streamable
                batch (set value / set values / set a formula) rewrites the workbook in place via the SAX writer —
                see xlsx/sheet streaming and office_help {topic:"xlsx/table"}/{topic:"xlsx/group"}.
                """,
                ["xlsx/cell", "xlsx/table", "xlsx/group"]),

            ["xlsx/table"] = (
                """
                ## xlsx Excel Tables / ListObjects (M6) — add type "table" (path = the range, first row = headers)
                {"op":"add","path":"/Sheet1/A1:D20","type":"table","props":{"name":"Sales","style":"medium2",
                 "headerRow":true,"totalsRow":false,"totals":{"Amount":"sum"},"bandedRows":true}}
                style: none | light1-21 | medium1-28 | dark1-11 (e.g. "medium2"), or the full "TableStyleMedium2".
                totals: column->function map (sum/average/count/countNumbers/min/max/stdDev/var); turns the totals row on.
                Structured references evaluate (the table exists before save): {"value":"=SUM(Sales[Amount])"}.
                get /Sheet1/table[@name=Sales] describes range/style/columns/totals. remove drops the ListObject but
                KEEPS the cell data.
                """,
                ["xlsx/cell", "xlsx/group"]),

            ["xlsx/group"] = (
                """
                ## xlsx outline grouping (M6) — add type "group" over a row/column span
                rows: {"op":"add","path":"/Sheet1/row[2]:row[6]","type":"group","props":{"collapsed":true}}
                cols: {"op":"add","path":"/Sheet1/col[B]:col[E]","type":"group"}
                The span is ONE path segment: row[a]:row[b] or col[a]:col[b]. collapsed:true collapses the new group;
                nesting raises the outline level. remove over the same span ungroups one level.
                A row/column get and read {view:"structure"} report outlineLevel/collapsed; Excel draws the symbols.
                """,
                ["xlsx/sheet", "edit-ops"]),

            ["pptx/chart"] = (
                """
                ## pptx charts (M3) — add type "chart" on /slide[i]
                kind        bar | line | pie (pie: exactly one series)
                categories  ["Q1","Q2",…]            series  [{"name":"Sales","values":[10,20]}] (null = gap)
                title, x, y, w, h   optional
                Data is cached literally (c:strLit/c:numLit, no embedded workbook): projections report dataEditable:false.
                Cross-doc dataFrom: {"props":{"kind":"bar","dataFrom":"book.xlsx!Sheet1/A1:B5"}} replaces categories/series —
                first column = categories, header row = series names, remaining columns = series values; the workbook is
                sandbox-resolved; a wrong range fails typed with candidates. Quote sheet names: book.xlsx!'Q3 Data'/A1:C5.
                Address /slide[i]/chart[k] for get (full data readback), set (title) and remove.
                """,
                ["pptx/slide", "edit-ops"]),

            ["pptx/transition"] = (
                """
                ## pptx slide transitions (M3) — slide-level set
                {"op":"set","path":"/slide[2]","props":{"transition":"fade|push|wipe|none","transitionDuration":"0.5s"}}
                Duration maps to p14:dur milliseconds (PowerPoint 2010+; older clients use the spd fallback).
                office_get /slide[2] and read {view:"outline"} report transition/transitionDuration.
                """,
                ["pptx/slide", "edit-ops"]),
        };

    /// <summary>All topic names (index first, then alphabetical).</summary>
    public static IReadOnlyList<string> Names { get; } =
        [IndexTopic, .. Topics.Keys.Where(k => k != IndexTopic).Order(StringComparer.Ordinal)];

    /// <summary>
    /// Resolves a topic (null/empty → index) to the <c>office_help</c> data payload,
    /// or throws <c>invalid_args</c> with nearest-match candidates.
    /// </summary>
    public static object Get(string? topic)
    {
        var key = string.IsNullOrWhiteSpace(topic) ? IndexTopic : topic.Trim();
        if (Topics.TryGetValue(key, out var entry))
        {
            return new { topic = key.ToLowerInvariant(), doc = entry.Doc, related = entry.Related };
        }

        throw new AiofficeException(
            ErrorCodes.InvalidArgs,
            $"Unknown help topic: '{topic}'.",
            "Call office_help with no topic to list all topics.",
            candidates: Nearest(key));
    }

    private static IReadOnlyList<string> Nearest(string requested)
    {
        var hits = Names
            .Where(n => n.Contains(requested, StringComparison.OrdinalIgnoreCase) ||
                        requested.Contains(n, StringComparison.OrdinalIgnoreCase))
            .ToList();
        return hits.Count > 0 ? hits : Names;
    }
}
