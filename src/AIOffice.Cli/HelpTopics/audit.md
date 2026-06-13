# audit

`aioffice audit <file>` inspects a document for **accessibility** and **quality**
problems and reports them as structured findings. Findings are *data*, never
errors: a successful audit always exits `0`, even when it surfaces
error-severity findings — the same way `aioffice validate` returns `ok:true`
while reporting issues.

    aioffice audit report.docx
    aioffice audit metrics.xlsx --category accessibility --severity warning
    aioffice audit deck.pptx --fix

## options

| option       | values                              | default | meaning                                  |
|--------------|-------------------------------------|---------|------------------------------------------|
| `--category` | `accessibility` · `quality` · `all` | `all`   | which checks to run                      |
| `--severity` | `error` · `warning` · `info`        | `info`  | minimum severity to report (a floor)     |
| `--fix`      | (flag)                              | off     | apply only the safe, non-destructive autofixes |

`--severity info` reports everything; `warning` hides info findings; `error`
hides info and warning findings.

## result shape

    {
      "findings": [
        { "id": "a11y_no_alt_text#/slide[2]/shape[@id=5]",
          "severity": "error", "category": "accessibility",
          "code": "a11y_no_alt_text", "path": "/slide[2]/shape[@id=5]",
          "message": "...", "suggestion": "...", "autofixable": true }
      ],
      "summary": { "errors": 1, "warnings": 2, "infos": 0 }
    }

With `--fix` the envelope also carries `"fixed": N` (how many autofixes landed)
and `"remaining": [ids]` (findings still present after the fix pass — the audit
is re-run so this list is authoritative).

Each finding's `id` is `code#path` and is stable within one run, so a future
`--fix` could target a single finding. Today `--fix` (no ids) applies every
*autofixable* finding.

## accessibility codes (`category: accessibility`)

| code                    | severity | autofix | fires on                                            |
|-------------------------|----------|---------|-----------------------------------------------------|
| `a11y_no_alt_text`      | error    | yes     | image/picture with no alternative text              |
| `a11y_no_table_header`  | error    | yes     | table whose first row is not marked a header        |
| `a11y_no_doc_title`     | warning  | yes     | docx/xlsx core property Title is empty              |
| `a11y_no_slide_title`   | warning  | yes     | pptx slide with no title placeholder text           |
| `a11y_heading_skip`     | warning  | no      | heading level jumps (H1 → H3) — report only         |
| `a11y_low_contrast`     | warning  | no      | text/background contrast below WCAG AA 4.5:1        |
| `a11y_tiny_font`        | warning  | no      | pptx text below ~12pt (error below ~8pt)            |
| `a11y_merged_data_cells`| warning  | no      | xlsx merged cells inside a data region              |
| `a11y_reading_order`    | info     | no      | pptx shapes whose z-order fights the visual order   |

## quality codes (`category: quality`)

| code                     | severity | autofix | fires on                                         |
|--------------------------|----------|---------|--------------------------------------------------|
| `quality_broken_ref`     | error    | no      | reference to a missing target (e.g. a name/range)|
| `quality_formula_error`  | error    | no      | xlsx cell holding an error value (#DIV/0!, #REF!)|
| `quality_broken_link`    | error    | no      | internal hyperlink to a missing bookmark         |
| `quality_empty_heading`  | warning  | no      | heading paragraph with no text                   |
| `quality_off_canvas`     | warning  | no      | pptx shape positioned outside the slide bounds   |
| `quality_empty_placeholder` | warning | yes/no | placeholder left with prompt text only           |
| `quality_orphan_bookmark`| info     | yes     | docx bookmark referenced by nothing              |
| `quality_duplicate_id`   | warning  | no      | duplicate element id (e.g. two shapes sharing one)|

## what `--fix` will and will not touch

`--fix` only applies **safe, non-destructive** autofixes:

- set a placeholder alt text `"(describe this image)"` on an image with none;
- mark a table's first row as a repeating header;
- set the document/slide title from the first heading, the file name, or a
  placeholder when nothing better exists;
- remove an orphan bookmark referenced by nothing.

Everything else (heading skips, low contrast, empty headings, off-canvas
shapes, formula errors, broken links/refs, tiny fonts, merged data cells) is
**report-only** — `--fix` never rewrites your content, restructures headings,
recolors text or moves shapes. After a `--fix` run those findings remain in
`remaining` so you can address them by hand.

Every `--fix` pass snapshots the pre-image first, so it is undoable with
`aioffice snapshot restore <file>`.
