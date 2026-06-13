# audit-demo/ — deliberately-bad files for `aioffice audit`

These three files are **intentionally flawed** so you can run the M7 audit verb
against them and see real findings (and `--fix` in action). They are the
*before* state; the polished, audited+fixed versions live one directory up in
`fixtures/manual-check/`.

```bash
# from a workspace that contains this folder
aioffice audit audit-demo/report.docx
aioffice audit audit-demo/report.docx --fix      # fixes alt text, table header, doc title
aioffice audit audit-demo/report.docx            # re-run: heading-skip + empty-heading remain (report-only)
aioffice validate audit-demo/report.docx         # still 0 schema errors
```

What each file trips (findings are *data* — the command exits `0`):

- **report.docx** — image with no alt text (`a11y_no_alt_text`, error, autofix),
  H1 → H3 heading skip (`a11y_heading_skip`, warning), an empty Heading2
  (`quality_empty_heading`, warning), a header-less table
  (`a11y_no_table_header`, error, autofix), no document title
  (`a11y_no_doc_title`, warning, autofix).
- **metrics.xlsx** — a `=B2/0` cell evaluating to `#DIV/0!`
  (`quality_formula_error`, error), a merged range inside the data region
  (`a11y_merged_data_cells`, warning), an image with no alt
  (`a11y_no_alt_text`, warning, autofix), no document title
  (`a11y_no_doc_title`, warning, autofix).
- **deck.pptx** — a shape positioned entirely off the slide
  (`quality_off_canvas`, warning), 6pt text (`a11y_tiny_font`, error), a picture
  with no alt (`a11y_no_alt_text`, warning, autofix), a slide with no title
  (`a11y_no_slide_title`, warning, autofix), plus a reading-order hint
  (`a11y_reading_order`, info).

`--fix` only applies the **safe, non-destructive** autofixes (placeholder alt
text, a marked table header row, a document/slide title, orphan-bookmark
removal). Everything else is reported for you to fix by hand — `--fix` never
rewrites your content. See `aioffice help audit` for the full code list.
