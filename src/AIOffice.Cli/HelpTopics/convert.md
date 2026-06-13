# convert — move a document between formats

`aioffice convert <src> <dest>` opens the source read-only, builds the
destination **fresh**, and prints one envelope:
`{ok, data:{from, to, blocksWritten, dropped:[…], written}, meta}`.

Conversion is **content transfer** and inherently lossy across formats. Whatever
the target cannot represent (animations, charts, exact styling, formulas as
values, run colors in markdown, …) is named honestly in `data.dropped` and
surfaced as a single `convert_lossy` meta warning. Nothing is silently lost.

## Supported (src → dest) matrix

Routing is by the (source extension, destination extension) pair.

### Office ↔ office (via the content-neutral model)

A handler projects its file down to a neutral model — headings, paragraphs,
formatted runs, lists, tables, images — and the target handler rebuilds a fresh
file from it.

| from \ to | docx | xlsx | pptx | md |
|-----------|------|------|------|-----|
| **docx**  | —    | ✓    | ✓    | ✓  |
| **xlsx**  | ✓    | —    | ✓    | ✓  |
| **pptx**  | ✓    | ✓    | —    | ✓  |

- `docx → pptx`: each heading opens a slide; the bullets under it become that
  slide's body; tables become native pptx tables.
- `xlsx → docx`: each sheet becomes a heading + a table of its **display** values
  (formulas cross as their cached values).
- `pptx → docx`: slide titles become headings, bullets become a list.

### Text bridges

- `docx ↔ md` — the markdown bridge (headings, lists, tables, links).
- `xlsx ↔ csv` — the csv bridge (one sheet window, RFC 4180).
- `xlsx/docx/pptx → md` and `md → xlsx/docx/pptx` — markdown serialized to/from
  the neutral model, so **any** office format reaches markdown.

### Render targets (open src, render to dest)

`any office doc → .pdf | .png | .svg | .html | .txt` routes to the render layer
(Chrome/Chromium needed for `.pdf`/`.png`).

## Errors

- **same extension** (`convert a.docx b.docx`) → `invalid_args`: "use edit".
- **unknown destination** (`convert a.docx a.xyz`) → `unsupported_feature`
  naming the supported targets.
- **csv to/from a non-xlsx office format** → `unsupported_feature`: go via a
  workbook (csv → xlsx → docx).

## Examples

```
aioffice convert report.docx deck.pptx      # a slide per heading, bullets below
aioffice convert report.docx report.md      # headings/lists/tables as markdown
aioffice convert data.xlsx data.docx        # a table per sheet, cached values
aioffice convert deck.pptx outline.docx     # slide titles + bullets
aioffice convert report.docx report.pdf     # paged PDF via headless Chrome
```
