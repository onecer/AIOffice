# structural-fields (v1.2 — table of figures, index, mail-merge fields, docx)

v1.2 adds four docx structural-field `add` types. Each writes a real Word field
(complex field or SDT block) and passes `validate` with 0 OpenXML errors. Word
recomputes page numbers / regenerates the field when the document opens or the
field is refreshed; the cached values aioffice writes are marked with a warning.

These are not tracked-changes operations: run them **without** `--track`
(tracking a structural field add returns `unsupported_feature`).

## Table of figures — `add type:tableOfFigures`

Collects the document's captions of one label into a navigable list:

    aioffice edit r.docx --ops '[{"op":"add","path":"/body","type":"tableOfFigures","props":{"label":"Figure","position":"before /body/p[1]"}}]'

- `props.label`: `Figure` (default), `Table` or `Equation` — the same labels
  captions carry (see `add type:caption`).
- `props.title` (optional): a heading above the list.
- `props.position`: where to place the block (`before/after <path>`), or the
  op's `position`.
- The entries are pre-rendered from the cached caption text; a **`figures_cached`**
  warning notes Word repaginates the page numbers on open/refresh.
- `read --view structure` lists each table of figures with its label and path.

## Index entry — `add type:indexEntry`

Marks a term with an `XE` field so the index can pick it up:

    aioffice edit r.docx --ops '[{"op":"add","path":"/body/p[3]","type":"indexEntry","props":{"text":"AI"}}]'

- `props.text` (required): the term to index (no double quotes).
- `props.subEntry` (optional): a sub-entry under `text`.
- `props.find` (optional): place the `XE` field next to a specific occurrence of
  this text in the paragraph; omit to append it.

## Index — `add type:index`

Builds the index from every `XE` field, alphabetized:

    aioffice edit r.docx --ops '[{"op":"add","path":"/body","type":"index","props":{"columns":2}}]'

- `props.columns` (optional, default 1): the index column count.
- Entries are alphabetized from the marked `XE` fields; page numbers are cached
  as `?` and an **`index_cached`** warning explains Word computes them on
  open/refresh. Re-running `add type:index` refreshes the block.
- `read --view structure` lists each index with its column count and path.

## Mail-merge field — `add type:mergeField`

Inserts a `MERGEFIELD` complex field the `template` verb fills by name:

    aioffice edit r.docx --ops '[{"op":"add","path":"/body/p[2]","type":"mergeField","props":{"name":"Name"}}]'

- `props.name` (required): the merge-field name (a Word field-name token:
  letters, digits, `_`). The field shows `«Name»` until merged.
- `props.find` (optional): place it next to specific text instead of appending.
- The `template` verb fills both `MERGEFIELD` fields (by name) **and** `{{key}}`
  text placeholders from the same `--data` map:

      aioffice template r.docx --data '{"Name":"Acme"}' -o out.docx

  Unresolved field names join `data.leftoverPlaceholders` in the result.
- `read --view fields` lists merge fields alongside content controls.

## Notes

- All four are docx-only. xlsx/pptx return `unsupported_feature` naming the
  workaround.
- Caption authoring (`add type:caption`) feeds the table of figures; see
  `aioffice help properties-docx`.
