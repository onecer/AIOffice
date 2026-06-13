# docx properties

Properties appear in `get` output and are written via `edit` ops
(`set` props, `add` type+props). Unknown properties fail with `invalid_args`
listing the supported set; capabilities not built yet answer
`unsupported_feature` with a workaround.

## p (paragraph)

| prop      | type   | notes                                            |
|-----------|--------|--------------------------------------------------|
| text      | string | plain text; replaces all runs on `set`           |
| style     | string | named paragraph style, e.g. Heading1, Normal     |
| align     | string | left · center · right · justify                  |
| rtl       | bool   | M6: right-to-left flow (`w:bidi`); turning it on also right-aligns the paragraph. `get` reports `rtl` |

## run

| prop      | type   | notes                       |
|-----------|--------|-----------------------------|
| text      | string | the run text                |
| bold      | bool   |                             |
| italic    | bool   |                             |
| underline | bool   |                             |
| size      | number | points                      |
| color     | string | hex RGB, e.g. FF0000        |
| font      | string | font family name            |
| rtl       | bool   | M6: right-to-left mark (`w:rtl`) for mixed-direction text |

## table / tr / tc

| prop      | type   | notes                                  |
|-----------|--------|----------------------------------------|
| rows      | number | `add` only: initial row count          |
| cols      | number | `add` only: initial column count       |
| text      | string | on `tc`: cell text                     |
| style     | string | named table style                      |

Deep tables (M5) — `set` on the table:

| prop          | type     | notes                                        |
|---------------|----------|----------------------------------------------|
| borders       | string   | all · outer · none (+ borderColor, borderWidthPt) |
| shading       | string   | hex RGB fill                                 |
| headerRow     | bool     | bolds row 1 + repeats it on every page       |
| width         | length   | `"12cm"` or `"100%"`                        |
| columnWidths  | array    | one per column, e.g. `["3cm","auto","2.5cm"]`|
| alignment     | string   | left · center · right                       |
| cellPaddingCm | number   | uniform cell padding                        |
| rtl           | bool     | M6: right-to-left table (`w:bidiVisual` — mirrors column order). `get` reports `rtl` |

and `set` on a cell (`/body/table[1]/tr[1]/tc[1]`): `mergeRight: N`
(gridSpan), `mergeDown: N` (vMerge chain; 1 unmerges), `shading`, `valign:
top|center|bottom`. Merged-away cells disappear from `get`/`render`;
`render --to html` emits real `colspan`/`rowspan`.

## add positions

`{op:"add", path:"/body", type:"p", position:"inside"}` appends to the body;
`position:"before"|"after"` inserts relative to the node at `path`.

Headers/footers are addressable (`/header[1]/p[1]`) for read/set, and created
with `{op:"add", path:"/header[1]", type:"header", props:{text, align?}}`
(same for `footer`). M5 variants: `/header[firstPage]` and `/header[even]`
(same for footers) — adding them wires `w:titlePg` / `w:evenAndOddHeaders`
automatically, all three variants coexist, and `get` reports the variant.

## field (M5)

| prop        | type   | notes                                          |
|-------------|--------|------------------------------------------------|
| kind        | string | pageNumber (PAGE) · numPages (NUMPAGES) · date (DATE) · docTitle (TITLE) |
| format      | string | date only: a format picture, e.g. `"yyyy"`     |
| leadingText | string | literal text emitted before the field          |

`'Page X of Y'` footer: add a `pageNumber` field to `/footer[1]/p[1]`, then a
`numPages` field with `leadingText:" of "`.

## markdown bridge (M5)

`aioffice create report.docx --from notes.md` imports GFM markdown: headings →
Heading styles, bold/italic/strike/code runs, nested bullet/number lists, pipe
tables (header row bolded), links, images (sandbox-resolved), blockquotes,
code blocks, horizontal rules; raw HTML and missing images degrade to
`meta.warnings`. `aioffice read report.docx --view markdown` exports the body
back as GFM — structure round-trips with the importer.

## list (M3)

`add` a paragraph with list props to make it a list item:

| prop        | type   | notes                                          |
|-------------|--------|------------------------------------------------|
| list        | string | bullet · number                                |
| level       | number | 0-based nesting level (default 0)              |
| listRestart | bool   | number lists: restart numbering at this item   |

`{op:"add", path:"/body", type:"p", props:{text:"Step one", list:"number"}}`;
nested: `props:{text:"Detail", list:"number", level:1}`. `read --view text`
shows `1.` / `•` markers; `render --to html` emits real `<ol>`/`<ul>`.

## link / bookmark / footnote (M3)

| type     | props                       | notes                                        |
|----------|-----------------------------|----------------------------------------------|
| link     | text, url OR anchor        | inline at the end of the target paragraph    |
| bookmark | name                       | wraps the target paragraph; name ≤ 40 chars  |
| footnote | text                       | reference at the paragraph end + note text   |

`{op:"add", path:"/body/p[3]", type:"link", props:{text:"site", url:"https://…"}}`;
internal jump: `props:{text:"see intro", anchor:"Intro"}` (a bookmark name).

## section (M3 page setup + M6 columns, `/section[1]`)

| prop                       | type   | notes                              |
|----------------------------|--------|------------------------------------|
| pageSize                   | string | A4 · Letter · Legal · A3 · A5      |
| orientation                | string | portrait · landscape               |
| marginTop/Bottom/Left/Right| length | e.g. "2cm", "72pt"                 |
| columns                    | number | M6: newspaper-style column count (1 clears) |
| columnGap                  | length | M6: space between columns, e.g. "1.25cm" |

`{op:"set", path:"/section[1]", props:{pageSize:"A4", orientation:"landscape"}}`;
two-column layout: `{op:"set", path:"/section[1]", props:{columns:2}}`.
`get /section[1]` reflects everything (sizes in cm, incl. `columns`,
`columnGapCm`). Inserting NEW section breaks: `{op:"add", type:"sectionBreak"}`.

## equation (M6, docx LaTeX → Office Math)

`add` an equation with a LaTeX `latex` prop; `display:true` makes a centered
equation block, `display:false` (default) an inline one.

| prop    | type   | notes                                                  |
|---------|--------|--------------------------------------------------------|
| latex   | string | required; the LaTeX source (stored for faithful read-back) |
| display | bool   | true = centered block (`m:oMathPara`); false = inline (default) |

    aioffice edit r.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"E = mc^2"}}]'
    aioffice edit r.docx --ops '[{"op":"add","path":"/body/p[1]","type":"equation","props":{"latex":"\\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}","display":true}}]'

Inline equations are addressed `/body/p[i]/omath[j]`; `get` returns the stored
`latex` and `display`. `read --view text` shows `$…$` (inline) / `$$…$$`
(display) markers. Unrecognized LaTeX commands degrade to literal runs and
raise an `equation_partial` warning (the file still validates). See
`aioffice help equations` for the supported LaTeX subset.

## columnBreak (M6)

`{op:"add", path:"/body/p[3]", type:"columnBreak"}` inserts a `w:br
w:type="column"` so the following text flows into the next column of a
multi-column section.

## tracked changes (M2)

Pass `--track` (and optionally `--author NAME`) on `edit`: text `set`/`add`/
`remove` ops are recorded as `w:ins`/`w:del` revisions instead of being applied
silently. Only text changes can be tracked; tracked formatting answers
`invalid_args` naming the workaround.

    aioffice edit r.docx --track --author "Reviewer" --set /body/p[1] text='New wording'
    aioffice read r.docx --view revisions     # -> [{id, kind, author, date, text, path}]
    aioffice edit r.docx --ops '[{"op":"accept","path":"/revision[@id=3]"}]'
    aioffice edit r.docx --ops '[{"op":"reject","path":"/body"}]'   # scope form: all in /body

Author resolution: op `props.author` > `--author` > `AIOFFICE_AUTHOR` env >
`AIOffice`. Revisions are never `remove`d — only accepted or rejected.
M3: `accept`/`reject` also resolve formatting revisions (w:rPrChange /
w:pPrChange) — reject restores the preserved old formatting. *Writing*
tracked formatting is still `invalid_args` with the workaround named.

## comment (M2)

| prop   | type   | notes                                                |
|--------|--------|------------------------------------------------------|
| text   | string | required                                             |
| author | string | optional; defaults like tracked changes              |

`{op:"add", path:"/body/p[2]", type:"comment", props:{text:"check this"}}` —
the path is the anchored content (paragraph or run). Read back with
`read --view comments`; `get /comment[@id=2]`; `remove /comment[@id=2]`.
M3 threaded replies: `{op:"add", path:"/comment[@id=1]", type:"reply",
props:{text:"agreed", author?:"Name"}}` — `read --view comments` nests them
under `replies`.

## style (M2)

| prop          | type   | notes                                         |
|---------------|--------|-----------------------------------------------|
| id            | string | required on `add`; letters/digits/_/-         |
| kind          | string | paragraph (default) · character               |
| name, basedOn | string | display name; parent style id                 |
| bold/italic/underline | bool |                                        |
| color         | string | hex RGB                                       |
| fontSize      | number | points                                        |
| alignment     | string | left · center · right · justify (paragraph)   |
| spacingBefore/spacingAfter | number | points (paragraph)               |

`{op:"add", path:"/styles", type:"style", props:{id:"Callout", bold:true}}` defines;
`{op:"set", path:"/style[@id=Callout]", props:{color:"FF0000"}}` modifies;
apply with `{op:"set", path:"/body/p[2]", props:{style:"Callout"}}`.
`read --view styles` lists all; `remove` works on custom styles only.

## image (M2)

| prop          | type   | notes                                          |
|---------------|--------|------------------------------------------------|
| src           | string | required; PNG/JPEG path inside the workspace   |
| width, height | length | e.g. "10cm", "72pt"; omit either to keep aspect|
| alt (M7)      | string | accessibility description (`wp:docPr/@descr`); also settable on the image-carrying paragraph/run with `set {alt}` |

`{op:"add", path:"/body", type:"image", props:{src:"logo.png", width:"10cm", alt:"Company logo"}}`
appends; `{op:"add", path:"/body/p[1]", type:"image", position:"before"}`
places it relative to a paragraph. An `src` escaping the workspace sandbox
fails with `sandbox_denied`. The audit verb flags images with no `alt` as
`a11y_no_alt_text` and `--fix` inserts a placeholder.

## document properties (M7, `/properties`)

Core + custom package metadata live on a virtual `/properties` node. Read with
`get /properties` or `read --view properties`; write with `set /properties`.

| prop        | type   | notes                                                       |
|-------------|--------|-------------------------------------------------------------|
| title       | string | core Title (also what `audit`'s `a11y_no_doc_title` checks) |
| subject · author · keywords · category · comments · lastModifiedBy · revision | string | core properties (`author` = Creator) |
| created · modified | date | ISO 8601, e.g. "2026-06-13" or "2026-06-13T10:00:00Z" |
| custom      | object | typed custom props: string/number/bool/date round-trip with their JSON type |

    {op:"set", path:"/properties", props:{title:"Q3 Report", author:"AIOffice",
      custom:{Project:"Aurora", Reviewed:true, Budget:1000}}}

`read --view properties` returns `{core:{…}, custom:{…}}`; a custom value that
parses as a number/bool/ISO-date is stored as `vt:r8`/`vt:bool`/`vt:filetime`
and reads back with that type. A `null` custom value clears the property.

## content controls (M7, `add type:contentControl`, `read --view fields`)

Block-level structured document tags (`w:sdt`) for template fields. Addressed by
`/sdt[@tag=X]` (or positionally `/sdt[i]`); listed by `read --view fields`.

| prop  | type   | notes                                                         |
|-------|--------|---------------------------------------------------------------|
| kind  | string | text · dropdown · date · checkbox (default text)              |
| tag   | string | required; unique stable identifier                            |
| title | string | optional display title (`w:alias`)                            |
| items | array  | dropdown only: the choices, e.g. `["Draft","Final"]`          |
| text  | string | initial value (dropdown: must be one of `items`; date: ISO)   |
| checked | bool | checkbox only                                                 |

    {op:"add", path:"/body/p[2]", type:"contentControl",
      props:{kind:"dropdown", tag:"status", title:"Status", items:["Draft","Final"]}}
    {op:"set", path:"/sdt[@tag=status]", props:{text:"Final"}}    # writes the value
    {op:"set", path:"/sdt[@tag=signed]", props:{checked:true}}    # checkbox

`read --view fields` returns each control's `{path, kind, tag, title, value, items?}`.
`remove /sdt[@tag=X]` unwraps the control but keeps its content (`keepContent:false`
drops it whole). Position is `before`/`after` a paragraph (default `after`).
