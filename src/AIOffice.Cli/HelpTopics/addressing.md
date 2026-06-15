# Addressing grammar

Paths are slash-separated, indices are 1-based, and `query` returns canonical
paths that `get` and `edit` accept unchanged.

## docx

    /body/p[3]                  3rd paragraph of the body
    /body/p[3]/run[2]           2nd run inside that paragraph
    /body/table[1]/tr[2]/tc[1]  table 1, row 2, cell 1
    /header[1]/p[1]             1st paragraph of the 1st header
    /revision[@id=3]            tracked revision by id (read --view revisions lists ids;
                                /revision[1] positional form also works)
    /comment[@id=2]             comment by id (read --view comments lists ids)
    /styles                     the style library (add type:style targets it)
    /style[@id=Callout]         one style definition by id (set/get/remove)
    /section[1]                 page setup + columns (set/get)
    /body/p[3]/omath[1]         M6: an inline equation in a paragraph (get/remove)
    /equation[@num=1.1]         1.7: a numbered display equation by its number (matches a
                                "(1.1)" label by its numeric core); get reports the number
    /caption[@label=Figure][1]  M8: a Figure/Table/Equation caption (1-based per label; get)
    /embed[1]                   M10: an embedded OLE/package object (read --view embeds lists
                                them; get/remove; extract op writes the bytes out)
    /properties                 core + custom document properties (get; set /properties)

## xlsx

    /Sheet1/A1                  one cell
    /Sheet1/A1:C10              rectangular range
    /Sheet1/row[3]              whole row 3
    /Sheet1/col[C]              whole column C
    /Sheet1/row[2]:row[6]       M6: a row span (outline group: add type:group / remove)
    /Sheet1/col[B]:col[E]       M6: a column span (outline group)
    /Sheet1/table[@name=Sales]  M6: an Excel Table (ListObject) by name
    /Sheet1/chart[1]            1st chart anchored on the sheet
    /Pivot/pivot[1]             1st pivot table ON ITS TARGET SHEET
    /Pivot/pivot[@name=SalesPivot]  same pivot by name (canonical form in results)
    /Sheet1/conditionalFormat[1]    1st conditional-format rule on the sheet
    /Sheet1/image[1]            1st anchored picture on the sheet
    /Sheet1/linkedPicture[1]    1.7: 1st linked picture (camera tool) on the sheet
                                (a static snapshot of a cell range; get/remove)
    /Sheet1/slicer[1]           M8: 1st slicer on the sheet (table-column / pivot-field)
    /Sheet1/slicer[@name=X]     M8: the same slicer by name (canonical form in results)
    /Sheet1/embed[1]            M10: an embedded OLE/package object on the sheet
                                (read --view embeds lists them; get/remove; extract op)
    /'Q3 Data'/B2               sheet names with spaces use single quotes
                                (escape a quote by doubling it: 'O''Brien')
    /properties                 core + custom workbook properties (get; set /properties)

## pptx

    /                           M6: the presentation root — slide size + sections (set/get)
    /slide[2]                   2nd slide
    /slide[2]/shape[3]          3rd shape on it
    /slide[2]/shape[@id=7]      same shape by stable id (canonical form in results)
    /slide[2]/shape[3]/p[1]     1st paragraph inside that shape
    /slide[2]/notes             the slide's speaker notes (whole body; no /p[j] beneath)
    /slide[2]/shape[@id=9]/omath[1]  M10: an equation inside a slide text box
                                (get reports the LaTeX; add type:equation creates it)
    /slide[2]/embed[@id=7]      M10: an embedded OLE/package object on the slide, by id
                                (read --view embeds lists them; get/remove; extract op)
    /slide[2]/animation[1]      M6: an animation in the slide timeline (set/move/remove)
    /section[1]                 M6: a slide section (set name / get / remove)
    /master[1]                  1st slide master (M6: editable — background, accents)
    /master[1]/layout[2]        2nd layout under that master (M6: editable; add clones)
    /master[1]/shape[1]         shape on the master (M6: editable — set/add/remove)
    /notesMaster                1.7: the notes master (set background/bodyFont; get)
    /handoutMaster              1.7: the handout master (set headerFooter/slidesPerPage; get)
    /properties                 core + custom presentation properties (get; set /properties)

Masters, layouts and their shapes are editable since M6 (background, theme
accent colors, shapes, cloned layouts).

## Rules

- A path always starts with `/`. `/` alone is the document root (pptx slide
  size / sections; document-level `get`); every other path has at least one
  segment and never ends with `/`.
- Element segments are `name` or `name[i]` with `i >= 1` (or `col[C]` letters).
- An xlsx outline group is a span segment: `row[a]:row[b]` / `col[a]:col[b]`.
- Cell/range segments use absolute A1 notation (column letters + row number).
- Quoted segments (`'Q3 Data'`) are name lookups, used for sheet names.
- When a path does not resolve, the error is `invalid_path` and
  `error.candidates` lists the nearest existing paths — pick one and retry.
