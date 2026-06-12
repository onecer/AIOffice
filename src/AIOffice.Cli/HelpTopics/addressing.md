# Addressing grammar

Paths are slash-separated, indices are 1-based, and `query` returns canonical
paths that `get` and `edit` accept unchanged.

## docx

    /body/p[3]                  3rd paragraph of the body
    /body/p[3]/run[2]           2nd run inside that paragraph
    /body/table[1]/tr[2]/tc[1]  table 1, row 2, cell 1
    /header[1]/p[1]             1st paragraph of the 1st header

## xlsx

    /Sheet1/A1                  one cell
    /Sheet1/A1:C10              rectangular range
    /Sheet1/row[3]              whole row 3
    /'Q3 Data'/B2               sheet names with spaces use single quotes
                                (escape a quote by doubling it: 'O''Brien')

## pptx

    /slide[2]                   2nd slide
    /slide[2]/shape[3]          3rd shape on it
    /slide[2]/shape[3]/p[1]     1st paragraph inside that shape

Master/layout addressing is reserved for M1 and currently answers
`unsupported_feature`.

## Rules

- A path always starts with `/` and never ends with one.
- Element segments are `name` or `name[i]` with `i >= 1`.
- Cell/range segments use absolute A1 notation (column letters + row number).
- Quoted segments (`'Q3 Data'`) are name lookups, used for sheet names.
- When a path does not resolve, the error is `invalid_path` and
  `error.candidates` lists the nearest existing paths — pick one and retry.
