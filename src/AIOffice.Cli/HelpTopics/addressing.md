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
    /Sheet1/chart[1]            1st chart anchored on the sheet
    /'Q3 Data'/B2               sheet names with spaces use single quotes
                                (escape a quote by doubling it: 'O''Brien')

## pptx

    /slide[2]                   2nd slide
    /slide[2]/shape[3]          3rd shape on it
    /slide[2]/shape[@id=7]      same shape by stable id (canonical form in results)
    /slide[2]/shape[3]/p[1]     1st paragraph inside that shape
    /master[1]                  1st slide master (read-only: get/query)
    /master[1]/layout[2]        2nd layout under that master (read-only)
    /master[1]/shape[1]         shape on the master (read-only)

Masters and layouts are readable (`get`, `query`); editing them and
paragraph/run addressing beneath them stay `unsupported_feature` until M2.

## Rules

- A path always starts with `/` and never ends with one.
- Element segments are `name` or `name[i]` with `i >= 1`.
- Cell/range segments use absolute A1 notation (column letters + row number).
- Quoted segments (`'Q3 Data'`) are name lookups, used for sheet names.
- When a path does not resolve, the error is `invalid_path` and
  `error.candidates` lists the nearest existing paths — pick one and retry.
