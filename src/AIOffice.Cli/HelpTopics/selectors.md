# Selector grammar (aioffice query)

    element[attr OP value]:contains('text')

- `element` — a node type (`p`, `table`, `cell`, `row`, `slide`, `shape`, `run`)
  or `*` for any.
- `[attr OP value]` — zero or more attribute predicates. Operators:
  `=` `!=` `>` `>=` `<` `<=` (numeric when both sides are numbers) and
  `*=` (substring).
- `:contains('text')` — the node's text contains `text`. Use doubled quotes to
  escape: `:contains('it''s')`.
- Values may be bare (`Heading1`), single- or double-quoted.

## Examples

    p[style=Heading1]            docx paragraphs styled Heading1
    p:contains('Q3')             paragraphs mentioning Q3
    cell[value>100]              xlsx cells with numeric value over 100
    cell[formula*=SUM]           cells whose formula contains SUM
    shape:contains('Revenue')    pptx shapes whose text mentions Revenue
    *[name*=Total]               any element whose name contains Total

`query` returns canonical paths (see `aioffice help addressing`) plus a small
property preview per match. Feed the paths straight into `get` or `edit`.
