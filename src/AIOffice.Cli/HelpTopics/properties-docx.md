# docx properties (v0)

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

## table / tr / tc

| prop      | type   | notes                                  |
|-----------|--------|----------------------------------------|
| rows      | number | `add` only: initial row count          |
| cols      | number | `add` only: initial column count       |
| text      | string | on `tc`: cell text                     |
| style     | string | named table style                      |

## add positions

`{op:"add", path:"/body", type:"p", position:"inside"}` appends to the body;
`position:"before"|"after"` inserts relative to the node at `path`.

Headers/footers are addressable (`/header[1]/p[1]`) for read/set; creating new
header parts is M1.
