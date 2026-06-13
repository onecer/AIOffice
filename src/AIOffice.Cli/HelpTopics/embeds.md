# embeds — embed, list and extract files (M10)

A document can carry **another file** as an embedded OLE/package object: a source
`.xlsx` attached to a report, a `.pdf`, a `.zip`, anything. aioffice can embed
one, list the embeds a document holds, and extract the payload back out
**byte-identical**. Works in all three formats (docx, xlsx, pptx).

## Embed a file (add type:embed)

    aioffice edit report.docx --add /body --type embed src=data.xlsx name="Q3 model"

- `src` (required) is a workspace-relative file of any kind — sandbox-resolved,
  so a path escaping the workspace is `sandbox_denied`.
- `name` (optional) is the display name; defaults to the source file name.
- `icon` (optional) is a PNG/JPEG shown for the object (a neutral placeholder
  otherwise).
- The **container** is per-format:
  - docx → `/body` (or a table cell / header / footer)
  - xlsx → `/Sheet1` (anchored on the sheet)
  - pptx → `/slide[i]`
- The media type is sniffed from the **source file**; the op returns the
  canonical embed path.

## List the embeds (read --view embeds)

    aioffice read report.docx --view embeds

Returns every embed's `{path, name, mediaType, size, container}`. The
`structure` view includes embeds too. `get <embed path>` returns the same
metadata — **not** the bytes.

## Extract the payload (the extract op)

    aioffice edit report.docx --ops '[{"op":"extract","path":"/embed[1]","props":{"to":"out/data.xlsx"}}]'

- `extract` is a **producing** op: it writes `props.to` (sandbox-resolved) and
  does **not** modify the source document.
- The extracted bytes equal what was embedded, byte-for-byte — even after the
  document has been opened and saved again.

## Addressing

- docx: `/embed[i]` (1-based, document order; header/footer embeds included)
- xlsx: `/Sheet1/embed[i]` (1-based per sheet)
- pptx: `/slide[i]/embed[j]` or `/slide[i]/embed[@id=N]` (the host frame's id)

## Remove

    aioffice edit report.docx --remove /embed[1]

Removes the object and its backing parts. Validates clean afterward.

## Notes

- The embed source and the extract destination are both sandbox-resolved; an
  escaping path is `sandbox_denied` and nothing is read or written.
- The MCP surface uses the same ops on `office_edit` / `office_read` — no new
  tool was added (still 17).
