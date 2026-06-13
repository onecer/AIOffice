# rtl (M6 — right-to-left / bidirectional text)

aioffice supports right-to-left flow for Arabic, Hebrew and other bidi scripts
at three docx levels. RTL is a boolean prop; `get` reports it back.

## Paragraph

`{op:"set", path:"/body/p[5]", props:{rtl:true}}` sets `w:bidi` so the paragraph
flows right-to-left; aioffice also right-aligns it (the usual expectation for an
RTL paragraph). Clear with `rtl:false`.

    aioffice edit r.docx --ops '[{"op":"add","path":"/body","type":"p","props":{"text":"مرحبا بالعالم"}},{"op":"set","path":"/body/p[5]","props":{"rtl":true}}]'
    aioffice get r.docx /body/p[5]      # -> "rtl": true, "alignment": "right"

## Run

`{op:"set", path:"/body/p[5]/run[2]", props:{rtl:true}}` sets `w:rtl` on a single
run — use it for a right-to-left span embedded in an otherwise left-to-right
paragraph (mixed-direction text).

## Table

`{op:"set", path:"/body/table[1]", props:{rtl:true}}` sets `w:bidiVisual`, which
mirrors the visual column order (column 1 renders on the right). `get
/body/table[1]` reports `rtl`.

## Notes

- RTL props are docx-only in M6.
- The flag is independent of the text content: aioffice never reshapes glyphs;
  it only sets the OOXML direction so Word/LibreOffice lay the text out RTL.
- Combine with paragraph `align` to override the auto right-alignment if needed.
