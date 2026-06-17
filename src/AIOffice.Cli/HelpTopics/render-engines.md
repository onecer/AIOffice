# render engines (1.9) — `render --engine chromium|soffice|auto`

`render --to png|pdf` can use two different engines. The default is **chromium**;
**soffice** is an optional TRUE-fidelity engine. The `--engine` flag is additive —
omit it (or pass `chromium`) and every existing render is byte-for-byte unchanged.

## chromium (default)

Screenshots / prints aioffice's own HTML (docx/xlsx) or SVG (pptx) projection with
a headless Chromium (Chrome/Edge/Chromium). Fast, no Office install, but the layout
is aioffice's reconstruction, not Office's. Needs a Chromium browser on PATH (or
`AIOFFICE_BROWSER`). This is the engine used when `--engine` is omitted.

```
aioffice render deck.pptx --to png --scope /slide[3] -o slide3.png
aioffice render report.docx --to pdf -o report.pdf
```

## soffice (true fidelity, optional)

Hands the ORIGINAL .docx/.xlsx/.pptx to a headless LibreOffice, so the output
matches Office's own layout closely (the high-fidelity path).

- `--to pdf` → `soffice --headless --convert-to pdf` (the WHOLE document; `--scope`
  is reported back but does not narrow the PDF — use `--to png` for one page).
- `--to png` → soffice to PDF, then `pdftoppm` rasterizes the selected page
  (pptx: the slide from `--scope /slide[N]`, default 1; docx/xlsx: page 1).
  PNG therefore needs **pdftoppm** from poppler as well as LibreOffice.
- `--to svg|html|text` → not supported by soffice; falls back to the native/chromium
  engine and emits an `engine_fallback` warning. soffice does png+pdf only.

```
aioffice render deck.pptx --to pdf  --engine soffice -o deck.pdf
aioffice render deck.pptx --to png  --engine soffice --scope /slide[2] -o s2.png
```

Requirements:
- LibreOffice — macOS: the `LibreOffice.app` bundle; linux: the `libreoffice`
  package; windows: the LibreOffice installer. Override with `AIOFFICE_SOFFICE`.
- poppler `pdftoppm` (PNG only) — macOS/linux: `brew install poppler` or your
  package manager's `poppler-utils`; windows: a poppler build on PATH. Override
  with `AIOFFICE_PDFTOPPM`.

If `--engine soffice` is chosen explicitly but LibreOffice is missing, the render
fails with a clear install hint (no silent fallback). For PNG, a missing pdftoppm
is an `unsupported_feature` error suggesting `brew install poppler` — use `--to pdf`
for a soffice render without poppler, or `--engine chromium` for a screenshot PNG.

## auto

`--engine auto` picks **soffice when `doctor` finds LibreOffice on this machine,
else chromium**. Use it when you want fidelity where available without hard-failing
on machines that lack LibreOffice.

## checking what's installed

`aioffice doctor` reports a `renderers` object:

```
"renderers": {
  "chromium":    { "engine": "chromium", "found": true,  "path": "…", "kind": "chrome" },
  "libreoffice": { "engine": "soffice",  "found": true,  "path": "/Applications/LibreOffice.app/Contents/MacOS/soffice" },
  "poppler":     { "tool": "pdftoppm",   "found": true,  "path": "/opt/homebrew/bin/pdftoppm" }
}
```

`found:false` for libreoffice means `--engine soffice` will error (and `auto`
falls back to chromium); `found:false` for poppler means soffice `--to png` will
error while soffice `--to pdf` still works.
