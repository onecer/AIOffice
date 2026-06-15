#!/usr/bin/env bash
# =============================================================================
# examples/tour.sh — regenerate the entire AIOffice showcase from scratch.
#
# Builds the three artifacts shown in SHOWCASE.md — a dark pitch deck (.pptx),
# a regional revenue dashboard (.xlsx), and a capability report (.docx) — and
# renders the PNGs you see in the gallery. Every build line below is a real
# aioffice command; there is no Microsoft Office, no template, no manual touch-up.
#
# The committed gallery PNGs were rendered with LibreOffice (headless `soffice`),
# which rasterizes the real OOXML exactly as a desktop Office app shows it. If
# `soffice` + `pdftoppm` are present this script reproduces those images; if not,
# it falls back to aioffice's own `render --to png` (a fast, approximate preview).
#
# Requirements:
#   * `aioffice` on your PATH               (see ../docs/INSTALL.md)
#   * LibreOffice (`soffice`) + `pdftoppm`  (for the high-fidelity gallery PNGs)
#   * Chrome / Chromium                     (only for the aioffice render fallback)
#
# Usage:
#   ./examples/tour.sh                 # build into a fresh temp dir, render PNGs
#   OUTDIR=./my-showcase ./tour.sh     # build into a directory you keep
#
# The script is the proof: read it top to bottom and you have seen every
# command that produced the gallery.
# =============================================================================

set -euo pipefail

# --- setup -------------------------------------------------------------------
AIO="${AIO:-aioffice}"
if ! command -v "$AIO" >/dev/null 2>&1; then
  echo "error: '$AIO' not found on PATH. Install it first (see docs/INSTALL.md)," >&2
  echo "       or run with AIO=/path/to/aioffice ./examples/tour.sh" >&2
  exit 127
fi

WS="${OUTDIR:-$(mktemp -d "${TMPDIR:-/tmp}/aioffice-tour.XXXXXX")}"
mkdir -p "$WS"
WS="$(cd "$WS" && pwd)"   # absolute — aioffice sandboxes to --workspace
echo "==> Building the AIOffice showcase in: $WS"
echo "==> Using binary: $($AIO version | sed 's/.*"version":"//; s/".*//' 2>/dev/null || echo "$AIO")"
echo

run() { echo "  \$ $*"; "$@" >/dev/null; }

# --- rendering ---------------------------------------------------------------
# Prefer LibreOffice (real-Office fidelity) for the gallery PNGs; fall back to
# aioffice's own renderer when soffice/pdftoppm aren't installed.
SOFFICE="$(command -v soffice 2>/dev/null || true)"
[ -z "$SOFFICE" ] && [ -x /Applications/LibreOffice.app/Contents/MacOS/soffice ] \
  && SOFFICE=/Applications/LibreOffice.app/Contents/MacOS/soffice
have_lo() { [ -n "$SOFFICE" ] && [ -x "$SOFFICE" ] && command -v pdftoppm >/dev/null 2>&1; }
_lo_pdf() {  # <file> -> echoes the produced pdf path
  "$SOFFICE" --headless --convert-to pdf --outdir "$WS" "$1" >/dev/null 2>&1
  echo "$WS/$(basename "${1%.*}").pdf"
}

# render_pages <file> <out-prefix> [maxPages]  ->  out-prefix-1.png … (one per slide/page)
render_pages() {
  local file="$1" prefix="$2" max="${3:-0}" i
  if have_lo; then
    echo "  \$ soffice --convert-to pdf $(basename "$file") && pdftoppm -> $(basename "$prefix")-N.png"
    local pdf; pdf="$(_lo_pdf "$file")"
    pdftoppm -png -scale-to-x 1280 -scale-to-y -1 "$pdf" "$prefix" >/dev/null 2>&1
    # pdftoppm zero-pads multi-digit page counts; normalise to prefix-1.png …
    for f in "$prefix"-0*.png; do [ -e "$f" ] && mv "$f" "${f/-0/-}"; done 2>/dev/null || true
  else
    [ "$max" -gt 0 ] || max=6
    for ((i=1;i<=max;i++)); do
      run "$AIO" render "$file" --workspace "$WS" --to png --scope "/slide[$i]" -o "$prefix-$i.png"
    done
  fi
}

# render_page1 <file> <out.png> [aioffice-scope]  ->  a single page-1 PNG
render_page1() {
  local file="$1" out="$2" scope="${3:-}"
  if have_lo; then
    echo "  \$ soffice --convert-to pdf $(basename "$file") && pdftoppm -f1 -l1 -> $(basename "$out")"
    local pdf; pdf="$(_lo_pdf "$file")"
    pdftoppm -png -scale-to-x 1400 -scale-to-y -1 -f 1 -l 1 "$pdf" "${out%.png}-p" >/dev/null 2>&1
    mv "${out%.png}-p-1.png" "$out" 2>/dev/null || mv "${out%.png}-p"-*.png "$out"
  elif [ -n "$scope" ]; then
    run "$AIO" render "$file" --workspace "$WS" --to png --scope "$scope" -o "$out"
  else
    run "$AIO" render "$file" --workspace "$WS" --to png -o "$out"
  fi
}

# =============================================================================
# 1) deck.pptx — a dark, 16:9 product pitch deck (6 slides, native chart)
# =============================================================================
echo "==> [1/3] deck.pptx — product pitch deck (dark theme, native chart)"

# create the deck, set 16:9 size, and the dark master theme
run "$AIO" create "$WS/deck.pptx" --kind pptx --workspace "$WS"
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"set","path":"/","props":{"slideSize":"16:9"}},
  {"op":"set","path":"/master[1]","props":{"background":"0F172A","accent1":"38BDF8","accent2":"818CF8","accent3":"34D399"}}
]'

# grow to 6 slides, all on the dark background
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"set","path":"/slide[1]","props":{"background":"0F172A"}},
  {"op":"add","path":"/slide[1]","type":"slide","position":"after","props":{"background":"0F172A"}},
  {"op":"add","path":"/slide[2]","type":"slide","position":"after","props":{"background":"0F172A"}},
  {"op":"add","path":"/slide[3]","type":"slide","position":"after","props":{"background":"0F172A"}},
  {"op":"add","path":"/slide[4]","type":"slide","position":"after","props":{"background":"0F172A"}},
  {"op":"add","path":"/slide[5]","type":"slide","position":"after","props":{"background":"0F172A"}}
]'

# slide 1: title (wordmark, tagline, accent shapes, footer)
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"rect","x":3.0,"y":5.6,"w":7.2,"h":0.16,"fill":"38BDF8","name":"accentBar"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"rect","x":29.6,"y":13.4,"w":4.27,"h":3.2,"fill":"38BDF8","name":"cornerBlock"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"shape":"rect","x":28.0,"y":2.4,"w":3.4,"h":3.4,"fill":"1E293B","name":"ghostBlock"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"text":"AIOffice","x":2.85,"y":6.45,"w":20.0,"h":2.6,"fontSize":86,"bold":true,"color":"F8FAFC","align":"left","name":"wordmark"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"text":"An AI-native Office engine for agents","x":3.0,"y":9.7,"w":20.0,"h":1.1,"fontSize":26,"color":"94A3B8","align":"left","name":"tagline"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"text":"One self-contained binary  ·  docx · xlsx · pptx  ·  CLI + MCP","x":3.0,"y":11.0,"w":20.0,"h":0.8,"fontSize":15,"color":"475569","align":"left","name":"subtag"}},
  {"op":"add","path":"/slide[1]","type":"shape","props":{"text":"Every slide in this deck was built only by aioffice — no Office installed.","x":3.0,"y":17.8,"w":20.0,"h":0.7,"fontSize":12,"color":"475569","align":"left","name":"footer"}}
]'

# slide 2: "One binary, three formats" — 3 format cards
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[2]","type":"shape","props":{"shape":"rect","x":2.0,"y":2.1,"w":3.6,"h":0.14,"fill":"38BDF8","name":"hdrBar"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"One binary, three formats","x":2.0,"y":2.5,"w":26.0,"h":1.6,"fontSize":34,"bold":true,"color":"F8FAFC","name":"hdr"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"Read, edit, render and validate real Office files — no Microsoft runtime, no cloud.","x":2.0,"y":4.2,"w":29.0,"h":1.0,"fontSize":15,"color":"94A3B8","name":"sub"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"shape":"roundRect","x":2.0,"y":6.6,"w":9.4,"h":9.4,"fill":"1E293B","name":"card1"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"shape":"roundRect","x":12.2,"y":6.6,"w":9.4,"h":9.4,"fill":"1E293B","name":"card2"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"shape":"roundRect","x":22.4,"y":6.6,"w":9.4,"h":9.4,"fill":"1E293B","name":"card3"}}
]'
# card accent bars, color-coded format names, titles and bodies (manual \n line breaks)
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[2]","type":"shape","props":{"shape":"rect","x":2.7,"y":7.4,"w":1.6,"h":0.12,"fill":"38BDF8","name":"c1bar"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"docx","x":2.7,"y":7.9,"w":8.0,"h":1.5,"fontSize":40,"bold":true,"color":"38BDF8","name":"c1num"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"Word documents","x":2.7,"y":9.9,"w":8.0,"h":0.9,"fontSize":17,"bold":true,"color":"F8FAFC","name":"c1t"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"Headings, tables, styles,\ntracked changes, comments\nand mail-merge — straight\nto clean OOXML.","x":2.7,"y":11.0,"w":8.0,"h":3.5,"fontSize":13,"color":"94A3B8","name":"c1b"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"shape":"rect","x":12.9,"y":7.4,"w":1.6,"h":0.12,"fill":"818CF8","name":"c2bar"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"xlsx","x":12.9,"y":7.9,"w":8.0,"h":1.5,"fontSize":40,"bold":true,"color":"818CF8","name":"c2num"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"Spreadsheets","x":12.9,"y":9.9,"w":8.0,"h":0.9,"fontSize":17,"bold":true,"color":"F8FAFC","name":"c2t"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"Real formulas, dynamic\narrays, charts, pivots,\nconditional formats,\nwhat-if & goal-seek.","x":12.9,"y":11.0,"w":8.0,"h":3.5,"fontSize":13,"color":"94A3B8","name":"c2b"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"shape":"rect","x":23.1,"y":7.4,"w":1.6,"h":0.12,"fill":"34D399","name":"c3bar"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"pptx","x":23.1,"y":7.9,"w":8.0,"h":1.5,"fontSize":40,"bold":true,"color":"34D399","name":"c3num"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"Presentations","x":23.1,"y":9.9,"w":8.0,"h":0.9,"fontSize":17,"bold":true,"color":"F8FAFC","name":"c3t"}},
  {"op":"add","path":"/slide[2]","type":"shape","props":{"text":"Slides, shapes, native\ncharts, SmartArt, tables,\ntransitions and\nanimations.","x":23.1,"y":11.0,"w":8.0,"h":3.5,"fontSize":13,"color":"94A3B8","name":"c3b"}}
]'

# slide 3: a REAL native chart (data labels + legend + axis titles), framed in a card
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[3]","type":"shape","props":{"shape":"rect","x":2.0,"y":2.1,"w":3.6,"h":0.14,"fill":"38BDF8","name":"hdrBar"}},
  {"op":"add","path":"/slide[3]","type":"shape","props":{"text":"Documents generated per week","x":2.0,"y":2.5,"w":28.0,"h":1.6,"fontSize":34,"bold":true,"color":"F8FAFC","name":"hdr"}},
  {"op":"add","path":"/slide[3]","type":"shape","props":{"text":"A native chart with data labels, legend and axis titles — written straight into the .pptx, opens in real PowerPoint.","x":2.0,"y":4.2,"w":30.0,"h":1.0,"fontSize":15,"color":"94A3B8","name":"sub"}}
]'
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[3]","type":"chart","props":{
    "kind":"bar",
    "categories":["W1","W2","W3","W4","W5","W6"],
    "series":[
      {"name":"docx","values":[120,180,240,310,360,420]},
      {"name":"xlsx","values":[90,150,210,260,330,390]},
      {"name":"pptx","values":[60,110,170,230,300,370]}
    ],
    "title":"Weekly output by format",
    "x":2.0,"y":6.0,"w":29.8,"h":11.5,
    "dataLabels":{"show":"value","position":"outEnd"},
    "legend":"bottom",
    "axisTitles":{"category":"Week","value":"Documents"},
    "gridlines":{"major":true,"minor":false}
  }}
]'
# frame the native chart in a rounded white "card" panel, then send the panel behind it
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[{"op":"add","path":"/slide[3]","type":"shape","props":{"shape":"roundRect","x":1.7,"y":5.7,"w":30.4,"h":12.2,"fill":"F8FAFC","name":"chartPanel"}}]'
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"move","path":"/slide[3]/shape[@id=6]","position":"back"},
  {"op":"set","path":"/slide[3]/shape[@id=5]","props":{"x":2.1,"y":6.0,"w":29.6,"h":11.4}}
]'

# slide 4: render -> look -> fix (3 numbered step cards + arrows)
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[4]","type":"shape","props":{"shape":"rect","x":2.0,"y":2.1,"w":3.6,"h":0.14,"fill":"38BDF8","name":"hdrBar"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"render -> look -> fix","x":2.0,"y":2.5,"w":28.0,"h":1.6,"fontSize":34,"bold":true,"color":"F8FAFC","name":"hdr"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"AIOffice renders any node to PNG so an agent can SEE its own output, judge it, and edit until it is right.","x":2.0,"y":4.2,"w":30.0,"h":1.0,"fontSize":15,"color":"94A3B8","name":"sub"}}
]'
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[4]","type":"shape","props":{"shape":"roundRect","x":2.2,"y":7.6,"w":8.6,"h":7.2,"fill":"1E293B","name":"step1"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"shape":"roundRect","x":12.6,"y":7.6,"w":8.6,"h":7.2,"fill":"1E293B","name":"step2"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"shape":"roundRect","x":23.0,"y":7.6,"w":8.6,"h":7.2,"fill":"1E293B","name":"step3"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"shape":"arrow","x":10.9,"y":10.9,"w":1.6,"h":0.7,"fill":"38BDF8","name":"arr1"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"shape":"arrow","x":21.3,"y":10.9,"w":1.6,"h":0.7,"fill":"38BDF8","name":"arr2"}}
]'
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"01","x":2.9,"y":8.3,"w":7.0,"h":1.2,"fontSize":30,"bold":true,"color":"38BDF8","name":"s1n"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"render","x":2.9,"y":9.8,"w":7.0,"h":1.0,"fontSize":20,"bold":true,"color":"F8FAFC","name":"s1t"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"aioffice render deck.pptx\n--to png --scope /slide[3]\n\nThe deck becomes a PNG\nthe agent can read back.","x":2.9,"y":11.0,"w":7.2,"h":3.4,"fontSize":12,"color":"94A3B8","name":"s1b"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"02","x":13.3,"y":8.3,"w":7.0,"h":1.2,"fontSize":30,"bold":true,"color":"818CF8","name":"s2n"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"look","x":13.3,"y":9.8,"w":7.0,"h":1.0,"fontSize":20,"bold":true,"color":"F8FAFC","name":"s2t"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"The agent inspects the\nimage: spacing, contrast,\nalignment, overflow.\n\nIs this actually good?","x":13.3,"y":11.0,"w":7.2,"h":3.4,"fontSize":12,"color":"94A3B8","name":"s2b"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"03","x":23.7,"y":8.3,"w":7.0,"h":1.2,"fontSize":30,"bold":true,"color":"34D399","name":"s3n"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"fix","x":23.7,"y":9.8,"w":7.0,"h":1.0,"fontSize":20,"bold":true,"color":"F8FAFC","name":"s3t"}},
  {"op":"add","path":"/slide[4]","type":"shape","props":{"text":"aioffice edit ... nudges\nsizes, colors, positions.\n\nLoop until every slide\nlooks deliberate.","x":23.7,"y":11.0,"w":7.2,"h":3.4,"fontSize":12,"color":"94A3B8","name":"s3b"}}
]'

# slide 5: errors that teach (agent tries | aioffice answers)
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[5]","type":"shape","props":{"shape":"rect","x":2.0,"y":2.1,"w":3.6,"h":0.14,"fill":"38BDF8","name":"hdrBar"}},
  {"op":"add","path":"/slide[5]","type":"shape","props":{"text":"Errors that teach","x":2.0,"y":2.5,"w":28.0,"h":1.6,"fontSize":34,"bold":true,"color":"F8FAFC","name":"hdr"}},
  {"op":"add","path":"/slide[5]","type":"shape","props":{"text":"Every failure returns ONE JSON envelope with an actionable suggestion — so the agent can self-correct, not guess.","x":2.0,"y":4.2,"w":30.0,"h":1.0,"fontSize":15,"color":"94A3B8","name":"sub"}}
]'
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[5]","type":"shape","props":{"shape":"roundRect","x":2.0,"y":6.4,"w":13.0,"h":11.2,"fill":"1E293B","name":"leftCard"}},
  {"op":"add","path":"/slide[5]","type":"shape","props":{"text":"THE AGENT TRIES","x":2.8,"y":7.1,"w":11.5,"h":0.8,"fontSize":12,"bold":true,"color":"38BDF8","name":"lLabel"}},
  {"op":"add","path":"/slide[5]","type":"shape","props":{"text":"aioffice edit deck.pptx\n  --set /slide[9]/shape[7]\n  text=\"Hi\"","x":2.8,"y":8.1,"w":11.5,"h":2.6,"fontSize":15,"color":"E2E8F0","name":"lCmd"}},
  {"op":"add","path":"/slide[5]","type":"shape","props":{"text":"Slide 9 does not exist. A naive\ntool would crash or hallucinate.","x":2.8,"y":13.2,"w":11.5,"h":2.4,"fontSize":13,"color":"94A3B8","name":"lNote"}},
  {"op":"add","path":"/slide[5]","type":"shape","props":{"shape":"roundRect","x":16.5,"y":6.4,"w":15.4,"h":11.2,"fill":"0B1220","name":"rightCard"}},
  {"op":"add","path":"/slide[5]","type":"shape","props":{"text":"AIOFFICE ANSWERS","x":17.3,"y":7.1,"w":13.5,"h":0.8,"fontSize":12,"bold":true,"color":"34D399","name":"rLabel"}},
  {"op":"add","path":"/slide[5]","type":"shape","props":{"text":"{\n  \"ok\": false,\n  \"error\": {\n    \"code\": \"path_not_found\",\n    \"message\": \"No slide[9];\n       the deck has 6 slides.\",\n    \"suggestion\": \"Use /slide[1..6],\n       or add a slide first.\"\n  }\n}","x":17.3,"y":8.1,"w":14.0,"h":9.0,"fontSize":13,"color":"E2E8F0","name":"rEnv"}}
]'

# slide 6: closing (headline, npx install card, repo url, accent blocks)
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"add","path":"/slide[6]","type":"shape","props":{"shape":"rect","x":3.0,"y":4.6,"w":7.2,"h":0.16,"fill":"38BDF8","name":"accentBar"}},
  {"op":"add","path":"/slide[6]","type":"shape","props":{"shape":"rect","x":29.8,"y":1.6,"w":4.07,"h":3.0,"fill":"38BDF8","name":"cornerBlock"}},
  {"op":"add","path":"/slide[6]","type":"shape","props":{"shape":"rect","x":0.0,"y":15.2,"w":3.4,"h":3.85,"fill":"818CF8","name":"cornerBlock2"}},
  {"op":"add","path":"/slide[6]","type":"shape","props":{"text":"Give your agent an\nOffice engine.","x":2.85,"y":5.2,"w":24.0,"h":4.2,"fontSize":52,"bold":true,"color":"F8FAFC","name":"closeTitle"}},
  {"op":"add","path":"/slide[6]","type":"shape","props":{"shape":"roundRect","x":3.0,"y":11.4,"w":18.0,"h":2.0,"fill":"0B1220","name":"installCard"}},
  {"op":"add","path":"/slide[6]","type":"shape","props":{"text":"$  npx @aioffice/cli create deck.pptx","x":3.7,"y":11.9,"w":17.0,"h":1.1,"fontSize":18,"color":"38BDF8","name":"installCmd"}},
  {"op":"add","path":"/slide[6]","type":"shape","props":{"text":"github.com/onecer/aioffice","x":3.0,"y":14.2,"w":24.0,"h":1.0,"fontSize":18,"bold":true,"color":"E2E8F0","name":"repo"}},
  {"op":"add","path":"/slide[6]","type":"shape","props":{"text":"One binary · 18 verbs · 17 MCP tools · docx + xlsx + pptx","x":3.0,"y":15.3,"w":24.0,"h":0.9,"fontSize":14,"color":"64748B","name":"closeSub"}}
]'

# document metadata
run "$AIO" edit "$WS/deck.pptx" --workspace "$WS" --ops '[
  {"op":"set","path":"/properties","props":{
    "title":"AIOffice — An AI-native Office engine for agents",
    "author":"AIOffice","subject":"Product pitch deck",
    "keywords":"aioffice;cli;mcp;office;docx;xlsx;pptx","category":"Pitch Deck",
    "comments":"Built entirely by aioffice — no Office installed."}}
]'

# prove it is sound, then export the six slide PNGs (~1280 wide)
run "$AIO" validate "$WS/deck.pptx" --workspace "$WS"
render_pages "$WS/deck.pptx" "$WS/deck" 6
echo

# =============================================================================
# 2) dashboard.xlsx — an FY2025 regional revenue dashboard (formulas + charts)
# =============================================================================
echo "==> [2/3] dashboard.xlsx — FY2025 revenue dashboard (live formulas, charts)"

# create the workbook (the first sheet takes the --title name: "Dashboard")
run "$AIO" create "$WS/dashboard.xlsx" --kind xlsx --title "Dashboard" --workspace "$WS"

# title, subtitle, table headers, and the 8-region quarterly dataset (one bulk write)
run "$AIO" edit "$WS/dashboard.xlsx" --workspace "$WS" --ops '[
  {"op":"set","path":"/Dashboard/B2","props":{"value":"Northwind Cloud — FY2025 Revenue Dashboard"}},
  {"op":"set","path":"/Dashboard/B3","props":{"value":"Quarterly revenue, units, and margin by sales region · built entirely with aioffice"}},
  {"op":"set","path":"/Dashboard/B9","props":{"value":"Region"}},
  {"op":"set","path":"/Dashboard/C9","props":{"value":"Q1"}},
  {"op":"set","path":"/Dashboard/D9","props":{"value":"Q2"}},
  {"op":"set","path":"/Dashboard/E9","props":{"value":"Q3"}},
  {"op":"set","path":"/Dashboard/F9","props":{"value":"Q4"}},
  {"op":"set","path":"/Dashboard/G9","props":{"value":"FY Revenue"}},
  {"op":"set","path":"/Dashboard/H9","props":{"value":"Units"}},
  {"op":"set","path":"/Dashboard/I9","props":{"value":"Margin %"}},
  {"op":"set","path":"/Dashboard/B10:F17","props":{"values":[
    ["North America", 412000, 458000, 503000, 561000],
    ["EMEA",          318000, 342000, 366000, 401000],
    ["LATAM",         128000, 141000, 152000, 168000],
    ["APAC",          274000, 309000, 358000, 412000],
    ["UK & Ireland",  186000, 197000, 205000, 223000],
    ["DACH",          221000, 238000, 251000, 269000],
    ["Nordics",       97000,  104000, 112000, 124000],
    ["MEA",           74000,  82000,  91000,  103000]
  ]}},
  {"op":"set","path":"/Dashboard/H10","props":{"value":3120}},
  {"op":"set","path":"/Dashboard/H11","props":{"value":2410}},
  {"op":"set","path":"/Dashboard/H12","props":{"value":1180}},
  {"op":"set","path":"/Dashboard/H13","props":{"value":2640}},
  {"op":"set","path":"/Dashboard/H14","props":{"value":1390}},
  {"op":"set","path":"/Dashboard/H15","props":{"value":1720}},
  {"op":"set","path":"/Dashboard/H16","props":{"value":760}},
  {"op":"set","path":"/Dashboard/H17","props":{"value":540}},
  {"op":"set","path":"/Dashboard/I10","props":{"value":0.342}},
  {"op":"set","path":"/Dashboard/I11","props":{"value":0.311}},
  {"op":"set","path":"/Dashboard/I12","props":{"value":0.268}},
  {"op":"set","path":"/Dashboard/I13","props":{"value":0.357}},
  {"op":"set","path":"/Dashboard/I14","props":{"value":0.298}},
  {"op":"set","path":"/Dashboard/I15","props":{"value":0.324}},
  {"op":"set","path":"/Dashboard/I16","props":{"value":0.281}},
  {"op":"set","path":"/Dashboard/I17","props":{"value":0.252}}
]'

# live formulas: per-row FY revenue (=SUM), a totals row (=SUM / =AVERAGE) — aioffice
# evaluates and caches every result into the saved file
run "$AIO" edit "$WS/dashboard.xlsx" --workspace "$WS" --ops '[
  {"op":"set","path":"/Dashboard/G10","props":{"value":"=SUM(C10:F10)"}},
  {"op":"set","path":"/Dashboard/G11","props":{"value":"=SUM(C11:F11)"}},
  {"op":"set","path":"/Dashboard/G12","props":{"value":"=SUM(C12:F12)"}},
  {"op":"set","path":"/Dashboard/G13","props":{"value":"=SUM(C13:F13)"}},
  {"op":"set","path":"/Dashboard/G14","props":{"value":"=SUM(C14:F14)"}},
  {"op":"set","path":"/Dashboard/G15","props":{"value":"=SUM(C15:F15)"}},
  {"op":"set","path":"/Dashboard/G16","props":{"value":"=SUM(C16:F16)"}},
  {"op":"set","path":"/Dashboard/G17","props":{"value":"=SUM(C17:F17)"}},
  {"op":"set","path":"/Dashboard/B18","props":{"value":"Total"}},
  {"op":"set","path":"/Dashboard/C18","props":{"value":"=SUM(C10:C17)"}},
  {"op":"set","path":"/Dashboard/D18","props":{"value":"=SUM(D10:D17)"}},
  {"op":"set","path":"/Dashboard/E18","props":{"value":"=SUM(E10:E17)"}},
  {"op":"set","path":"/Dashboard/F18","props":{"value":"=SUM(F10:F17)"}},
  {"op":"set","path":"/Dashboard/G18","props":{"value":"=SUM(G10:G17)"}},
  {"op":"set","path":"/Dashboard/H18","props":{"value":"=SUM(H10:H17)"}},
  {"op":"set","path":"/Dashboard/I18","props":{"value":"=AVERAGE(I10:I17)"}}
]'

# KPI band (row 5 labels / row 6 values). Top Region and Best Quarter are real
# =XLOOKUP(MAX(...),...) results that aioffice evaluates at write time; Regions is =COUNTA
run "$AIO" edit "$WS/dashboard.xlsx" --workspace "$WS" --ops '[
  {"op":"set","path":"/Dashboard/B5","props":{"value":"Total FY Revenue"}},
  {"op":"set","path":"/Dashboard/C5","props":{"value":"Units Sold"}},
  {"op":"set","path":"/Dashboard/D5","props":{"value":"Avg Margin"}},
  {"op":"set","path":"/Dashboard/E5","props":{"value":"Top Region"}},
  {"op":"set","path":"/Dashboard/F5","props":{"value":"Best Quarter"}},
  {"op":"set","path":"/Dashboard/G5","props":{"value":"Regions"}},
  {"op":"set","path":"/Dashboard/B6","props":{"value":"=G18"}},
  {"op":"set","path":"/Dashboard/C6","props":{"value":"=H18"}},
  {"op":"set","path":"/Dashboard/D6","props":{"value":"=I18"}},
  {"op":"set","path":"/Dashboard/E6","props":{"value":"=XLOOKUP(MAX(G10:G17),G10:G17,B10:B17)"}},
  {"op":"set","path":"/Dashboard/F6","props":{"value":"=XLOOKUP(MAX(C18:F18),C18:F18,C9:F9)"}},
  {"op":"set","path":"/Dashboard/G6","props":{"value":"=COUNTA(B10:B17)"}}
]'

# number formats: compact USD currency, percent, thousands; KPI value formats
run "$AIO" edit "$WS/dashboard.xlsx" --workspace "$WS" --ops '[
  {"op":"set","path":"/Dashboard/C10:G18","props":{"numberFormat":"\"$\"#,##0"}},
  {"op":"set","path":"/Dashboard/H10:H18","props":{"numberFormat":"#,##0"}},
  {"op":"set","path":"/Dashboard/I10:I18","props":{"numberFormat":"0.0%"}},
  {"op":"set","path":"/Dashboard/B6","props":{"numberFormat":"\"$\"#,##0"}},
  {"op":"set","path":"/Dashboard/C6","props":{"numberFormat":"#,##0"}},
  {"op":"set","path":"/Dashboard/D6","props":{"numberFormat":"0.0%"}},
  {"op":"set","path":"/Dashboard/G6","props":{"numberFormat":"0"}}
]'

# light navy theme: header band, KPI cards, zebra-banded data rows, totals band,
# bold region names. (Real fills/colors in the .xlsx; visible when opened in Excel.)
run "$AIO" edit "$WS/dashboard.xlsx" --workspace "$WS" --ops '[
  {"op":"set","path":"/Dashboard/A1:J19","props":{"fill":"FFFFFF"}},
  {"op":"set","path":"/Dashboard/B2","props":{"bold":true,"color":"0F2742"}},
  {"op":"set","path":"/Dashboard/B3","props":{"color":"64748B"}},
  {"op":"set","path":"/Dashboard/B5:G5","props":{"fill":"E8F1FB","color":"35597E","bold":true}},
  {"op":"set","path":"/Dashboard/B6:G6","props":{"fill":"F4F9FF","color":"0F2742","bold":true}},
  {"op":"set","path":"/Dashboard/B9:I9","props":{"fill":"143C66","color":"FFFFFF","bold":true}},
  {"op":"set","path":"/Dashboard/B10:I10","props":{"fill":"FFFFFF","color":"1E293B"}},
  {"op":"set","path":"/Dashboard/B11:I11","props":{"fill":"F1F6FC","color":"1E293B"}},
  {"op":"set","path":"/Dashboard/B12:I12","props":{"fill":"FFFFFF","color":"1E293B"}},
  {"op":"set","path":"/Dashboard/B13:I13","props":{"fill":"F1F6FC","color":"1E293B"}},
  {"op":"set","path":"/Dashboard/B14:I14","props":{"fill":"FFFFFF","color":"1E293B"}},
  {"op":"set","path":"/Dashboard/B15:I15","props":{"fill":"F1F6FC","color":"1E293B"}},
  {"op":"set","path":"/Dashboard/B16:I16","props":{"fill":"FFFFFF","color":"1E293B"}},
  {"op":"set","path":"/Dashboard/B17:I17","props":{"fill":"F1F6FC","color":"1E293B"}},
  {"op":"set","path":"/Dashboard/B10:B17","props":{"bold":true,"color":"0F2742"}},
  {"op":"set","path":"/Dashboard/B18:I18","props":{"fill":"D6E6F7","color":"0F2742","bold":true}}
]'

# column widths and row heights for a tidy used range
run "$AIO" edit "$WS/dashboard.xlsx" --workspace "$WS" --ops '[
  {"op":"set","path":"/Dashboard/col[A]","props":{"width":2.5}},
  {"op":"set","path":"/Dashboard/col[B]","props":{"width":16}},
  {"op":"set","path":"/Dashboard/col[C]","props":{"width":14}},
  {"op":"set","path":"/Dashboard/col[D]","props":{"width":14}},
  {"op":"set","path":"/Dashboard/col[E]","props":{"width":14}},
  {"op":"set","path":"/Dashboard/col[F]","props":{"width":14}},
  {"op":"set","path":"/Dashboard/col[G]","props":{"width":16}},
  {"op":"set","path":"/Dashboard/col[H]","props":{"width":10}},
  {"op":"set","path":"/Dashboard/col[I]","props":{"width":11}},
  {"op":"set","path":"/Dashboard/row[2]","props":{"height":26}},
  {"op":"set","path":"/Dashboard/row[3]","props":{"height":18}},
  {"op":"set","path":"/Dashboard/row[5]","props":{"height":16}},
  {"op":"set","path":"/Dashboard/row[6]","props":{"height":24}},
  {"op":"set","path":"/Dashboard/row[9]","props":{"height":20}},
  {"op":"set","path":"/Dashboard/row[10]","props":{"height":18}},
  {"op":"set","path":"/Dashboard/row[11]","props":{"height":18}},
  {"op":"set","path":"/Dashboard/row[12]","props":{"height":18}},
  {"op":"set","path":"/Dashboard/row[13]","props":{"height":18}},
  {"op":"set","path":"/Dashboard/row[14]","props":{"height":18}},
  {"op":"set","path":"/Dashboard/row[15]","props":{"height":18}},
  {"op":"set","path":"/Dashboard/row[16]","props":{"height":18}},
  {"op":"set","path":"/Dashboard/row[17]","props":{"height":18}},
  {"op":"set","path":"/Dashboard/row[18]","props":{"height":20}}
]'

# conditional formatting: a data bar on Margin %, a color scale on FY Revenue
run "$AIO" edit "$WS/dashboard.xlsx" --workspace "$WS" --ops '[
  {"op":"add","path":"/Dashboard/I10:I17","type":"conditionalFormat","props":{"kind":"dataBar","color":"38BDF8"}},
  {"op":"add","path":"/Dashboard/G10:G17","type":"conditionalFormat","props":{"kind":"colorScale","minColor":"DCEBFB","maxColor":"143C66"}}
]'

# chart source: a compact Region + FY-Revenue helper block in K:L (the bar chart
# needs two adjacent columns). Values reference the table cells.
run "$AIO" edit "$WS/dashboard.xlsx" --workspace "$WS" --ops '[
  {"op":"set","path":"/Dashboard/K9","props":{"value":"Region"}},
  {"op":"set","path":"/Dashboard/L9","props":{"value":"FY Revenue"}},
  {"op":"set","path":"/Dashboard/K10","props":{"value":"=B10"}},
  {"op":"set","path":"/Dashboard/K11","props":{"value":"=B11"}},
  {"op":"set","path":"/Dashboard/K12","props":{"value":"=B12"}},
  {"op":"set","path":"/Dashboard/K13","props":{"value":"=B13"}},
  {"op":"set","path":"/Dashboard/K14","props":{"value":"=B14"}},
  {"op":"set","path":"/Dashboard/K15","props":{"value":"=B15"}},
  {"op":"set","path":"/Dashboard/K16","props":{"value":"=B16"}},
  {"op":"set","path":"/Dashboard/K17","props":{"value":"=B17"}},
  {"op":"set","path":"/Dashboard/L10","props":{"value":"=G10"}},
  {"op":"set","path":"/Dashboard/L11","props":{"value":"=G11"}},
  {"op":"set","path":"/Dashboard/L12","props":{"value":"=G12"}},
  {"op":"set","path":"/Dashboard/L13","props":{"value":"=G13"}},
  {"op":"set","path":"/Dashboard/L14","props":{"value":"=G14"}},
  {"op":"set","path":"/Dashboard/L15","props":{"value":"=G15"}},
  {"op":"set","path":"/Dashboard/L16","props":{"value":"=G16"}},
  {"op":"set","path":"/Dashboard/L17","props":{"value":"=G17"}}
]'

# two native, polished charts: a bar chart of FY revenue by region (data labels +
# axis titles), and a multi-series line chart of the quarterly trend
run "$AIO" edit "$WS/dashboard.xlsx" --workspace "$WS" --ops '[
  {"op":"add","path":"/Dashboard","type":"chart","props":{
    "kind":"bar","dataRange":"K9:L17","anchor":"B20","title":"FY Revenue by Region",
    "legend":"none","dataLabels":{"show":"value"},
    "axisTitles":{"category":"Region","value":"FY Revenue (USD)"}}},
  {"op":"add","path":"/Dashboard","type":"chart","props":{
    "kind":"line","dataRange":"B9:F17","anchor":"B40","title":"Quarterly Revenue Trend by Region",
    "legend":"bottom","axisTitles":{"category":"Quarter","value":"Revenue (USD)"},
    "gridlines":{"major":true,"minor":false}}}
]'

# workbook metadata (gives the file a document Title, clears the audit check)
run "$AIO" edit "$WS/dashboard.xlsx" --workspace "$WS" --ops '[
  {"op":"set","path":"/properties","props":{"title":"Northwind Cloud FY2025 Revenue Dashboard","author":"AIOffice"}}
]'

# prove the file is sound, then render page 1 (KPI band + table + bar chart).
# LibreOffice paginates the sheet; the aioffice fallback scopes to B5:I18.
run "$AIO" validate "$WS/dashboard.xlsx" --workspace "$WS"
run "$AIO" audit    "$WS/dashboard.xlsx" --workspace "$WS"
render_page1 "$WS/dashboard.xlsx" "$WS/dashboard.png" "/Dashboard/B5:I18"
echo

# =============================================================================
# 3) report.docx — a capability report (styles, table formula, equation, citations)
# =============================================================================
echo "==> [3/3] report.docx — capability report (table formula, LaTeX math, citations)"

# create the document (Heading1 title seeded from --title)
run "$AIO" create "$WS/report.docx" --kind docx --title "AIOffice Capability Report" --workspace "$WS"

# custom accent theme, package metadata, and US-Letter page setup
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"set","path":"/theme","props":{"accent1":"2563EB","accent2":"0EA5E9","dk1":"0F172A","dk2":"1E293B","lt2":"F1F5F9","majorFont":"Calibri Light","minorFont":"Calibri"}},
  {"op":"set","path":"/properties","props":{"title":"AIOffice Capability Report","subject":"What a single self-built binary can author","author":"AIOffice","keywords":"AIOffice, OOXML, AI-native, CLI, MCP","category":"Product","custom":{"Version":"1.7.0","Engine":"C#/.NET single binary","SurfaceVersion":1.0}}},
  {"op":"set","path":"/section[1]","props":{"pageSize":"Letter","marginTop":"2.2cm","marginBottom":"2cm","marginLeft":"2.4cm","marginRight":"2.4cm"}}
]'

# custom paragraph styles for the masthead and lead text
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/styles","type":"style","props":{"id":"Subtitle","kind":"paragraph","name":"Subtitle","basedOn":"Normal","color":"475569","fontSize":13,"spacingBefore":2,"spacingAfter":14}},
  {"op":"add","path":"/styles","type":"style","props":{"id":"Kicker","kind":"paragraph","name":"Kicker","basedOn":"Normal","color":"2563EB","fontSize":10,"bold":true,"spacingAfter":2}},
  {"op":"add","path":"/styles","type":"style","props":{"id":"Lead","kind":"paragraph","name":"Lead","basedOn":"Normal","color":"334155","fontSize":11,"spacingAfter":8}}
]'

# masthead: kicker eyebrow before the title, subtitle after it
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/body/p[1]","type":"p","position":"before","props":{"text":"AIOFFICE  ·  CAPABILITY REPORT","style":"Kicker"}},
  {"op":"set","path":"/body/p[3]","props":{"text":"One self-built C#/.NET binary that authors real Office files","style":"Subtitle"}}
]'

# intro paragraph + the first H2 section and its table intro
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"AIOffice is a single self-built binary and MCP server that reads, edits, and renders genuine .docx, .xlsx, and .pptx files. Every element below — the heading hierarchy, the computed table total, the typeset equation, and the formatted bibliography — was produced by aioffice commands alone. This page is the artifact and the proof.","style":"Lead"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Coverage by surface","style":"Heading2"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"The Total row is a live Word table formula, =SUM(ABOVE), that aioffice computed headlessly when the file was written.","style":"Lead"}}
]'

# a 5x3 data table
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/body","type":"table","position":"inside","props":{"rows":5,"cols":3}}
]'

# fill the cells; the Total row carries a real =SUM(ABOVE) table formula that
# aioffice computes now and stores as a Word field
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"text":"Surface"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"text":"Editable properties"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[3]","props":{"text":"Live operations"}},
  {"op":"set","path":"/body/table[1]/tr[2]/tc[1]","props":{"text":"Word (.docx)"}},
  {"op":"set","path":"/body/table[1]/tr[2]/tc[2]","props":{"text":"58"}},
  {"op":"set","path":"/body/table[1]/tr[2]/tc[3]","props":{"text":"31"}},
  {"op":"set","path":"/body/table[1]/tr[3]/tc[1]","props":{"text":"Excel (.xlsx)"}},
  {"op":"set","path":"/body/table[1]/tr[3]/tc[2]","props":{"text":"46"}},
  {"op":"set","path":"/body/table[1]/tr[3]/tc[3]","props":{"text":"24"}},
  {"op":"set","path":"/body/table[1]/tr[4]/tc[1]","props":{"text":"PowerPoint (.pptx)"}},
  {"op":"set","path":"/body/table[1]/tr[4]/tc[2]","props":{"text":"39"}},
  {"op":"set","path":"/body/table[1]/tr[4]/tc[3]","props":{"text":"19"}},
  {"op":"set","path":"/body/table[1]/tr[5]/tc[1]","props":{"text":"Total"}},
  {"op":"set","path":"/body/table[1]/tr[5]/tc[2]","props":{"formula":"=SUM(ABOVE)","numberFormat":"integer"}},
  {"op":"set","path":"/body/table[1]/tr[5]/tc[3]","props":{"formula":"=SUM(ABOVE)","numberFormat":"integer"}}
]'

# style the table: header row + total band, accent shading, clean borders, widths
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"set","path":"/body/table[1]","props":{"borders":"all","borderColor":"E2E8F0","borderWidthPt":0.75,"headerRow":true,"width":"100%","columnWidths":["7.5cm","4.5cm","4.5cm"],"cellPaddingCm":0.18}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[1]","props":{"shading":"0F172A"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[2]","props":{"shading":"0F172A"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[3]","props":{"shading":"0F172A"}},
  {"op":"set","path":"/body/table[1]/tr[5]/tc[1]","props":{"shading":"EFF6FF"}},
  {"op":"set","path":"/body/table[1]/tr[5]/tc[2]","props":{"shading":"EFF6FF"}},
  {"op":"set","path":"/body/table[1]/tr[5]/tc[3]","props":{"shading":"EFF6FF"}}
]'
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"set","path":"/body/table[1]/tr[1]/tc[1]/p[1]/run[1]","props":{"color":"FFFFFF"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[2]/p[1]/run[1]","props":{"color":"FFFFFF"}},
  {"op":"set","path":"/body/table[1]/tr[1]/tc[3]/p[1]/run[1]","props":{"color":"FFFFFF"}},
  {"op":"set","path":"/body/table[1]/tr[5]/tc[1]/p[1]/run[1]","props":{"bold":true,"color":"0F172A"}}
]'

# equation section + a NUMBERED display equation: the quadratic formula.
# LaTeX in -> native Office Math (m:oMathPara) out, no LaTeX install.
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Typeset mathematics","style":"Heading2"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"LaTeX in, native Office Math out. The expression below is real m:oMathPara markup from the built-in converter — no LaTeX install, no image.","style":"Lead"}}
]'
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/body","type":"equation","props":{"latex":"x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}","display":true,"number":true}}
]'

# bullet-list section
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Built entirely by the CLI","style":"Heading2"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Each item below maps to a documented capability:","style":"Lead"}},
  {"op":"add","path":"/body","type":"p","props":{"text":"Heading hierarchy, a running header, and a Page X of Y footer field","list":"bullet"}},
  {"op":"add","path":"/body","type":"p","props":{"text":"A formatted table whose total is a live =SUM(ABOVE) field computed headlessly","list":"bullet"}},
  {"op":"add","path":"/body","type":"p","props":{"text":"A numbered display equation typeset from LaTeX into native Office Math","list":"bullet"}},
  {"op":"add","path":"/body","type":"p","props":{"text":"A managed citation and an APA bibliography drawn from the document source store","list":"bullet"}},
  {"op":"add","path":"/body","type":"p","props":{"text":"A custom accent theme, document metadata, and validation that reports zero issues","list":"bullet"}}
]'

# sources go into the document store; the References section cites both
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/sources","type":"source","props":{"tag":"ECMA376","kind":"report","author":"ECMA International","title":"Office Open XML File Formats (ECMA-376)","year":2016}},
  {"op":"add","path":"/sources","type":"source","props":{"tag":"MCP2025","kind":"website","author":"Anthropic","title":"Model Context Protocol Specification","year":2025}}
]'
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"Standards and protocol","style":"Heading2"}},
  {"op":"add","path":"/body","type":"p","position":"inside","props":{"text":"AIOffice writes the published Office Open XML formats and speaks the Model Context Protocol.","style":"Lead"}}
]'

# drop the two CITATION fields into the section paragraph (/body/p[18])
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/body/p[18]","type":"citation","props":{"source":"ECMA376"}},
  {"op":"add","path":"/body/p[18]","type":"citation","props":{"source":"MCP2025"}}
]'

# render the APA bibliography of the cited sources
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/body","type":"bibliography","props":{"style":"APA"}}
]'

# running header + a Page X of Y footer built from PAGE/NUMPAGES fields
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/header[1]","type":"header","props":{"text":"AIOffice Capability Report"}},
  {"op":"add","path":"/footer[1]","type":"footer","props":{"text":"Generated entirely by aioffice 1.7.0     "}}
]'
run "$AIO" edit "$WS/report.docx" --workspace "$WS" --ops '[
  {"op":"add","path":"/footer[1]/p[1]","type":"field","props":{"kind":"pageNumber","leadingText":"Page "}},
  {"op":"add","path":"/footer[1]/p[1]","type":"field","props":{"kind":"numPages","leadingText":" of "}}
]'

# prove the file is sound (0 issues) and render the PNG used in the showcase
run "$AIO" validate "$WS/report.docx" --workspace "$WS"
render_page1 "$WS/report.docx" "$WS/report.png"
echo

# =============================================================================
# done
# =============================================================================
echo "==> Showcase rebuilt in: $WS"
echo "    Files:"
echo "      deck.pptx        + deck-1.png … deck-6.png   (6 dark slides, native chart)"
echo "      dashboard.xlsx   + dashboard.png             (formulas, charts, conditional formats)"
echo "      report.docx      + report.png                (table formula, LaTeX math, citations)"
echo
echo "    Every file above was authored only by aioffice — no Microsoft Office, no template."
if have_lo; then
  echo "    PNGs rendered with LibreOffice (real-Office fidelity)."
else
  echo "    PNGs rendered with aioffice's own renderer (install LibreOffice + poppler for full fidelity)."
fi
echo "    Open the PNGs to look; see ../SHOWCASE.md for the gallery and ../docs/COOKBOOK.md for recipes."
