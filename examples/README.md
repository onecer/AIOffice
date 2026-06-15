# Examples

Runnable examples for [AIOffice](../README.md) — the AI-native Office CLI + MCP
server.

## `tour.sh` — regenerate the whole showcase

[`tour.sh`](tour.sh) rebuilds the three artifacts in [SHOWCASE.md](../SHOWCASE.md)
from scratch — a dark pitch deck (`.pptx`), a regional revenue dashboard
(`.xlsx`), and a capability report (`.docx`) — and renders the gallery PNGs.
Every line in the script is a real `aioffice` command; there is no Microsoft
Office, no template, and no manual touch-up. The script is the proof: read it
top to bottom and you have seen exactly what produced the gallery.

```bash
# from the repo root, with aioffice on your PATH (see ../docs/INSTALL.md)
./examples/tour.sh                  # builds into a fresh temp dir, renders PNGs
OUTDIR=./out ./examples/tour.sh     # …or into a directory you keep
AIO=/path/to/aioffice ./examples/tour.sh   # …or point at a specific binary
```

**Requirements**

- `aioffice` on your PATH — or set `AIO=/path/to/aioffice`.
- Chrome / Chromium installed — only needed for the `render --to png` steps;
  the build (`create` / `edit` / `validate`) works without it.

**What it produces** (in the build directory it prints at the end):

| File | What's in it |
|---|---|
| `deck.pptx` + `deck-1.png`…`deck-6.png` | 6 dark 16:9 slides, format cards, a native bar chart, the teaching-error envelope |
| `dashboard.xlsx` + `dashboard.png` | a KPI band, an 8-region table with live `=SUM` totals and `=XLOOKUP` KPIs, two native charts, conditional formatting |
| `report.docx` + `report.png` | a styled masthead, a table with a computed `=SUM(ABOVE)` total, a numbered LaTeX→Office-Math equation, a managed citation pair + APA bibliography |

The script ends with `aioffice validate` (0 issues) on each file before
rendering, so a clean run is also a clean-OOXML proof.

---

## Agent quickstart

AIOffice is built for agents: one command in, exactly one JSON envelope out.
Drive it from a shell, or over **MCP** (the same surface, 17 tools that mirror
the CLI verbs 1:1).

**1 · Register the MCP server.** Run `aioffice mcp --workspace <dir>` and point
your host at it. The `--workspace` directory is the sandbox root — the agent can
only touch files inside it. Example for a generic stdio host / Claude Desktop:

```json
{
  "mcpServers": {
    "aioffice": {
      "command": "aioffice",
      "args": ["mcp", "--workspace", "/path/to/your/documents"]
    }
  }
}
```

Per-host snippets (Claude Desktop, Claude Code, Cursor, TonoBraid) are in
[../docs/MCP-SETUP.md](../docs/MCP-SETUP.md).

**2 · Point the agent at the skill.** Put [../SKILL.md](../SKILL.md) in the
agent's system prompt — it's the AI-facing onboarding guide: the envelope shape,
the addressing grammar (`/body/p[3]`, `/Sheet1/A1:C10`, `/slide[2]/shape[3]`),
and the read-before-write loop. Copy-paste task recipes live in
[../docs/COOKBOOK.md](../docs/COOKBOOK.md).

**3 · Run a first task.** A minimal end-to-end loop the agent can run today:

```bash
aioffice create report.docx --title "Q3 Report"
aioffice edit   report.docx --ops '[
  {"op":"add","path":"/body","type":"p","position":"inside",
   "props":{"text":"Revenue grew 12% quarter-over-quarter.","style":"Heading2"}}]'
aioffice read   report.docx --view outline     # orient: headings + canonical paths
aioffice validate report.docx                  # prove the OOXML is clean
aioffice render report.docx --to png -o report.png   # look at what you made
```

Then **look** at `report.png` and **fix** with another `edit` — the loop the
whole tool is designed around. When you're ready for the full picture, run
`./tour.sh` and read it alongside [SHOWCASE.md](../SHOWCASE.md).
