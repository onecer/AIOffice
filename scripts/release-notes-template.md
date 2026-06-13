<!--
  Rendered by .github/workflows/release.yml on every v* tag:
  {{version}} -> 0.5.0, {{tag}} -> v0.5.0. Edit the highlights per release
  BEFORE tagging — this file is the single source of the GitHub release notes.
-->
**AIOffice** is a 100% self-built, AI-native CLI + MCP server for real Office files — create, query, edit, render, preview and validate `.docx` / `.xlsx` / `.pptx` with one command in, exactly one JSON envelope out. Pure C#/.NET on DocumentFormat.OpenXml + ClosedXML — **no Microsoft Office, no runtime dependencies, no wrapped third-party engines.** One self-contained single-file binary per platform.

## Highlights in {{tag}}

The first release since `v0.4.0` — four milestones (M4–M7) of new capability, **1366 passing tests**, **15 MCP tools**.

- **`audit` — a new top-level verb (15th capability).** Accessibility + quality lint across all three formats (`aioffice audit file [--category accessibility|quality] [--fix]`): missing alt-text, heading-level skips, tables/slides without headers, low contrast (WCAG luminance), broken refs (`#REF!`/`#DIV/0!`), broken links, off-canvas shapes, empty placeholders — with `--fix` for safe autofixes. No other office CLI does this.
- **LaTeX → equations.** A self-built LaTeX→OMML converter (no LaTeX dependency): fractions, roots, sub/superscript, sums/integrals with limits, matrices, Greek, operators. `add type:equation` (inline + display); original LaTeX is archived for faithful read-back; unknown commands degrade gracefully.
- **Markdown / CSV bridges.** `create report.docx --from notes.md` (headings, lists, GFM tables, links, code) and `read --view markdown`; `create data.xlsx --from data.csv` (RFC-4180) and `read --view csv`.
- **Document-wide find/replace**, **TOC / watermark / endnotes / multi-section** docx, **editable-data pptx charts** + entrance/exit/emphasis animations, **pivot tables / data validation / threaded comments / Excel Tables**, **in-place streaming writes** for large workbooks, **pptx master/layout editing & sections**, **document properties** (core + custom), **content controls**, **PDF export**, and **cross-document data flow** (`dataFrom: "book.xlsx!Sheet1/A1:B5"`).

## Install

Download the binary for your platform below, then make it executable and put it on your `PATH`:

```bash
# macOS (Apple Silicon)
curl -L -o aioffice https://github.com/onecer/AIOffice/releases/download/{{tag}}/aioffice-mac-arm64
chmod +x aioffice && ./aioffice doctor
```

Verify the download against `SHA256SUMS`:

```bash
shasum -a 256 -c SHA256SUMS --ignore-missing
```

> **macOS Gatekeeper:** these binaries are unsigned for now. If macOS blocks the first run, clear the quarantine attribute: `xattr -d com.apple.quarantine ./aioffice` (or right-click → Open once). Signing & notarization are planned.

## MCP quickstart

```json
{
  "mcpServers": {
    "aioffice": {
      "command": "/path/to/aioffice",
      "args": ["mcp", "--workspace", "/path/to/your/documents"]
    }
  }
}
```

## Assets

| File | Platform |
|---|---|
| `aioffice-mac-arm64` | macOS Apple Silicon |
| `aioffice-mac-x64` | macOS Intel |
| `aioffice-linux-x64` | Linux x64 |
| `aioffice-linux-arm64` | Linux ARM64 |
| `aioffice-win-x64.exe` | Windows x64 |
| `aioffice-win-arm64.exe` | Windows ARM64 |
| `SHA256SUMS` | Checksums for all binaries |

See the [README](https://github.com/onecer/AIOffice/blob/main/README.md) · [中文 README](https://github.com/onecer/AIOffice/blob/main/README.zh-CN.md) for the full command surface and capability ledger.

License: Apache-2.0.
