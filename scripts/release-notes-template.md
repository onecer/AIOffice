<!--
  Rendered by .github/workflows/release.yml on every v* tag:
  {{version}} -> 0.5.0, {{tag}} -> v0.5.0. Edit the highlights per release
  BEFORE tagging — this file is the single source of the GitHub release notes.
-->
**AIOffice** is a 100% self-built, AI-native CLI + MCP server for real Office files — create, query, edit, render, preview and validate `.docx` / `.xlsx` / `.pptx` with one command in, exactly one JSON envelope out. Pure C#/.NET on DocumentFormat.OpenXml + ClosedXML — **no Microsoft Office, no runtime dependencies, no wrapped third-party engines.** One self-contained single-file binary per platform.

## Highlights in {{tag}}

- **Document-wide find/replace** — `aioffice edit f --find X --replace Y [--regex] [--match-case] [--whole-word]` (and the `office_edit` replace op, scope `"/"`): docx body + headers + footers, every sheet, every slide including notes; matches split across runs are found; on docx `--track` records every replacement as a w:del+w:ins revision pair.
- **docx structure** — table of contents (`add toc`, read back via `/toc[1]`), text watermarks, endnotes, multi-section documents (`sectionBreak`, per-section page setup).
- **xlsx at scale** — bulk anchor writes (values + formulas in one op), row/column insert/delete with formula shifting, column width/hidden, legacy cell notes.
- **pptx motion & data** — entrance/exit/emphasis animations, slide comments, and charts whose embedded workbook makes them **Edit-Data-able in PowerPoint** (plus retrofitting old charts via `embedData:true`).

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
