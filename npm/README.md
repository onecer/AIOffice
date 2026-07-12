# aioffice

**AIOffice** is an AI-native command-line tool and [MCP](https://modelcontextprotocol.io)
server for working with real Office documents — `.docx`, `.xlsx`, and `.pptx`.
It creates, reads, queries, edits, renders, validates, and converts Office files
through a single stable JSON surface designed for AI agents and scripts.

This npm package is a thin installer: on install it downloads the correct
self-contained native binary for your platform from the official
[GitHub release](https://github.com/onecer/AIOffice/releases) and verifies it
against the release's `SHA256SUMS`. There are no other dependencies.

## Install

Global install (adds an `aioffice` command to your PATH):

```sh
npm install -g aioffice
aioffice version
aioffice doctor
```

Or run on demand without installing, via `npx`:

```sh
npx aioffice doctor
npx aioffice create report.docx --title "Q3 Report"
```

> On the first run, `npx` will download and SHA256-verify the native binary if
> the postinstall step did not already do so.

## Use as an MCP server

AIOffice speaks MCP over stdio. Point your MCP client at the `aioffice mcp`
command:

```json
{
  "mcpServers": {
    "aioffice": {
      "command": "aioffice",
      "args": ["mcp"]
    }
  }
}
```

If you installed globally, `aioffice` resolves from your PATH. If you prefer not
to install globally, use `npx` as the command instead:

```json
{
  "mcpServers": {
    "aioffice": {
      "command": "npx",
      "args": ["-y", "aioffice", "mcp"]
    }
  }
}
```

The MCP transport is plain stdio JSON-RPC; the npm shim passes stdin/stdout
through transparently.

## How installation works

1. `install.js` (run as a postinstall script) detects your platform and CPU
   architecture and maps them to the matching release asset:

   | OS / arch        | Asset                      |
   | ---------------- | -------------------------- |
   | macOS arm64      | `aioffice-mac-arm64`       |
   | macOS x64        | `aioffice-mac-x64`         |
   | Linux x64        | `aioffice-linux-x64`       |
   | Linux arm64      | `aioffice-linux-arm64`     |
   | Windows x64      | `aioffice-win-x64.exe`     |
   | Windows arm64    | `aioffice-win-arm64.exe`   |

2. It downloads `SHA256SUMS` and the binary from the release
   `v{package version}`, computes the SHA256 of the download, and **refuses to
   install** if it does not match the published checksum.
3. On success the binary is placed in this package's `bin/` directory and made
   executable (`chmod +x` on Unix). The install is idempotent — re-running it
   skips work when a verified binary is already present.

If the download or verification fails, installation aborts with a clear message
that includes the direct download URL so you can install the binary manually.

## Environment overrides

Both are useful for testing or for serving binaries from a private mirror:

| Variable                     | Default                                                           | Purpose                                  |
| ---------------------------- | ---------------------------------------------------------------- | ---------------------------------------- |
| `AIOFFICE_DOWNLOAD_VERSION`  | `v{package version}` (e.g. `v1.26.0`)                             | Release tag to download.                 |
| `AIOFFICE_DOWNLOAD_BASEURL`  | `https://github.com/onecer/AIOffice/releases/download`           | Base URL for the assets + `SHA256SUMS`.  |

The binary is fetched from `{BASEURL}/{VERSION}/{asset}` and the checksum file
from `{BASEURL}/{VERSION}/SHA256SUMS`.

```sh
# Example: install a specific version from a local mirror
AIOFFICE_DOWNLOAD_VERSION=v1.26.0 \
AIOFFICE_DOWNLOAD_BASEURL=https://mirror.example.com/aioffice \
npm install -g aioffice
```

## License

MIT. See the [AIOffice repository](https://github.com/onecer/AIOffice)
for source, full documentation, and the AI-facing surface contract.
