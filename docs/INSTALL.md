# Installing AIOffice · 安装 AIOffice

> **中文摘要** — AIOffice 是一个单文件原生二进制（`aioffice`），无需安装 Office、无运行时依赖。四种安装方式：(1) **npm**：`npm i -g aioffice` 或 `npx aioffice`；(2) **Homebrew**：`brew install onecer/tap/aioffice`；(3) **一行脚本**（mac/Linux 用 `curl … | sh`，Windows 用 `irm … | iex`）；(4) **直接下载** Release 二进制并校验 SHA256。每个安装方式都会下载与你平台匹配的二进制，并用发行版的 `SHA256SUMS` 做完整性校验。macOS 首次运行被 Gatekeeper 拦截时，执行 `xattr -d com.apple.quarantine`（见文末）。

AIOffice ships as a **single self-contained native binary** — no .NET runtime, no Microsoft Office, no extra files. Pick whichever method fits your setup. Every method downloads the binary that matches your OS/CPU and verifies it against the release's `SHA256SUMS`.

| Platform | npm | Homebrew | One-line script | Direct download |
|---|---|---|---|---|
| macOS (arm64 / x64) | ✅ | ✅ | ✅ | ✅ |
| Linux (x64 / arm64) | ✅ | ✅ | ✅ | ✅ |
| Windows (x64 / arm64) | ✅ | — | ✅ (PowerShell) | ✅ |

The current release is **v1.6.0**. All commands below assume that tag; the npm/Homebrew/script flows track the latest release automatically.

---

## 1. npm (Node ≥ 18) · npm 安装

The npm package is a tiny wrapper: on install it downloads the matching native binary from the GitHub release and SHA256-verifies it. The binary itself is **not** shipped inside the tarball.

**Global install** (puts `aioffice` on your `PATH`):

```sh
npm i -g aioffice
aioffice version
```

**One-shot, no install** (great for CI and for MCP host configs):

```sh
npx aioffice version
npx aioffice mcp --workspace ./docs
```

**Pin a version / use a mirror** (env vars read by the postinstall step):

```sh
# install a specific release tag
AIOFFICE_DOWNLOAD_VERSION=v1.6.0 npm i -g aioffice

# download assets from a private mirror instead of github.com
AIOFFICE_DOWNLOAD_BASEURL=https://mirror.example.com/aioffice npm i -g aioffice
```

Notes:
- If you install with `--ignore-scripts` (postinstall skipped), the binary is fetched **lazily on first run** of `aioffice`.
- Uninstall with `npm rm -g aioffice`.
- The download is always SHA256-checked against the release `SHA256SUMS`; a mismatch aborts the install and nothing is written.

---

## 2. Homebrew (macOS / Linux) · Homebrew 安装

```sh
brew install onecer/tap/aioffice
aioffice version
```

or, in two steps:

```sh
brew tap onecer/tap
brew install aioffice
```

Upgrade / uninstall:

```sh
brew upgrade aioffice
brew uninstall aioffice
```

Homebrew installs the prebuilt single-file binary for your platform (no source build) and runs `aioffice version` as its formula test. The formula lives at [`dist/Formula/aioffice.rb`](../dist/Formula/aioffice.rb); the human publishes it into the `onecer/homebrew-tap` repo as `Formula/aioffice.rb`.

---

## 3. One-line install script · 一行脚本安装

### macOS / Linux

```sh
curl -fsSL https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.sh | sh
```

Installs into `~/.local/bin` by default. The script detects your OS/CPU, downloads the matching asset, **verifies its SHA256**, marks it executable, and (on macOS) strips the Gatekeeper quarantine attribute for you.

Customize with env vars:

```sh
# install a specific version
VERSION=v1.6.0 curl -fsSL https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.sh | sh

# install to a different directory (e.g. a dir already on PATH)
AIOFFICE_BIN=/usr/local/bin curl -fsSL https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.sh | sh
```

If the install dir is not on your `PATH`, the script prints the exact `export PATH=…` line to add to your shell profile.

> Prefer to read before you pipe to a shell? Download the script, inspect it, then run it:
> ```sh
> curl -fsSL https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.sh -o install.sh
> less install.sh
> sh install.sh
> ```

### Windows (PowerShell)

```powershell
irm https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.ps1 | iex
```

Installs `aioffice.exe` into `%LOCALAPPDATA%\aioffice` and adds that directory to your **user** `PATH` (effective in new terminals). Pin a version or change the directory first:

```powershell
$env:VERSION = "v1.6.0"; $env:AIOFFICE_BIN = "C:\tools\aioffice"
irm https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.ps1 | iex
```

The script verifies the download's SHA256 against the release `SHA256SUMS` before installing.

---

## 4. Direct binary download · 直接下载二进制

Grab the asset for your platform from the [releases page](https://github.com/onecer/AIOffice/releases), verify its checksum, make it executable, and put it on your `PATH`.

### Pick your asset

| OS | CPU | Asset |
|---|---|---|
| macOS | Apple Silicon (arm64) | `aioffice-mac-arm64` |
| macOS | Intel (x64) | `aioffice-mac-x64` |
| Linux | x64 | `aioffice-linux-x64` |
| Linux | arm64 | `aioffice-linux-arm64` |
| Windows | x64 | `aioffice-win-x64.exe` |
| Windows | arm64 | `aioffice-win-arm64.exe` |

Each release also publishes `SHA256SUMS` (one line per asset).

### macOS / Linux

```sh
# 1. download the binary + the checksums file (example: macOS arm64, v1.6.0)
VERSION=v1.6.0
ASSET=aioffice-mac-arm64
BASE="https://github.com/onecer/AIOffice/releases/download/$VERSION"
curl -fsSL "$BASE/$ASSET"      -o aioffice
curl -fsSL "$BASE/SHA256SUMS"  -o SHA256SUMS

# 2. verify (must print "<asset>: OK")
grep " $ASSET\$" SHA256SUMS | sed "s/ $ASSET\$/  aioffice/" | shasum -a 256 -c -

# 3. install
chmod +x aioffice
mkdir -p ~/.local/bin && mv aioffice ~/.local/bin/aioffice

# 4. confirm
aioffice version
```

If `~/.local/bin` is not on your `PATH`, add `export PATH="$HOME/.local/bin:$PATH"` to your `~/.zshrc` / `~/.bashrc`.

### Windows (PowerShell)

```powershell
# 1. download (example: x64, v1.6.0)
$Version = "v1.6.0"; $Asset = "aioffice-win-x64.exe"
$Base = "https://github.com/onecer/AIOffice/releases/download/$Version"
Invoke-WebRequest "$Base/$Asset"     -OutFile aioffice.exe
Invoke-WebRequest "$Base/SHA256SUMS" -OutFile SHA256SUMS

# 2. verify — compare these two values; they must match
(Get-FileHash aioffice.exe -Algorithm SHA256).Hash.ToLower()
(Select-String -Path SHA256SUMS -Pattern $Asset).Line.Split()[0].ToLower()

# 3. install (move aioffice.exe somewhere on your PATH) and confirm
.\aioffice.exe version
```

### macOS Gatekeeper note · macOS Gatekeeper 提示

The binaries are **not yet code-signed or notarized** (see [SIGNING.md](SIGNING.md)). On macOS, a directly downloaded binary may be quarantined and blocked on first run with *"cannot be opened because the developer cannot be verified."* Strip the quarantine attribute:

```sh
xattr -d com.apple.quarantine ~/.local/bin/aioffice
```

(The `install.sh` script and Homebrew do this for you; only direct downloads need it.) Alternatively, allow it once via **System Settings → Privacy & Security → Open Anyway**.

---

## Verify your download · 验证下载

Always confirm integrity against the release `SHA256SUMS` before running an unsigned binary.

```sh
# from a directory containing the downloaded asset(s) and SHA256SUMS:
shasum -a 256 -c SHA256SUMS 2>/dev/null            # macOS / older Linux
sha256sum -c SHA256SUMS 2>/dev/null                # Linux
```

A line ending in `OK` means the file matches; `FAILED` means do **not** run it — re-download.

---

## Quick start after install · 安装后快速上手

```sh
# create a doc, inspect it, validate it (all inside a sandbox workspace)
aioffice --workspace ./work create report.docx --title "Q3 Report"
aioffice --workspace ./work read report.docx --view outline
aioffice --workspace ./work validate report.docx

# run as an MCP server for an AI agent (see docs/MCP-SETUP.md)
aioffice mcp --workspace ./work
```

Every command emits a JSON envelope (`{ok,data,error,meta}`) and resolves all file paths inside the `--workspace` sandbox. See the [README](../README.md) for the full command surface, [docs/MCP-SETUP.md](MCP-SETUP.md) to wire it into Claude/Cursor/other hosts, and [docs/SIGNING.md](SIGNING.md) for the signing roadmap.
