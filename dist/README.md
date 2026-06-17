# AIOffice distribution

Packaging and one-line install artifacts for the `aioffice` binary. None of this
touches the C#/.NET source tree under `src/`.

| File | Purpose |
| --- | --- |
| `install.sh` | POSIX one-line installer for macOS + Linux. |
| `install.ps1` | PowerShell installer for Windows. |
| `Formula/aioffice.rb` | Homebrew formula (for the `onecer/tap` tap). |

All three pull prebuilt binaries from the GitHub release assets
`https://github.com/onecer/AIOffice/releases/download/v{version}/{asset}` and
verify each download's SHA256 against the release `SHA256SUMS` before installing.

Platform → asset map:

| OS | Arch | Asset |
| --- | --- | --- |
| macOS | arm64 | `aioffice-mac-arm64` |
| macOS | x64 | `aioffice-mac-x64` |
| Linux | arm64 | `aioffice-linux-arm64` |
| Linux | x64 | `aioffice-linux-x64` |
| Windows | x64 | `aioffice-win-x64.exe` |
| Windows | arm64 | `aioffice-win-arm64.exe` |

## End-user install

macOS / Linux:

```sh
curl -fsSL https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.sh | sh
```

Pin a version or install dir:

```sh
VERSION=v1.9.0 AIOFFICE_BIN=/usr/local/bin \
  sh -c "$(curl -fsSL https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.sh)"
```

Windows (PowerShell):

```powershell
irm https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.ps1 | iex
```

Homebrew (once the tap exists):

```sh
brew install onecer/tap/aioffice
```

## Publish checklist (human-only — these need credentials)

The agent built and locally tested everything below but did NOT run any
outward-facing command. Run these yourself after tagging `v1.9.0`.

### 1. Cut the v1.9.0 release

Tag `v1.9.0` and upload the 6 binaries + `SHA256SUMS` (your existing release
flow). The install scripts default to the GitHub "latest" release, so they pick
up v1.9.0 automatically once it is published.

### 2. Homebrew tap

1. Create a public repo `onecer/homebrew-tap`.
2. Copy `dist/Formula/aioffice.rb` into it as `Formula/aioffice.rb`.
3. Fill the four `sha256` values (each marked `TODO(human)`) from the v1.9.0
   `SHA256SUMS`:

   ```sh
   gh release download v1.9.0 -R onecer/AIOffice -p SHA256SUMS -O - | sort
   ```

   Map asset → sha256 line (the comment on each formula line names the asset).
   The values currently in the formula are the **v1.5.0** examples — replace
   them. The formula already has `version "1.9.0"`.
4. Commit and push the tap repo. Verify with:

   ```sh
   brew install onecer/tap/aioffice
   aioffice version
   ```

### 3. Make `releases/download/...` URLs publicly resolvable

The install scripts use the unauthenticated browser download URL
(`https://github.com/onecer/AIOffice/releases/download/...`). While the repo is
**private**, that URL returns 404 for anyone without a token — only
`gh release download` (authenticated) works. Make the repo (or at least its
releases) public before advertising the one-liner, or the curl/irm install will
fail for end users.

## How this was tested locally

- `ruby -c` and `brew style` on the formula → no offenses.
- `sh -n` on `install.sh` → syntax OK.
- `install.sh` run end-to-end against a local HTTP mirror of the **real v1.5.0**
  `aioffice-mac-arm64` + `SHA256SUMS`: detected platform, downloaded, verified
  sha256 `63a5d98…`, installed executable, printed PATH hint, and
  `aioffice version` returned the valid JSON envelope (`"version":"1.5.0"`).
- Tamper test: appending bytes to the served binary made the sha256 mismatch →
  installer refused and wrote nothing.

The only reason the scripts could not be tested against the live public URL is
that the repo is private today (see step 3); the URL format itself is verified
correct against the v1.5.0 asset set.
