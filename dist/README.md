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
VERSION=v1.20.0 AIOFFICE_BIN=/usr/local/bin \
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

## Release automation (what runs on a `v*` tag)

Tagging `vX.Y.Z` fans out to three workflows — no manual outward-facing step is
needed once the two publish secrets exist:

| Workflow | Does | Secret |
| --- | --- | --- |
| `release.yml`      | builds the 6 binaries + `SHA256SUMS`, creates the GitHub release | `GITHUB_TOKEN` (built in) |
| `npm-publish.yml`  | waits for `SHA256SUMS`, publishes the `aioffice` npm shim     | `NPM_TOKEN` (automation token) |
| `homebrew-tap.yml` | waits for `SHA256SUMS`, rewrites the formula's version + 4 sha256s and pushes `Formula/aioffice.rb` to the tap, then syncs `dist/Formula/aioffice.rb` back to `main` | `HOMEBREW_TAP_TOKEN` (PAT, push to the tap) |

Each publish workflow is **gated on its secret**: absent the secret (forks, or
before setup) it logs a skip and stays green. The install scripts default to the
GitHub "latest" release, so they pick up the new version automatically.

`scripts/update-formula.py` is the formula rewriter (version + per-platform
sha256 from `SHA256SUMS`, anchored by each line's `# asset:` comment). To refresh
the tap **by hand** if the automation is ever off:

```sh
gh release download vX.Y.Z -R onecer/AIOffice -p SHA256SUMS
python3 scripts/update-formula.py --version X.Y.Z --sums SHA256SUMS --formula dist/Formula/aioffice.rb
# then commit dist/Formula/aioffice.rb and copy it into onecer/homebrew-tap as Formula/aioffice.rb
```

Verify either way with:

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
