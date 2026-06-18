# aioffice npm package — build & publish notes

This directory is a self-contained npm package whose only job is to download the
correct AIOffice native binary for the user's platform from the GitHub release
and verify it by SHA256. It contains **no** third-party dependencies and uses
only Node builtins (`https`, `fs`, `path`, `crypto`).

## What's in here

| File             | Role                                                                 |
| ---------------- | -------------------------------------------------------------------- |
| `package.json`   | name `aioffice`, version `1.12.0`, `bin`, `postinstall`, `files`.     |
| `platform.js`    | Single source of truth: platform/arch -> release asset + local name. |
| `install.js`     | Postinstall: download + SHA256-verify + chmod the binary into `bin/`.|
| `bin/aioffice.js`| Thin launcher: lazy-installs if needed, spawns binary, stdio pass-through. |
| `README.md`      | User-facing install + MCP usage + env overrides.                    |
| `.npmignore`     | Guards against publishing the downloaded binary.                     |

## The package name

The name is **`aioffice`** (scope-free). It is set in two obvious places only —
`package.json` `"name"` and the URLs in `README.md`. The code itself never hard-
codes the package name (it reads `pkg.version` from `package.json`). If the
scope-free name turns out to be taken on the public registry, the human can
rename to **`@onecer/aioffice`** by editing `package.json` `"name"` (and the
README install commands); nothing else needs to change.

## Versioning

`package.json` is at `1.12.0`. On install, `install.js` downloads release
`v{package.version}` = **`v1.12.0`** by default. That tag does not exist yet — the
human tags it (with its 6 binaries + `SHA256SUMS`) as part of this release. Once
tagged, `npm install -g aioffice` will Just Work with no env overrides.

## How it was tested locally (against the existing v1.5.0 release)

Because `v1.12.0` is not tagged yet, and because the `onecer/AIOffice` repo is
**private** (anonymous `https://github.com/.../releases/download/...` returns
404 — a real public release will not), testing used the real **v1.5.0** assets
served from a local HTTPS mirror:

1. Downloaded the real `v1.5.0` `SHA256SUMS` + `aioffice-mac-arm64` with
   authenticated `gh release download` into a local dir; confirmed the bytes'
   SHA256 matches the published `SHA256SUMS` (`63a5d987…aee6c`).
2. Served that dir over local HTTPS and ran:

   ```sh
   AIOFFICE_DOWNLOAD_VERSION=v1.5.0 \
   AIOFFICE_DOWNLOAD_BASEURL=https://127.0.0.1:8799 \
   NODE_TLS_REJECT_UNAUTHORIZED=0 \
   node install.js
   ```

   (`NODE_TLS_REJECT_UNAUTHORIZED=0` is only for the self-signed local cert —
   the real GitHub URL has valid TLS and needs no such flag.)

Results — all green:

- Download + SHA256 verify + `chmod 755` -> `bin/aioffice` (`-rwxr-xr-x`).
- `./bin/aioffice doctor` -> `ok:true`, all 3 handlers (docx/xlsx/pptx) `ready`.
- Idempotent re-run -> "already installed and verified", no re-download.
- Corrupted local binary -> detected, re-downloaded, re-verified.
- Tampered download (wrong bytes vs. listed sum) -> **aborts, installs nothing**,
  exit code 1, prints manual-download URL.
- Unsupported platform (`sunos/sparc`) -> clear thrown error listing supported
  platforms.
- Lazy install via shim (`node bin/aioffice.js …` with no binary present) ->
  auto-installs then runs (so `npx aioffice` works even if postinstall skipped).
- MCP stdio pass-through (`aioffice mcp` initialize handshake) -> transparent,
  identical JSON-RPC reply via shim and direct binary.
- Exit-code propagation -> shim returns the binary's exact code (e.g. 4 on a
  read error).
- `npm pack` -> 5 files, 6.3 kB; the downloaded binary is **not** in the tarball.
- `npm install <tarball>` into an isolated prefix -> postinstall downloaded +
  verified the binary, linked the `aioffice` bin; `aioffice doctor` and a real
  `create demo.docx` + `read` round-trip both succeeded.

When the real `v1.12.0` tag exists, the **same flow runs with no env overrides**:
`install.js` downloads `https://github.com/onecer/AIOffice/releases/download/v1.12.0/<asset>`
and the sibling `SHA256SUMS`.

## Publishing (for the human to run — NOT done here)

These are intentionally **not** executed by the build (no credentials are used).

```sh
cd npm

# 0. Make sure the v1.12.0 GitHub release exists with all 6 binaries + SHA256SUMS.

# 1. Sanity-check what will be published.
npm pack --dry-run         # should list exactly: README.md, bin/aioffice.js,
                           # install.js, package.json, platform.js (5 files)

# 2. Optional: dry-run the publish.
npm publish --dry-run

# 3. Publish (scope-free, public).
npm publish --access public

# If the name "aioffice" is taken, first edit package.json "name" to
# "@onecer/aioffice" (and the README install commands), then:
#   npm publish --access public
```

After publish, verify a clean machine:

```sh
npm install -g aioffice && aioffice doctor
# or
npx aioffice doctor
```

## Notes / gotchas

- `files` lists `bin/aioffice.js` explicitly (not the whole `bin/` dir) and
  `.npmignore` excludes `bin/aioffice`, `bin/aioffice.exe`, and
  `bin/*.download-*`, so a binary downloaded during local testing can never leak
  into a published tarball.
- The download follows GitHub's 302 redirect to its asset CDN (handled in
  `fetchBuffer`).
- `install.js` writes the binary atomically (temp file -> rename) and verifies
  the SHA256 **before** the rename, so a failed/aborted install never leaves a
  half-written or unverified binary in place.
