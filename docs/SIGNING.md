# Code signing & notarization · 代码签名与公证

> **中文摘要** — **现状：AIOffice 的发行二进制目前尚未签名 / 公证。** 这是诚实的说明，不是承诺。本文给出走向签名的完整路线：macOS 用 Developer ID 应用证书 + 强化运行时（hardened runtime）+ `com.apple.security.cclr.allow-jit` 入口权限（裁剪过的 CoreCLR 需要 JIT）+ `notarytool` 公证；Windows 用 OV（或 EV）代码签名证书。也说明这些步骤将在 `release.yml` 的哪个位置插入、需要哪些 CI secrets。在签名落地之前，用户侧的临时绕过方法是 macOS 上执行 `xattr -d com.apple.quarantine`（一行脚本与 Homebrew 已自动处理）。

## Status · 现状

**The published binaries are currently _unsigned_ and _not notarized_** on every platform. There is no Developer ID signature on the macOS builds, no notarization ticket, and no Authenticode signature on the Windows `.exe` files. This document is the roadmap to change that, written honestly so users and the maintainer know exactly where things stand and what each step requires.

Consequences today:
- **macOS** — a directly downloaded binary is quarantined; Gatekeeper blocks it on first run ("developer cannot be verified"). Workaround below.
- **Windows** — SmartScreen may warn on first run of an unsigned `.exe`; users click *More info → Run anyway*.
- **Linux** — no OS-level signature gate; integrity is established via the published `SHA256SUMS` (see [INSTALL.md](INSTALL.md#verify-your-download)).

### User-facing workaround until signing lands · 临时绕过

The integrity story today is **SHA256 verification against the release `SHA256SUMS`**, which every install path performs. For the macOS Gatekeeper block specifically:

```sh
xattr -d com.apple.quarantine /path/to/aioffice
```

The `install.sh` one-line installer and the Homebrew formula already strip the quarantine attribute for you; only **direct downloads** need this manual step. Alternatively: **System Settings → Privacy & Security → Open Anyway** after the first blocked launch.

### Installer note: replace the inode, don't overwrite it · 升级时换 inode，别原地覆盖

macOS keeps a code-signature validation cache **keyed by inode**. Overwriting an installed binary *in place* (`cp`/`>` onto the existing file) reuses its inode, so the cached signature from the old bytes no longer matches the new ones and AMFI kills/hangs the process on next launch — a real upgrade failure (fixed in `dist/install.sh`, commit `12eec48`). Rule for any updater: **never overwrite the binary in place** — write a sibling temp file and `mv` it over the destination (atomic same-dir rename = fresh inode), or `rm` then write. `mv` is also safe while the old binary is still running. `npm/install.js` (`fs.renameSync`) and the fixed `install.sh` / `install.ps1` all do this; Homebrew / `.pkg` / `.dmg` replace the whole file and get a fresh inode for free.

**Interaction with signing (revisit when signing lands):** the fixed `install.sh` ad-hoc re-signs the staged macOS binary **only when it carries no real `Authority=` (Developer ID / notarized) signature** — so a properly signed, notarized release binary is left untouched (an ad-hoc re-sign would *strip* the Developer ID signature and its notarization, reintroducing the Gatekeeper prompt). Once real signing ships: confirm the installer takes the skip branch on the signed asset (`codesign -dvv aioffice-mac-arm64` shows an `Authority=` line), and consider dropping the ad-hoc fallback if it's no longer wanted.

---

## macOS: Developer ID signing + notarization · macOS 签名与公证

The target is a hardened-runtime, Developer-ID-signed, notarized, stapled binary that runs without any quarantine prompt.

### 1. Prerequisites

- An Apple Developer Program membership.
- A **Developer ID Application** signing certificate (not "Apple Distribution" — that's for the App Store) in the build keychain.
- A notarization credential: an **App Store Connect API key** (`.p8` + key id + issuer id) is the CI-friendly choice for `notarytool`.

### 2. The allow-jit entitlement (required) · 必需的 allow-jit 入口权限

The single-file binaries are **trimmed self-contained CoreCLR** builds. Trimmed CoreCLR still uses the JIT at runtime, so under the hardened runtime the process needs the JIT-allowing entitlement or it will crash on launch with a code-signing/JIT error. Create `build/aioffice.entitlements`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <!-- Trimmed CoreCLR JITs at runtime; the hardened runtime needs this. -->
  <key>com.apple.security.cs.allow-jit</key>
  <true/>
  <!-- Allow unsigned executable memory / DYLD env as a safety net for the runtime. -->
  <key>com.apple.security.cs.allow-unsigned-executable-memory</key>
  <true/>
  <key>com.apple.security.cs.disable-library-validation</key>
  <true/>
</dict>
</plist>
```

> Note: the canonical entitlement key is `com.apple.security.cs.allow-jit`. Start with `allow-jit`; only keep the two looser keys if a JIT/codesign launch failure persists after testing on a clean machine. Verify on hardware that has never seen the binary before (quarantine + Gatekeeper actually engage there).

### 3. Sign (per macOS asset) · 签名

```sh
# DEVID = "Developer ID Application: Your Name (TEAMID)"
codesign --force --timestamp --options runtime \
  --entitlements build/aioffice.entitlements \
  --sign "$DEVID" \
  release/aioffice-mac-arm64

codesign --verify --strict --verbose=2 release/aioffice-mac-arm64
codesign --display --entitlements - release/aioffice-mac-arm64   # confirm allow-jit present
```

Repeat for `aioffice-mac-x64`. `--timestamp` (secure timestamp) and `--options runtime` (hardened runtime) are both required for notarization to pass.

### 4. Notarize with notarytool · 公证

Notarization needs the submission wrapped in a zip (a bare Mach-O can't be submitted directly):

```sh
# store the credential once (CI: pass key/issuer inline instead, see below)
xcrun notarytool store-credentials AIOFFICE_NOTARY \
  --key   AuthKey_XXXXX.p8 \
  --key-id   "$ASC_KEY_ID" \
  --issuer   "$ASC_ISSUER_ID"

for asset in aioffice-mac-arm64 aioffice-mac-x64; do
  ditto -c -k --keepParent "release/$asset" "release/$asset.zip"
  xcrun notarytool submit "release/$asset.zip" \
    --keychain-profile AIOFFICE_NOTARY --wait
done
```

### 5. Stapling caveat · 装订说明

`xcrun stapler staple` attaches the notarization ticket to *bundles, disk images, and installer packages* — **not** to a bare Mach-O executable. For a standalone CLI binary there is nothing to staple; Gatekeeper performs an **online** notarization check on first run instead. Two honest options:

1. **Ship the bare binary** (matches today's asset layout). It is signed + notarized, so Gatekeeper clears it online with no prompt — but the very first launch needs network. This is the simplest path and keeps the `SHA256SUMS` / asset names unchanged.
2. **Ship a stapled `.pkg` or `.dmg`** for macOS so the ticket can be stapled and first launch works fully offline. This changes the macOS asset shape and the installer/Homebrew flows, so it's a larger, separate decision.

Recommendation: do (1) first (smallest change, removes the quarantine prompt), and only move to (2) if offline-first-launch is a hard requirement.

---

## Windows: Authenticode signing · Windows 代码签名

Goal: an Authenticode signature on `aioffice-win-x64.exe` and `aioffice-win-arm64.exe` so SmartScreen trusts them (reputation builds over time; an EV cert gets instant SmartScreen trust).

### Certificate options

- **OV (Organization Validation) code-signing certificate** — cheaper; SmartScreen reputation accrues with download volume over time. As of the 2023 baseline-requirement change, OV private keys must live on an **FIPS 140-2 hardware token / HSM** (or a cloud HSM such as Azure Key Vault), so signing happens against the HSM, not a local `.pfx`.
- **EV (Extended Validation) code-signing certificate** — pricier and always HSM-backed, but grants **immediate** SmartScreen reputation.

### Signing command (Azure-Key-Vault-backed, CI-friendly)

```powershell
# AzureSignTool (dotnet tool) signs against a cloud HSM without exporting the key.
AzureSignTool sign `
  --azure-key-vault-url        "$env:AKV_URL" `
  --azure-key-vault-client-id  "$env:AKV_CLIENT_ID" `
  --azure-key-vault-tenant-id  "$env:AKV_TENANT_ID" `
  --azure-key-vault-client-secret "$env:AKV_CLIENT_SECRET" `
  --azure-key-vault-certificate   "$env:AKV_CERT_NAME" `
  --timestamp-rfc3161 http://timestamp.digicert.com `
  --file-digest sha256 `
  release/aioffice-win-x64.exe

signtool verify /pa release/aioffice-win-x64.exe
```

A timestamp server (`/tr` with `signtool`, or `--timestamp-rfc3161`) is required so signatures stay valid after the cert expires.

---

## Where signing slots into CI · 在 CI 中的位置

All signing fits into [`.github/workflows/release.yml`](../.github/workflows/release.yml), **between** the existing *"Publish single-file binaries (6 rids)"* step (which produces `release/aioffice-*`) and the *"Checksums"* step (so `SHA256SUMS` is computed over the **signed** bytes). Concretely:

1. **Publish single-file binaries (6 rids)** — unchanged; produces the six `release/aioffice-*` assets.
2. **→ NEW: Sign macOS assets** — import the Developer ID cert into a temporary keychain, `codesign` both `aioffice-mac-*` with the entitlements above, then `notarytool submit --wait`. Runs on the existing `macos-latest` runner.
3. **→ NEW: Sign Windows assets** — `AzureSignTool` / `signtool` over both `aioffice-win-*.exe`. (The release job currently runs on macOS; Windows signing via Azure Key Vault works cross-platform with AzureSignTool, or split this into a dedicated `windows-latest` job that signs and re-uploads its two assets.)
4. **Checksums** — unchanged, but now hashes the **signed** binaries.
5. **Create GitHub release** — unchanged.

Keep each signing step **gated on its secrets being present** (same pattern as the npm-publish gate) so forks and secret-less runs still produce working unsigned releases instead of failing the pipeline.

### Secrets to add · 所需 CI Secrets

| Secret | Purpose |
|---|---|
| `MACOS_CERT_P12_BASE64` | Base64 of the Developer ID Application cert (`.p12`) to import into a CI keychain |
| `MACOS_CERT_PASSWORD` | Password for that `.p12` |
| `MACOS_KEYCHAIN_PASSWORD` | Throwaway password for the temporary CI keychain |
| `APPLE_ASC_KEY_ID` / `APPLE_ASC_ISSUER_ID` / `APPLE_ASC_KEY_P8_BASE64` | App Store Connect API key for `notarytool` |
| `APPLE_DEVELOPER_ID` | The `"Developer ID Application: … (TEAMID)"` identity string |
| `AKV_URL` / `AKV_CLIENT_ID` / `AKV_TENANT_ID` / `AKV_CLIENT_SECRET` / `AKV_CERT_NAME` | Azure Key Vault credentials for Windows Authenticode signing |

Until these secrets exist, releases remain unsigned and the SHA256-verification + quarantine-strip path above is the supported integrity story.

---

## Summary · 小结

- **Today:** unsigned, not notarized; integrity via `SHA256SUMS`; macOS users strip quarantine (`xattr -d com.apple.quarantine`), automated for `install.sh` + Homebrew.
- **macOS path:** Developer ID + hardened runtime + **`com.apple.security.cs.allow-jit`** (trimmed CoreCLR needs JIT) + `notarytool` (bare binary → online Gatekeeper check; `.pkg`/`.dmg` if offline-first-launch is required).
- **Windows path:** OV (HSM-backed) or EV Authenticode signing with a timestamp.
- **CI:** new gated steps in `release.yml` between *Publish* and *Checksums*; secrets listed above. Nothing is done yet — this is the plan, not a claim that it's shipped.
