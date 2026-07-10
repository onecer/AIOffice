#!/bin/sh
# AIOffice installer — POSIX sh, no bashisms.
#
#   curl -fsSL https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.sh | sh
#
# Downloads the matching prebuilt `aioffice` binary from a GitHub release,
# verifies its SHA256 against the release SHA256SUMS, and installs it.
#
# Environment variables (all optional):
#   VERSION       Release tag to install (e.g. v1.23.1). Default: latest release.
#   AIOFFICE_BIN  Install directory.                    Default: $HOME/.local/bin
#
# Examples:
#   VERSION=v1.23.1 sh install.sh
#   AIOFFICE_BIN=/usr/local/bin sh install.sh
set -eu

REPO="onecer/AIOffice"
# Fallback tag used only when the GitHub "latest" lookup fails and VERSION is unset.
FALLBACK_VERSION="v1.23.1"

# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
err() { printf 'aioffice-install: %s\n' "$1" >&2; }
die() { err "$1"; exit 1; }

have() { command -v "$1" >/dev/null 2>&1; }

# Download $1 (url) to $2 (path) using curl or wget.
download() {
  _url="$1"; _out="$2"
  if have curl; then
    curl -fsSL "$_url" -o "$_out"
  elif have wget; then
    wget -qO "$_out" "$_url"
  else
    die "neither curl nor wget is available; please install one and retry"
  fi
}

# Fetch a URL to stdout.
fetch() {
  if have curl; then
    curl -fsSL "$1"
  elif have wget; then
    wget -qO- "$1"
  else
    die "neither curl nor wget is available; please install one and retry"
  fi
}

# ---------------------------------------------------------------------------
# detect platform -> release asset name
# ---------------------------------------------------------------------------
detect_asset() {
  _os="$(uname -s)"
  _arch="$(uname -m)"

  case "$_os" in
    Darwin) _osname="mac" ;;
    Linux)  _osname="linux" ;;
    *) die "unsupported OS: $_os (this installer handles macOS and Linux; Windows users run dist/install.ps1)" ;;
  esac

  case "$_arch" in
    arm64|aarch64) _archname="arm64" ;;
    x86_64|amd64)  _archname="x64" ;;
    *) die "unsupported CPU architecture: $_arch" ;;
  esac

  ASSET="aioffice-${_osname}-${_archname}"
}

# ---------------------------------------------------------------------------
# resolve version
# ---------------------------------------------------------------------------
resolve_version() {
  if [ "${VERSION:-}" != "" ]; then
    RELEASE_TAG="$VERSION"
    return
  fi
  # Ask the GitHub API for the latest release tag.
  _api="https://api.github.com/repos/${REPO}/releases/latest"
  _tag="$(fetch "$_api" 2>/dev/null | sed -n 's/.*"tag_name"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' | head -n 1 || true)"
  if [ "${_tag:-}" = "" ]; then
    err "could not resolve latest release from GitHub API; falling back to ${FALLBACK_VERSION}"
    _tag="$FALLBACK_VERSION"
  fi
  RELEASE_TAG="$_tag"
}

# ---------------------------------------------------------------------------
# sha256
# ---------------------------------------------------------------------------
sha256_of() {
  if have sha256sum; then
    sha256sum "$1" | awk '{print $1}'
  elif have shasum; then
    shasum -a 256 "$1" | awk '{print $1}'
  else
    die "neither sha256sum nor shasum is available; cannot verify download integrity"
  fi
}

# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
main() {
  detect_asset
  resolve_version

  BASE="https://github.com/${REPO}/releases/download/${RELEASE_TAG}"
  BIN_URL="${BASE}/${ASSET}"
  SUMS_URL="${BASE}/SHA256SUMS"

  INSTALL_DIR="${AIOFFICE_BIN:-$HOME/.local/bin}"
  DEST="${INSTALL_DIR}/aioffice"
  STAGE="${DEST}.new.$$"

  printf 'aioffice-install: %s %s -> %s\n' "$RELEASE_TAG" "$ASSET" "$DEST"

  _tmp="$(mktemp -d "${TMPDIR:-/tmp}/aioffice-install.XXXXXX")" || die "failed to create temp dir"
  # shellcheck disable=SC2064
  trap "rm -rf '$_tmp' '$STAGE'" EXIT INT TERM

  # 1. download binary + checksums
  download "$BIN_URL" "${_tmp}/aioffice" \
    || die "download failed: ${BIN_URL} (check that release ${RELEASE_TAG} exists and has asset ${ASSET})"
  download "$SUMS_URL" "${_tmp}/SHA256SUMS" \
    || die "download failed: ${SUMS_URL}"

  # 2. verify SHA256 against the published SHA256SUMS line for this asset
  _expected="$(awk -v a="$ASSET" '$2 == a {print $1}' "${_tmp}/SHA256SUMS" | head -n 1)"
  [ "${_expected:-}" != "" ] || die "no SHA256 entry for ${ASSET} in SHA256SUMS"
  _actual="$(sha256_of "${_tmp}/aioffice")"
  if [ "$_expected" != "$_actual" ]; then
    err "SHA256 mismatch for ${ASSET}"
    err "  expected: ${_expected}"
    err "  actual:   ${_actual}"
    die "refusing to install a binary that failed integrity verification"
  fi
  printf 'aioffice-install: sha256 verified (%s)\n' "$_actual"

  # 3. install — stage the new binary in the install dir, then atomically rename it
  #    into place. A plain `cp` over an existing binary REUSES that file's inode; on
  #    macOS the kernel keeps a code-signature cache keyed by inode, so once the bytes
  #    change the cached signature no longer matches and the upgraded binary HANGS /
  #    is killed on first run. `mv` (rename) swaps in a fresh inode, sidestepping the
  #    stale cache — and it is safe to do while the OLD binary is still running.
  #    Staging inside INSTALL_DIR keeps the rename on one filesystem (atomic).
  mkdir -p "$INSTALL_DIR" || die "cannot create install dir: ${INSTALL_DIR}"
  cp "${_tmp}/aioffice" "$STAGE" || die "failed to write into ${INSTALL_DIR} (need write permission, or set AIOFFICE_BIN)"
  chmod +x "$STAGE" || die "failed to mark the new binary executable"

  # 4. macOS: clear quarantine + provenance xattrs, and (re-)establish an ad-hoc
  #    signature on the staged binary so the fresh inode gets a clean first-run
  #    Gatekeeper/AMFI assessment. Integrity is already guaranteed by the SHA256
  #    check above; we ad-hoc re-sign UNLESS the download already carries a real
  #    (Developer ID / notarized) signature — detected by an `Authority=` line —
  #    which we leave untouched.
  if [ "$(uname -s)" = "Darwin" ]; then
    if have xattr; then xattr -c "$STAGE" >/dev/null 2>&1 || true; fi
    if have codesign && ! codesign -dvv "$STAGE" 2>&1 | grep -q '^Authority='; then
      codesign --force --sign - "$STAGE" >/dev/null 2>&1 || true
    fi
  fi

  # 5. atomically replace the old binary with the freshly-staged one (new inode).
  mv -f "$STAGE" "$DEST" || die "failed to install ${DEST}"

  # 6. PATH hint
  case ":${PATH}:" in
    *":${INSTALL_DIR}:"*) : ;;
    *)
      printf '\naioffice-install: NOTE %s is not on your PATH.\n' "$INSTALL_DIR"
      printf '  Add this line to your shell profile (~/.bashrc, ~/.zshrc, ~/.profile):\n'
      printf '    export PATH="%s:$PATH"\n\n' "$INSTALL_DIR"
      ;;
  esac

  # 7. confirm install by running the binary
  printf 'aioffice-install: installed. '
  if "$DEST" version 2>/dev/null; then
    :
  else
    printf '\naioffice-install: installed to %s (run `aioffice version` to confirm).\n' "$DEST"
  fi
}

main "$@"
