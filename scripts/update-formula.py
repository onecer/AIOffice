#!/usr/bin/env python3
"""Rewrite the Homebrew formula's version + per-platform sha256 from a release's
SHA256SUMS file.

Used by .github/workflows/homebrew-tap.yml to auto-update the tap
(onecer/homebrew-tap, Formula/aioffice.rb) on each v* release, replacing the
manual per-release SHA refresh. Pure text rewrite, idempotent, no deps.

Each sha256 line in the formula is anchored by the comment that precedes it,
e.g.  `# asset: aioffice-mac-arm64 (v1.14.0 SHA256SUMS)` -> the next `sha256`
line gets aioffice-mac-arm64's checksum from SHA256SUMS. The version string and
those comment tags are rewritten too; the URLs interpolate `v#{version}` in Ruby
so they follow automatically.

  python3 scripts/update-formula.py --version 1.15.0 \
      --sums SHA256SUMS --formula dist/Formula/aioffice.rb
"""
import argparse
import re
import sys

ASSET_RE = re.compile(r"#\s*asset:\s*(aioffice-[a-z0-9.-]+)")
SHA_RE = re.compile(r'^(\s*sha256\s*")[0-9a-fA-F]{64}(".*)$')
VERSION_RE = re.compile(r'^(\s*version\s*")[^"]*(".*)$')
VTAG_RE = re.compile(r"v\d+\.\d+\.\d+")
EXPECTED_SHA_LINES = 4


def load_sums(path):
    """Parse a `<sha>  <asset>` SHA256SUMS file into {asset: sha}."""
    sums = {}
    with open(path, encoding="utf-8") as handle:
        for line in handle:
            parts = line.split()
            if len(parts) >= 2:
                sums[parts[-1]] = parts[0].lower()
    if not sums:
        sys.exit(f"error: {path} parsed to zero entries")
    return sums


def rewrite(formula_path, version, sums):
    with open(formula_path, encoding="utf-8") as handle:
        lines = handle.readlines()

    out = []
    pending_asset = None
    rewritten_shas = 0
    for line in lines:
        asset_match = ASSET_RE.search(line)
        if asset_match:
            pending_asset = asset_match.group(1)
            out.append(VTAG_RE.sub(f"v{version}", line))
            continue

        version_match = VERSION_RE.match(line)
        if version_match:
            out.append(f"{version_match.group(1)}{version}{version_match.group(2)}\n")
            continue

        sha_match = SHA_RE.match(line)
        if sha_match and pending_asset is not None:
            sha = sums.get(pending_asset)
            if sha is None:
                sys.exit(f"error: SHA256SUMS has no entry for {pending_asset}")
            out.append(f"{sha_match.group(1)}{sha}{sha_match.group(2)}\n")
            rewritten_shas += 1
            pending_asset = None
            continue

        out.append(line)

    if rewritten_shas != EXPECTED_SHA_LINES:
        sys.exit(
            f"error: expected to rewrite {EXPECTED_SHA_LINES} sha256 lines, "
            f"rewrote {rewritten_shas} (is the formula's `# asset:` anchoring intact?)"
        )

    with open(formula_path, "w", encoding="utf-8") as handle:
        handle.writelines(out)
    print(f"formula updated: version {version}, {rewritten_shas} sha256 lines")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--version", required=True, help="package version, e.g. 1.15.0 (no leading v)")
    parser.add_argument("--sums", required=True, help="path to the release SHA256SUMS file")
    parser.add_argument("--formula", required=True, help="path to the Homebrew formula to rewrite")
    args = parser.parse_args()

    version = args.version.lstrip("v")
    rewrite(args.formula, version, load_sums(args.sums))


if __name__ == "__main__":
    main()
