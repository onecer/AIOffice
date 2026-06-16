'use strict';

// install.js — postinstall step for the "aioffice" npm package.
//
// Downloads the platform-specific AIOffice binary from the matching GitHub
// release, verifies it against the release's SHA256SUMS file, and installs it
// (executable) into bin/. Uses ONLY Node builtins (https, fs, path, crypto) —
// no third-party dependencies.
//
// Environment overrides (handy for testing + private mirrors):
//   AIOFFICE_DOWNLOAD_VERSION  — release tag to download (default: v{package version})
//   AIOFFICE_DOWNLOAD_BASEURL  — base URL for the release assets
//                                (default: https://github.com/onecer/AIOffice/releases/download)
//
// The full asset URL is: {BASEURL}/{VERSION}/{asset}
// and the checksum file:  {BASEURL}/{VERSION}/SHA256SUMS

const fs = require('fs');
const path = require('path');
const https = require('https');
const crypto = require('crypto');

const { assetName, binaryName } = require('./platform');
const pkg = require('./package.json');

const DEFAULT_BASEURL = 'https://github.com/onecer/AIOffice/releases/download';
const REPO_RELEASES = 'https://github.com/onecer/AIOffice/releases';

const BIN_DIR = path.join(__dirname, 'bin');

// Resolve the release tag. A bare package version "1.8.0" becomes "v1.8.0";
// an explicit override is used verbatim (it may or may not carry a leading v).
function resolveVersion() {
  const override = process.env.AIOFFICE_DOWNLOAD_VERSION;
  if (override && override.trim()) return override.trim();
  return `v${pkg.version}`;
}

function resolveBaseUrl() {
  const override = process.env.AIOFFICE_DOWNLOAD_BASEURL;
  if (override && override.trim()) return override.trim().replace(/\/+$/, '');
  return DEFAULT_BASEURL;
}

// GET a URL into a Buffer, following redirects (GitHub asset URLs 302 to a CDN).
function fetchBuffer(url, redirectsLeft = 5) {
  return new Promise((resolve, reject) => {
    const req = https.get(
      url,
      { headers: { 'User-Agent': `aioffice-npm/${pkg.version}`, Accept: '*/*' } },
      (res) => {
        const { statusCode, headers } = res;
        if (statusCode >= 300 && statusCode < 400 && headers.location) {
          res.resume(); // drain
          if (redirectsLeft <= 0) {
            reject(new Error(`Too many redirects fetching ${url}`));
            return;
          }
          const next = new URL(headers.location, url).toString();
          resolve(fetchBuffer(next, redirectsLeft - 1));
          return;
        }
        if (statusCode !== 200) {
          res.resume();
          reject(
            new Error(
              `Download failed (HTTP ${statusCode}) for ${url}`,
            ),
          );
          return;
        }
        const chunks = [];
        res.on('data', (c) => chunks.push(c));
        res.on('end', () => resolve(Buffer.concat(chunks)));
      },
    );
    req.on('error', reject);
    req.setTimeout(120000, () => {
      req.destroy(new Error(`Download timed out after 120s for ${url}`));
    });
  });
}

function sha256(buf) {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

// Parse a sha256sum-style file: lines of "<hex>␠␠<filename>".
// Returns a Map of filename -> lowercase hex digest.
function parseSha256Sums(text) {
  const map = new Map();
  for (const raw of text.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line) continue;
    const m = line.match(/^([0-9a-fA-F]{64})[\s*]+(.+)$/);
    if (!m) continue;
    map.set(m[2].trim(), m[1].toLowerCase());
  }
  return map;
}

async function main() {
  const version = resolveVersion();
  const baseUrl = resolveBaseUrl();
  const asset = assetName(); // throws a clear error on unsupported platform
  const localName = binaryName();
  const dest = path.join(BIN_DIR, localName);

  const assetUrl = `${baseUrl}/${version}/${asset}`;
  const sumsUrl = `${baseUrl}/${version}/SHA256SUMS`;

  fs.mkdirSync(BIN_DIR, { recursive: true });

  // --- Idempotency: skip if a correct binary is already present. -----------
  // We need the expected digest to know it's correct, so fetch SHA256SUMS
  // first. If the network is unavailable but a binary already exists, trust it
  // (a prior install verified it) rather than failing a re-run.
  let sums;
  try {
    const sumsBuf = await fetchBuffer(sumsUrl);
    sums = parseSha256Sums(sumsBuf.toString('utf8'));
  } catch (err) {
    if (fs.existsSync(dest) && fs.statSync(dest).size > 0) {
      console.error(
        `aioffice: could not fetch SHA256SUMS (${err.message}); ` +
          'a binary is already installed, keeping it.',
      );
      return;
    }
    fail(
      `Could not download the checksum file.\n  ${err.message}`,
      version,
      asset,
      assetUrl,
    );
  }

  const expected = sums.get(asset);
  if (!expected) {
    fail(
      `Release ${version} has no checksum entry for "${asset}".\n` +
        `  Known entries: ${[...sums.keys()].join(', ') || '(none)'}`,
      version,
      asset,
      assetUrl,
    );
  }

  if (fs.existsSync(dest) && fs.statSync(dest).size > 0) {
    const have = sha256(fs.readFileSync(dest));
    if (have === expected) {
      ensureExecutable(dest);
      console.log(
        `aioffice: ${localName} already installed and verified (${version}).`,
      );
      return;
    }
    // Present but wrong (partial/stale download) — re-download below.
    console.error('aioffice: existing binary failed verification, re-downloading.');
  }

  // --- Download the binary. -------------------------------------------------
  console.log(`aioffice: downloading ${asset} (${version})...`);
  let binBuf;
  try {
    binBuf = await fetchBuffer(assetUrl);
  } catch (err) {
    fail(
      `Could not download the binary.\n  ${err.message}`,
      version,
      asset,
      assetUrl,
    );
  }

  // --- Verify SHA256 BEFORE writing the final file. ------------------------
  const actual = sha256(binBuf);
  if (actual !== expected) {
    fail(
      'SHA256 checksum mismatch — the download may be corrupt or tampered ' +
        'with. Nothing was installed.\n' +
        `  expected: ${expected}\n  actual:   ${actual}`,
      version,
      asset,
      assetUrl,
    );
  }

  // Write atomically: temp file -> rename.
  const tmp = `${dest}.download-${process.pid}`;
  fs.writeFileSync(tmp, binBuf);
  fs.renameSync(tmp, dest);
  ensureExecutable(dest);

  console.log(
    `aioffice: installed ${localName} (${version}), SHA256 verified.`,
  );
}

function ensureExecutable(file) {
  if (process.platform !== 'win32') {
    try {
      fs.chmodSync(file, 0o755);
    } catch (_) {
      /* best-effort */
    }
  }
}

function fail(message, version, asset, assetUrl) {
  const lines = [
    '',
    'aioffice: installation failed.',
    `  ${message.replace(/\n/g, '\n  ')}`,
    '',
    'You can download the binary manually from:',
    `  ${assetUrl}`,
    `  (release page: ${REPO_RELEASES}/tag/${version})`,
    'then place it in this package\'s bin/ directory as ' +
      `"${binaryName()}"` +
      (process.platform === 'win32' ? '.' : ' and run chmod +x on it.'),
    '',
    'Env overrides: AIOFFICE_DOWNLOAD_VERSION, AIOFFICE_DOWNLOAD_BASEURL.',
    '',
  ];
  const err = new Error(lines.join('\n'));
  err.handled = true;
  throw err;
}

// Allow `require('./install')` (used by the bin shim) to call main() directly.
module.exports = { main, parseSha256Sums, resolveVersion, resolveBaseUrl };

if (require.main === module) {
  main().catch((err) => {
    process.stderr.write((err && err.message ? err.message : String(err)) + '\n');
    process.exit(1);
  });
}
