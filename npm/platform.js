'use strict';

// platform.js — maps the running platform to the GitHub release asset name and
// to the local binary filename. This is the single source of truth for the
// platform mapping; install.js and bin/aioffice.js both import it.
//
// Asset map (must match the names of the assets attached to each GitHub
// release v{version}):
//   darwin + arm64 -> aioffice-mac-arm64
//   darwin + x64   -> aioffice-mac-x64
//   linux  + x64   -> aioffice-linux-x64
//   linux  + arm64 -> aioffice-linux-arm64
//   win32  + x64   -> aioffice-win-x64.exe
//   win32  + arm64 -> aioffice-win-arm64.exe

// platform -> arch -> release asset name
const ASSETS = {
  darwin: {
    arm64: 'aioffice-mac-arm64',
    x64: 'aioffice-mac-x64',
  },
  linux: {
    x64: 'aioffice-linux-x64',
    arm64: 'aioffice-linux-arm64',
  },
  win32: {
    x64: 'aioffice-win-x64.exe',
    arm64: 'aioffice-win-arm64.exe',
  },
};

/**
 * The local filename the downloaded binary is stored under, inside bin/.
 * Windows needs a .exe extension so the OS will execute it; every other
 * platform uses a plain "aioffice".
 * @param {string} [platform=process.platform]
 * @returns {string}
 */
function binaryName(platform = process.platform) {
  return platform === 'win32' ? 'aioffice.exe' : 'aioffice';
}

/**
 * Resolve the GitHub release asset name for the given platform/arch.
 * Throws a clear, actionable error on an unsupported combination.
 * @param {string} [platform=process.platform]
 * @param {string} [arch=process.arch]
 * @returns {string} the release asset filename
 */
function assetName(platform = process.platform, arch = process.arch) {
  const byArch = ASSETS[platform];
  const asset = byArch && byArch[arch];
  if (!asset) {
    const supported = Object.entries(ASSETS)
      .flatMap(([p, m]) => Object.keys(m).map((a) => `${p}/${a}`))
      .join(', ');
    throw new Error(
      `aioffice: unsupported platform "${platform}/${arch}". ` +
        `Supported: ${supported}. ` +
        'If you believe this platform should be supported, please open an ' +
        'issue at https://github.com/onecer/AIOffice/issues',
    );
  }
  return asset;
}

module.exports = { ASSETS, assetName, binaryName };
