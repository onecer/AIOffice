#!/usr/bin/env node
'use strict';

// bin/aioffice.js — thin launcher for the AIOffice native binary.
//
// Locates the downloaded binary (bin/aioffice[.exe]). If it is missing (e.g.
// postinstall was skipped, as `npm` does with --ignore-scripts, or under some
// `npx` flows), it runs the install step on demand, then spawns the binary
// with the inherited argv and stdio (full pass-through). The child's exit code
// is propagated. stdio:'inherit' keeps `aioffice mcp` (stdin/stdout JSON-RPC)
// transparent.

const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

const { binaryName } = require('../platform');

const BINARY = path.join(__dirname, binaryName());

function spawnBinary() {
  if (!fs.existsSync(BINARY)) {
    process.stderr.write(
      'aioffice: native binary not found after install. ' +
        'See messages above; you may need to reinstall or download it ' +
        'manually from https://github.com/onecer/AIOffice/releases\n',
    );
    process.exit(1);
  }

  const child = spawn(BINARY, process.argv.slice(2), { stdio: 'inherit' });

  child.on('error', (err) => {
    process.stderr.write(`aioffice: failed to launch binary: ${err.message}\n`);
    process.exit(1);
  });

  // Forward termination signals so Ctrl-C / kill reach the native process.
  for (const sig of ['SIGINT', 'SIGTERM', 'SIGHUP']) {
    process.on(sig, () => {
      if (!child.killed) {
        try {
          child.kill(sig);
        } catch (_) {
          /* ignore */
        }
      }
    });
  }

  child.on('exit', (code, signal) => {
    if (signal) {
      // Re-raise the signal so the parent's exit status reflects it.
      process.kill(process.pid, signal);
      return;
    }
    process.exit(code == null ? 0 : code);
  });
}

async function main() {
  if (!fs.existsSync(BINARY)) {
    // Lazy install (postinstall was skipped or `npx` first run).
    process.stderr.write('aioffice: binary not present, installing...\n');
    try {
      const install = require('../install');
      await install.main();
    } catch (err) {
      process.stderr.write(
        (err && err.message ? err.message : String(err)) + '\n',
      );
      process.exit(1);
    }
  }
  spawnBinary();
}

main();
