#!/usr/bin/env python3
"""Golden-script runner for the aioffice CLI.

Each script in fixtures/scripts/*.json is a list of CLI invocations (argv
arrays) with expected envelope assertions. Every step must print exactly one
JSON envelope on stdout; by default each step must be ok:true and exit 0.

Usage:
    python3 fixtures/scripts/run_goldens.py [--cli "<command prefix>"] [script.json ...]

Without --cli the runner uses:  dotnet run --project src/AIOffice.Cli --
(override with the AIOFFICE_CMD environment variable, e.g. a published binary).
Without script paths it runs every *.json in this directory.

The runner creates a fresh temporary workspace per script and passes it via
--workspace, so scripts reference files relative to the workspace root
(e.g. golden-out/hello.docx). Exit code: 0 when all scripts pass, 1 otherwise.
"""

from __future__ import annotations

import argparse
import json
import os
import pathlib
import shlex
import subprocess
import sys
import tempfile

SCRIPTS_DIR = pathlib.Path(__file__).resolve().parent
REPO_ROOT = SCRIPTS_DIR.parent.parent
DEFAULT_CLI = "dotnet run --project src/AIOffice.Cli --"


def fail(message: str) -> None:
    print(f"  FAIL  {message}")


def run_step(cli: list[str], workspace: str, step: dict) -> bool:
    argv = step["argv"]
    expect = step.get("expect", {"ok": True})
    command = cli + ["--workspace", workspace, "--json"] + argv

    proc = subprocess.run(
        command,
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        timeout=300,
    )

    stdout = proc.stdout.strip()
    try:
        envelope = json.loads(stdout)
    except json.JSONDecodeError:
        fail(f"{argv}: stdout is not a single JSON envelope:\n{stdout}\n{proc.stderr}")
        return False

    ok_expected = expect.get("ok", True)
    if envelope.get("ok") is not ok_expected:
        fail(f"{argv}: expected ok={ok_expected}, got: {stdout}")
        return False

    if ok_expected and proc.returncode != 0:
        fail(f"{argv}: ok envelope but exit code {proc.returncode}")
        return False

    if "errorCode" in expect:
        actual = (envelope.get("error") or {}).get("code")
        if actual != expect["errorCode"]:
            fail(f"{argv}: expected error.code={expect['errorCode']}, got {actual}")
            return False

    if "dataContains" in expect:
        blob = json.dumps(envelope.get("data"), ensure_ascii=False)
        if expect["dataContains"] not in blob:
            fail(f"{argv}: data does not contain {expect['dataContains']!r}: {blob[:400]}")
            return False

    print(f"  ok    {' '.join(str(a) for a in argv)}")
    return True


def run_script(cli: list[str], path: pathlib.Path) -> bool:
    script = json.loads(path.read_text(encoding="utf-8"))
    print(f"== {script.get('name', path.stem)}")
    with tempfile.TemporaryDirectory(prefix="aioffice-golden-") as workspace:
        os.makedirs(os.path.join(workspace, "golden-out"), exist_ok=True)
        return all(run_step(cli, workspace, step) for step in script["steps"])


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cli", default=os.environ.get("AIOFFICE_CMD", DEFAULT_CLI),
                        help="command prefix used to invoke aioffice")
    parser.add_argument("scripts", nargs="*", type=pathlib.Path,
                        help="golden scripts to run (default: all *.json here)")
    args = parser.parse_args()

    cli = shlex.split(args.cli, posix=(os.name != "nt"))
    scripts = args.scripts or sorted(SCRIPTS_DIR.glob("*.json"))
    if not scripts:
        print("no golden scripts found", file=sys.stderr)
        return 1

    results = [run_script(cli, path) for path in scripts]
    passed = sum(results)
    print(f"\n{passed}/{len(results)} golden scripts passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
