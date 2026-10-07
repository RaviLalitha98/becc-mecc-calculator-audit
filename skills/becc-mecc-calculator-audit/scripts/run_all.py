#!/usr/bin/env python3
"""Run the full automated evidence pass and write an index.

  python run_all.py WORKBOOK --out audit_out [--password PW] [--twin "Super structure=Sub structure"]
                    [--previous OLD_VERSION.xlsx]

Runs: inventory → formula_scan → formula_patterns → lookup_audit →
aggregation_check → (optional) twin-sheet comparison → (optional) version
comparison against a previous workbook. Writes audit_out/00_INDEX.md.

Everything produced here is *candidate evidence*. The auditor still has to
open the cells, trace them, and decide what is a defect.
"""
from __future__ import annotations

import argparse
import hashlib
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def run(args, log):
    t = time.time()
    p = subprocess.run([sys.executable] + args, capture_output=True, text=True, cwd=HERE)
    line = (p.stdout.strip().splitlines() or [""])[-1]
    log.append(f"- `{os.path.basename(args[0])}` ({time.time() - t:.1f}s): {line}")
    if p.returncode != 0:
        log.append(f"  - **FAILED**: {p.stderr.strip()[-800:]}")
    return p.returncode == 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--out", default="audit_out")
    ap.add_argument("--password")
    ap.add_argument("--twin", action="append", default=[], help="sheet pair A=B inside the same workbook")
    ap.add_argument("--previous", help="earlier version of the calculator to diff against")
    a = ap.parse_args()
    wb = os.path.abspath(a.workbook)
    out = os.path.abspath(a.out)
    os.makedirs(out, exist_ok=True)
    pw = ["--password", a.password] if a.password else []
    h0 = sha256(wb)
    log = [f"# Automated evidence pass — {os.path.basename(wb)}\n",
           f"SHA-256 at start: `{h0}`\n",
           "All outputs are *candidates* for review, not confirmed findings.\n"]
    for script in ("inventory.py", "formula_scan.py", "formula_patterns.py", "lookup_audit.py", "aggregation_check.py"):
        run([os.path.join(HERE, script), wb, "--out", out] + pw, log)
    for i, pair in enumerate(a.twin):
        run([os.path.join(HERE, "compare_versions.py"), wb, wb, "--pair", pair, "--out", os.path.join(out, f"twin_{i + 1}")], log)
    if a.previous:
        run([os.path.join(HERE, "compare_versions.py"), os.path.abspath(a.previous), wb, "--out",
             os.path.join(out, "version_diff")], log)
    h1 = sha256(wb)
    log.append(f"\nSHA-256 at end: `{h1}` — {'unchanged' if h0 == h1 else '**CHANGED — investigate**'}")
    log.append("\n## Files\n")
    for root, _, files in os.walk(out):
        for f in sorted(files):
            log.append(f"- {os.path.relpath(os.path.join(root, f), out)}")
    with open(os.path.join(out, "00_INDEX.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(log) + "\n")
    print("\n".join(log))


if __name__ == "__main__":
    main()
