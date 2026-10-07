"""Rebuild dist/becc-mecc-calculator-audit.skill from skills/becc-mecc-calculator-audit/.

Run from anywhere:  python build_skill.py

The .skill file is a zip whose top-level folder is becc-mecc-calculator-audit/.
Test-only and generated files (evals/, __pycache__/, *.pyc) are left out.
"""
import os
import zipfile

NAME = "becc-mecc-calculator-audit"
ROOT = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, "skills", NAME)
OUT = os.path.join(ROOT, "dist", NAME + ".skill")
SKIP_DIRS = {"evals", "__pycache__"}
TEXT_EXT = (".md", ".py", ".json", ".txt", ".csv")

if not os.path.isfile(os.path.join(SRC, "SKILL.md")):
    raise SystemExit(f"SKILL.md not found in {SRC}")

os.makedirs(os.path.dirname(OUT), exist_ok=True)
count = 0
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zf:
    for folder, dirs, files in os.walk(SRC):
        dirs[:] = sorted(d for d in dirs if d not in SKIP_DIRS)
        for f in sorted(files):
            if f.endswith(".pyc"):
                continue
            path = os.path.join(folder, f)
            arc = os.path.join(NAME, os.path.relpath(path, SRC)).replace(os.sep, "/")
            if f.endswith(TEXT_EXT):  # same LF line endings whatever git's checkout setting is
                with open(path, "rb") as fh:
                    zf.writestr(arc, fh.read().replace(b"\r\n", b"\n"), zipfile.ZIP_DEFLATED)
            else:
                zf.write(path, arc)
            count += 1

print(f"Built {os.path.relpath(OUT, ROOT)} ({count} files, {os.path.getsize(OUT):,} bytes)")
