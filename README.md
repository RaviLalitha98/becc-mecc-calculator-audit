# BECC / MECC Calculator Audit

A Claude skill that checks whether an Excel embodied-carbon calculator gives the right answers.

## What it does

Embodied-carbon calculators, such as a BECC (Building Embodied Carbon Calculator) or a
MECC (Material Embodied Carbon Calculator), are large spreadsheets with thousands of formulas.
A single wrong cell reference, unit mix-up or broken lookup can silently change a building's
carbon total by orders of magnitude, and it's very hard to spot by eye.

This tool gives Claude a structured audit method plus Python scripts that scan the whole
workbook. It finds problems like these:

- **Broken or inconsistent formulas.** One row points to the wrong cell, sheet or range, unlike its neighbours.
- **Wrong lookups.** A material picks up the wrong emission factor, a range is cut short, or duplicates exist.
- **Unit errors.** For example, an emission factor per m³ multiplied by a quantity in kg.
- **Totals that miss or double-count rows.**
- **Life-cycle stages (A1–A3, A4, A5, B, C, D)** mixed up or left out.
- **Hidden sheets, hidden logic and protection** that hide problems.
- **Excel-version issues.** Functions that break in older Excel.
- **Version changes.** What changed between v7 and v8, and whether it was intended.

Every issue comes with evidence: the sheet and cell, the current formula, why it's wrong,
the impact, a suggested fix, and a severity (Critical / High / Medium / Low).

**What you get:** an evaluation report (shown in the chat, or as a Word or PDF file, your
choice), a findings register (Excel + Markdown), and an evidence folder with the raw scan results. The original workbook is never modified.

**What it doesn't do:** it doesn't confirm that the emission factors or the methodology are
right. Those are flagged for the methodology owner to confirm.

## Repository layout

```
.
├── skills/
│   └── becc-mecc-calculator-audit/   Skill source code
│       ├── SKILL.md                  Audit method Claude follows (start here)
│       ├── scripts/                  Python tools that scan the workbook
│       ├── references/               Guidance notes the skill reads (units, modules, testing, …)
│       ├── assets/                   Report and test-scenario templates
│       └── evals/                    Test prompts for evaluating the skill
├── build_skill.py                    Rebuilds the package in dist/
└── dist/
    └── becc-mecc-calculator-audit.skill   Packaged skill, ready to install
```

Local-only folders (excluded by `.gitignore`, never committed):

| Folder | Contents |
|---|---|
| `input/` | Workbooks being audited |
| `output/audit_evidence/` | Raw evidence from the scripts |
| `output/deliverables/` | Final report and issue list |

## Ways to use it

| Option | Best for | Needs |
|---|---|---|
| [1. Claude.ai or Claude desktop app](#1-claudeai-or-claude-desktop-app) | Non-technical users | A Claude account with Skills enabled |
| [2. Claude Code, all projects](#2-claude-code-available-in-all-projects) | Regular use on your computer | Claude Code, git |
| [3. Claude Code, one project only](#3-claude-code-one-project-only) | Sharing with a team through a project repo | Claude Code, git |
| [4. Run the scripts without Claude](#4-run-the-scripts-without-claude) | Quick automated scan, CI, scripting | Python 3.9+ |

> In options 2 and 3, copy the `skills/becc-mecc-calculator-audit/` folder, the one with
> `SKILL.md` directly inside it. Don't copy the whole repo, or the skill ends up too deep and
> Claude won't find it.

### 1. Claude.ai or Claude desktop app

1. **[Download becc-mecc-calculator-audit.skill](https://github.com/RaviLalitha98/becc-mecc-calculator-audit/raw/main/dist/becc-mecc-calculator-audit.skill)**.
   There's no need to clone. If the link doesn't download, open
   [`dist/becc-mecc-calculator-audit.skill`](dist/becc-mecc-calculator-audit.skill) on GitHub and click **Download raw file**.
2. In Claude, go to **Settings → Capabilities → Skills** and upload the file.
3. Start a new chat, attach the calculator workbook (`.xlsx`) and ask Claude to audit it.

### 2. Claude Code, available in all projects

Clone the repo and copy the skill into your personal skills folder.

**Windows (PowerShell)**
```powershell
git clone https://github.com/RaviLalitha98/becc-mecc-calculator-audit.git
New-Item -ItemType Directory -Force "$HOME\.claude\skills" | Out-Null
Copy-Item -Recurse becc-mecc-calculator-audit\skills\becc-mecc-calculator-audit "$HOME\.claude\skills\"
```

**macOS / Linux**
```bash
git clone https://github.com/RaviLalitha98/becc-mecc-calculator-audit.git
mkdir -p ~/.claude/skills
cp -r becc-mecc-calculator-audit/skills/becc-mecc-calculator-audit ~/.claude/skills/
```

To update later, run `git pull` in the cloned repo and copy the folder again.

### 3. Claude Code, one project only

Copy the skill into the `.claude/skills/` folder of the project where you keep your calculator
workbooks. Anyone who clones that project and opens it in Claude Code gets the skill
automatically.

**Windows (PowerShell)**, run from your project folder:
```powershell
git clone https://github.com/RaviLalitha98/becc-mecc-calculator-audit.git $env:TEMP\becc-mecc
New-Item -ItemType Directory -Force .claude\skills | Out-Null
Copy-Item -Recurse $env:TEMP\becc-mecc\skills\becc-mecc-calculator-audit .claude\skills\
```

**macOS / Linux**, run from your project folder:
```bash
git clone https://github.com/RaviLalitha98/becc-mecc-calculator-audit.git /tmp/becc-mecc
mkdir -p .claude/skills
cp -r /tmp/becc-mecc/skills/becc-mecc-calculator-audit .claude/skills/
```

Commit `.claude/skills/` to that project's repo to share the skill with your team.

### Using the skill in Claude Code (options 2 and 3)

Start Claude Code in a folder that contains the workbook, then do either of these:

- Ask in plain words: *"Audit BECC_v7_2026_V1.xlsx"* or *"Check this carbon calculator for
  errors"*. Claude loads the skill automatically.
- Call it directly: `/becc-mecc-calculator-audit BECC_v7_2026_V1.xlsx`

To check it's installed, type `/` and look for `becc-mecc-calculator-audit` in the list.

### 4. Run the scripts without Claude

Requirements: Python 3.9+ and `openpyxl`. Optional: `msoffcrypto-tool` for password-to-open
workbooks, LibreOffice (with UNO) for `recalc.py` scenario testing, and `python-docx`,
`markdown` and `xhtml2pdf` for exporting the report to Word or PDF.

```bash
pip install openpyxl msoffcrypto-tool
python skills/becc-mecc-calculator-audit/scripts/run_all.py input/WORKBOOK.xlsx --out output/audit_evidence \
    --twin "Super structure=Sub structure"
```

Add `--previous input/OLD_VERSION.xlsx` to compare against an earlier version. The scripts read
the workbook only and never modify it. Their output is *candidate* evidence: each item still has
to be traced and confirmed before it becomes a finding (see `SKILL.md`).

To turn a finished Markdown report into Word and/or PDF:
```bash
pip install python-docx markdown xhtml2pdf
python skills/becc-mecc-calculator-audit/scripts/export_report.py audit_report.md --format docx pdf
```

## Editing the skill

Anyone can refine the skill: improve the audit method, add a check or fix a script.

1. **Edit the source files** in `skills/becc-mecc-calculator-audit/`:

   | To change… | Edit |
   |---|---|
   | The audit steps Claude follows | `SKILL.md` |
   | Guidance on units, life-cycle modules, severity, etc. | `references/*.md` |
   | The automated checks | `scripts/*.py` |
   | The report layout or test scenarios | `assets/` |

   Don't edit `dist/becc-mecc-calculator-audit.skill` directly. It's a generated zip file.

2. **Test script changes** (if you changed any) on a sample workbook:
   ```bash
   python skills/becc-mecc-calculator-audit/scripts/run_all.py input/WORKBOOK.xlsx --out output/test
   ```

3. **Rebuild the package** so the downloadable file includes your changes:
   ```bash
   python build_skill.py
   ```
   This works on Windows, macOS and Linux. It leaves out `evals/` and Python cache files.

4. **Commit and push** both the source changes and the rebuilt package:
   ```bash
   git add -A
   git commit -m "Describe what you changed"
   git push
   ```

If you skip step 3, people who download the `.skill` file get the old version. Claude Code
users who copy the `skills/` folder get your changes either way.
