---
name: becc-mecc-calculator-audit
description: Forensic technical audit of building/material embodied-carbon (BECC/MECC) calculators built in Excel or similar spreadsheets — formula integrity, cell linkage, material and emission-factor mapping, life-cycle modules (A1–A3, A4, A5, B, C1–C4, D), units, totals, hidden logic, protection, Excel-version compatibility and version-to-version regression, ending in an evidence-backed findings register and evaluation report. Use this whenever someone wants an embodied-carbon, whole-life-carbon, LCA, EPD-based or Green Mark carbon calculator workbook checked, reviewed, validated, QA'd, stress-tested, compared with a previous version, or made robust — even if they only say "can you check this carbon calculator", "is this spreadsheet right", "why did the total change between v6 and v7", or attach an .xlsx with emission factors and A1–A3 columns.
---

# BECC / MECC calculator audit

The aim is not to confirm that the workbook produces a number. It is to establish, with
reproducible evidence, whether the calculation architecture — formulas, references,
material mappings, emission-factor lookups, life-cycle boundaries, units, totals and
version evolution — can be relied on, and exactly where it cannot.

Two habits make or break this kind of audit:

1. **Automate the search, not the verdict.** The bundled scripts scan tens of thousands of
   formulas in seconds and surface *candidates*. Every candidate is then opened, traced and
   reasoned about before it becomes a finding. A pattern break is often a deliberate special
   case; a "clean" pattern can be consistently wrong. Neighbouring formulas are not assumed
   correct.
2. **Separate kinds of problem.** A spreadsheet defect (wrong cell referenced) and a
   maintainability risk (fixed `SUM(H10:H150)`) need different owners and different fixes.
   Mixing them makes the report unusable.
3. **Methodology is out of scope.** The audit checks whether the workbook calculates what it
   sets out to calculate — judged against its own labels, headers, units, notes and
   consistency — not whether the carbon-accounting approach or the emission-factor data are
   right. Do not review, question or comment on methodology (system boundary, Module D,
   biogenic carbon, wastage rates, transport assumptions, factor sources and so on), do not
   ask the user methodology questions, and do not raise them as findings.

## Ground rules

- Never modify the original. Record its SHA-256 at intake and again at the end. Experiments
  (recalculation, inserted rows, test inputs) happen on copies. Do not save a workbook
  through openpyxl and hand it back — openpyxl silently drops x14 data validations,
  dynamic-array metadata, some charts and protection details.
- Passwords: use only a password the authorised owner/reviewer supplied, only to read what
  the audit needs, on a copy. Never attempt to crack, guess or strip protection. Sheet and
  workbook-structure protection do not stop openpyxl reading formulas; password-to-open
  encryption needs the password (`--password`, uses `msoffcrypto-tool`). Record what was
  protected and what could not be inspected (see `references/structure-hidden-protection.md`).
- Every finding carries evidence another reviewer can reproduce: workbook, sheet, cell or
  range, current formula/value, linked sources, expected pattern, why, impact, confidence.
- Say what is unknown. Never state "the calculator is correct".

## Tools

Scripts live in `scripts/` (Python 3 + openpyxl; recalculation needs LibreOffice with UNO).
Run from that directory or with full paths. All write Markdown summaries plus CSV detail.

| Script | Purpose |
|---|---|
| `run_all.py WB --out DIR [--twin "A=B"] [--previous OLD.xlsx]` | Full automated pass + `00_INDEX.md` |
| `inventory.py` | Sheets (visible/hidden/veryHidden), protection, used ranges, hidden rows/cols, data validation (incl. x14), names, external links, VBA/Power Query/connections, calc settings, date system, function usage, Excel-version compatibility matrix, sheet dependency edges |
| `formula_patterns.py` | R1C1 pattern-break detection (off-by-one, wrong column/sheet, swapped pairs, hard-coded values among formulas) with the formula the pattern predicts |
| `formula_scan.py` | Errors, #REF!, external refs, IFERROR/IFNA masking, magic numbers, implicit intersection, legacy/dynamic arrays, float equality, rounding, blank precedents, text in arithmetic |
| `lookup_audit.py` | Every lookup: match mode, returned column header, truncated/inconsistent ranges, duplicate keys, fallback chains; material→table mapping matrix; emission-factor table profiles |
| `aggregation_check.py` | Totals: rows excluded after range end, sibling totals with different extents, nested subtotals (double counting), hidden rows, cached-total mismatch, hand-picked multi-area SUMs |
| `trace.py WB precedents|dependents REF` | Upstream tree to leaf inputs / downstream users |
| `compare_versions.py OLD NEW` | Layout-aligned diff: formula/reference/function changes, formulas overwritten by values, data (EF) changes, names, settings; also twin-sheet comparison within one workbook via `--pair` |
| `recalc.py --scenarios S.json --workbook V1 [--workbook V2]` | Recalculate copies with test inputs in headless LibreOffice, check expectations, reconcile versions, check entries against data-validation rules |
| `findings_register.py findings.json --out DIR` | Validate findings and build the register (xlsx + md) and severity counts |
| `export_report.py REPORT.md --format docx pdf` | Convert the Markdown report to Word and/or PDF (A4 landscape; needs `python-docx` / `markdown` + `xhtml2pdf`) |
| `dump_cells.py` | Every cell to CSV for ad-hoc pandas/grep work |

## Workflow

Create a task list from these phases. Phases 3–12 can be reordered to follow the evidence,
but none should be skipped silently — if one does not apply (no Module D, single version),
say so in the report.

### Phase 0 — Intake
Collect: workbook(s) and which is current; any user guide or version notes for the
workbook; supported Excel versions; authorised passwords; what "BECC" vs "MECC" means for this owner; any known
problem prompting the audit. Missing items become explicit assumptions, not blockers.
Hash the original(s): `sha256sum file.xlsx`.

### Phase 1 — Automated evidence pass
`python scripts/run_all.py WB --out audit_out [--twin "Super structure=Sub structure"] [--previous OLD]`

Add `--twin` for sheets that should mirror each other (super/sub-structure, building A/B,
scenario 1/2) — differences between twins are one of the richest sources of real defects.
Read `audit_out/00_INDEX.md`, then each `.md`. Expect roughly 30–60 s for a 40k-formula book.

### Phase 2 — Understand the workbook before judging it
Read `references/structure-hidden-protection.md`. From `inventory.md`, classify every sheet
(input, calculation, EF database, lookup/dropdown, summary/dashboard, config, version/history,
documentation, orphan). Reconstruct the calculation flow
**Inputs → quantities → classification → emission factors → module calculations →
aggregation → results** and write it down with the sheets and key columns at each step
(a small diagram helps the reader). Open representative rows of each calculation block and
note the intended formula architecture per material (quantity × conversion × factor, etc.).
Flag unclear paths, duplicated calculation blocks, unused/orphan sheets, hidden calculation
logic and hard-coded outputs.

### Phase 3 — Formula integrity and hard-codes
Read `references/formula-integrity.md`. Work through `pattern_anomalies.md` (likely-defect
first) and `formula_scan.md`. For each candidate: open the cell and its neighbours, read the
row and column labels, trace the reference to see *what* it points at, and decide —
defect / intentional special case / needs owner confirmation. Pattern breaks that only
differ in `$` anchoring are robustness issues, not wrong results.

### Phase 4 — Linkage and dependency tracing
Trace every headline output (each module total, each material total, normalised
kgCO2e/m²) with `trace.py precedents`. Confirm each path ends at the right inputs, the right
EF table and the right column. Look specifically for cross-material, cross-module,
cross-component and cross-scenario links (aluminium reading steel factors, A4 reading
A1–A3, sub-structure reading super-structure inputs).

### Phase 5 — Material mapping and emission-factor database
Read `references/materials-and-ef-database.md`. Use `lookup_audit.md`: approximate matches,
duplicate keys, truncated ranges, returned-column headers, fallback-to-generic chains,
mapping matrix, EF table profiles (source, year, version, geography, unit, boundary,
negatives, zeros, outliers). An EF is not "wrong" without evidence — classify as
"requires source verification" instead.

### Phase 6 — Life-cycle modules
Read `references/lifecycle-modules.md`. Audit each module present (A1–A3, A4, A5, B1–B7,
C1–C4, D) for linkage, quantity, factor, units and aggregation, and test the double-counting
traps listed there.

### Phase 7 — Units
Read `references/units.md`. Write the unit at every step of each module chain and check every
conversion and every magic number from `formula_scan.md` (1000, 0.001, densities, 44/12…).

### Phase 8 — Totals
Use `aggregation_check.md`. Verify subtotals reconcile to totals (recompute where useful),
check excluded rows, sibling extent mismatches, nested subtotals and hand-picked SUM lists.

### Phase 9 — Version and twin comparison
Read `references/version-robustness.md`. Classify each difference (documented intended
change, expected data update, implementation change, compatibility difference, probable
regression, requires investigation). Compare semantics, not text. Rate formula design
robust / moderately fragile / highly fragile with reasons. Check which application saved
each version: a file written by openpyxl or another library loses cross-sheet dropdowns,
charts, spill metadata and cached results — report that save damage separately from
content changes. Make sure every changed cell is exercised by at least one scenario, and
test changes individually as well as together (errors can partly cancel).

### Phase 10 — Excel compatibility
Read `references/excel-compatibility.md`. Build the compatibility matrix from detected
features only. Cover `_xlfn`/`_xlws` prefixes, dynamic arrays and `@`, legacy CSE arrays,
lookup semantics, calc mode/iteration/precision, date system, rounding, names, tables,
external links, Power Query, VBA (32/64-bit, `PtrSafe`), file format.

### Phase 11 — Boundary and scenario tests
Read `references/testing.md`. Build a scenario JSON (start from `assets/scenarios_template.json`)
covering the ten standard tests that the workbook's structure allows, run `recalc.py`, and
use the results to confirm or refute static findings. A demonstrated wrong number turns
"High confidence" into "Confirmed". When two versions exist, run the same scenarios through
both and reconcile every difference back to its cause.

### Phase 12 — Data validation, hidden logic, protection
Inventory data-validation rules and test whether invalid entries (negative quantities,
wrong units, blank selections) are blocked or silently produce zeros. Review hidden
rows/columns/sheets, names, helper cells, macros, queries and protection effects on
auditability.

### Phase 13 — (not performed) Methodology
Methodology and emission-factor data are out of scope (see principle 3). Skip this phase;
do not list methodology questions.

### Phase 14 — Findings, verification and report
Read `references/severity-confidence-findings.md`. Record findings in `findings.json`
(schema in that file) and run `python scripts/findings_register.py findings.json --out audit_out`
to validate fields and produce the register and counts. Then write the report from
`assets/report_template.md` as `audit_report.md`, and deliver it in the format(s) the user
chose (see Deliverables).

Before finalising, re-verify every Critical and High finding by re-opening the cells in the
original file (not your notes) — or, when subagents are available, give the finding list and
workbook to a fresh agent and ask it to try to refute each one. Downgrade anything that does
not survive. Re-hash the original to show it is unchanged.

## Classification, severity, confidence (summary)

Each finding gets one **class**: Confirmed error · Probable error requiring review ·
Unusual but potentially intentional logic · Maintainability
risk · Version-control/workbook-evolution risk.

**Severity** (impact if real): Critical (materially changes BECC/MECC results) · High
(significant, subset of results) · Medium (potentially incorrect, needs investigation) ·
Low (minor/maintainability/presentation) · Observation.

**Confidence** (how sure): Confirmed · High · Medium · Low.

Severity and confidence are independent: an unverified swapped reference in the headline
total can be Critical/Medium. Never assign severity from unusual syntax alone — assign it
from what the cell feeds. Full definitions and examples: `references/severity-confidence-findings.md`.

## Deliverables

1. **Evaluation report** following `assets/report_template.md` (executive summary, findings
   table, formula consistency, cell linkage, life-cycle modules, materials, units, version
   robustness, compatibility matrix and conclusion, recommended
   corrections with current/proposed formula and reason, test results, limitations).
   Always write it first as Markdown (`audit_report.md`), then deliver it in the format the
   user wants. If they have not said, ask once, before writing the report, offering:
   - **In the chat**: show the report as Markdown directly in the conversation (good for a
     quick read; long reports may be summarised in chat with the full file attached).
   - **Word (.docx)**: editable, for comments and tracked changes.
   - **PDF**: fixed layout, for sharing and sign-off.
   They can pick more than one. Produce Word/PDF with
   `python scripts/export_report.py audit_report.md --format docx pdf --out DIR`; if the host
   has its own Word or PDF skill, that may be used instead. Open or render the result and
   check the tables are readable before handing it over.
2. **Findings register** (`findings_register.xlsx`) — one row per finding, filterable.
3. **Evidence folder** — the `audit_out` outputs, scenario JSON and recalculation results,
   so another reviewer can reproduce every finding.

Proposed corrections are recommendations. Do not apply them to the workbook unless
explicitly asked; if asked, apply to a copy, re-run the scripts and scenarios on the copy,
and report what changed.

## Scaling notes

- Group by R1C1 signature. 18,000 formulas copied down 600 rows are ~50 distinct formulas;
  review the distinct ones, then the breaks.
- `pattern_map.csv` lists every distinct formula per column with its row runs — the fastest
  way to learn a calculation sheet's architecture.
- Empty templates have blank inputs, so BLANK_PRECEDENT and #DIV/0! noise is expected;
  populate a scenario with `recalc.py` to see real behaviour.
- LibreOffice is a stand-in for Excel. Functions it lacks (reported in recalc caveats) will
  error in LibreOffice but may work in Excel — never report those as calculator defects.
