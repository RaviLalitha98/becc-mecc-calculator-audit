# [Calculator name] — Technical Audit and Evaluation Report

- **Workbook(s) reviewed:** [file name(s), version, SHA-256 at intake]
- **Compared against:** [previous version / twin sheets / none]
- **Methodology reference:** [document + version, or "not provided — methodology verification requires confirmation"]
- **Supported Excel versions (stated):** [list, or "not documented — to confirm"]
- **Review date / reviewer:** [ ]
- **Scope and limits:** [what was and was not reviewed: protected/encrypted parts, VBA, Power Query, Excel vs LibreOffice recalculation, inputs used]

---

## 1. Executive summary

- **What the calculator is:** [purpose, BECC/MECC scope, modules covered, materials covered, sheet count, formula count]
- **Overall structure:** [one-paragraph calculation flow]
- **Issues by severity:** Critical [n] · High [n] · Medium [n] · Low [n] · Observation [n] (by class: confirmed errors [n], probable errors [n], intentional-but-unusual [n], methodology [n], maintainability [n], version/evolution [n])
- **Major calculation risks:** [top 3–5, one line each with finding IDs]
- **Major methodology risks:** [top items with question IDs]
- **Version-robustness concerns:** [fragility rating and the main break scenarios]
- **Excel compatibility:** [minimum version implied; main incompatibilities]
- **Overall assessment:** [evidence-based statement of reliability and the conditions attached — never "the calculator is correct"]

## 2. Workbook structure and calculation flow

[Sheet classification table: sheet · class · visibility · protection · role]
[Calculation flow chain per module]
[Orphan/legacy sheets, hidden logic, external links, names]

## 3. Detailed findings table

| ID | Severity | Confidence | Sheet | Cell/Range | Category | Current Logic | Expected/Reference Logic | Issue | Potential Impact | Recommendation |
|---|---|---|---|---|---|---|---|---|---|---|

(From `findings_register.md`; full register in `findings_register.xlsx`.)

## 4. Formula consistency findings

[Pattern anomalies reviewed: n candidates → n defects, n intentional, n robustness. Each defect: current vs predicted formula, labels, why.]
[Hard-coded values in calculated regions; masked errors; magic numbers]

## 5. Cell-linkage findings

[Suspicious cross-sheet, cross-material, cross-module, cross-component references; trace summaries of headline outputs]

## 6. Life-cycle module findings

### A1–A3
### A4
### A5
### B modules
### C1–C4
### D
[For modules not modelled, say so.]

## 7. Material-specific findings

[By material: mapping, lookup, EF table, formula architecture consistency]

## 8. Emission-factor database review

[Table profiles; duplicates; units; geography; source/year/version; boundaries; negatives/zeros/outliers — "requires source verification" where unproven]

## 9. Unit and conversion findings

[Each suspected unit error: formula — units in → units out — expected — effect]

## 10. Aggregation and totals

[Excluded rows, sibling mismatches, nested subtotals, hand-picked SUMs, reconciliation results]

## 11. Version robustness findings

[What breaks when materials/rows/modules/columns/sheets are added or renamed; fragility per area (robust / moderately fragile / highly fragile) with reasons]
[Cross-version / twin comparison: each difference classified as expected methodology change / expected data update / implementation change / compatibility difference / probable regression / requires investigation]

## 12. Excel version compatibility

| Feature (detected) | Where | 2016 | 2019 | 2021 | 2024 | M365 | Risk |
|---|---|---|---|---|---|---|---|

**Compatibility conclusion:** versions evaluated and how; versions expected to work; versions with unsupported formulas (cells); behaviour differences; recommended minimum version.

## 13. Data validation, hidden logic and protection

[Input controls and gaps; hidden content; protected components reviewed / not reviewable; auditability risks]

## 14. Test results

| Test | Inputs | Expected | Actual | Result | Finding supported |
|---|---|---|---|---|---|

[Engine used and caveats]

## 15. Methodology questions (for the methodology owner)

| ID | Question | Why it matters | Where in workbook | Evidence/assumption observed |
|---|---|---|---|---|

(Not defects — decisions that need confirmation.)

## 16. Recommended corrections

### [Finding ID] — [title]
**Current formula** `…`
**Proposed formula** `…`
**Reason** …
**Also change** [siblings, twin sheet, check cells]

(Recommendations only; the workbook was not modified.)

## 17. Limitations and reproducibility

[What could not be verified; assumptions; evidence folder contents; how to reproduce; original file hash unchanged at end]

**Outside the scope of this evaluation (Excel side).** Besides methodology and emission-factor
data, which are not verified, the following were not fully covered. Keep every item; say which
ones apply to this workbook (e.g. "no VBA present") and add anything else not reviewed.

- **Live Excel behaviour:** the review reads formulas and the values Excel last saved; scenario
  tests were recalculated in [LibreOffice / Excel version]. Behaviour that appears only when
  Excel itself recalculates is not covered unless tested by hand in Excel.
- **Other Excel versions and platforms:** compatibility with Excel 2016/2019/2021 is assessed
  from the functions used, not by opening the file in each version. Excel for Mac, the web and
  mobile were not tested.
- **Macros, Power Query and data connections:** detected but not executed or fully reviewed;
  locked VBA projects and encrypted parts without a password cannot be inspected. [present / not present]
- **Run-time references** (INDIRECT, OFFSET, text-built references): only partly traceable statically.
- **External linked files:** links are reported, but the linked workbooks themselves are not reviewed.
- **Presentation and usability:** charts are checked only against the totals they display;
  conditional formatting, number formats, print layout, user experience and accessibility are not reviewed.
- **Performance and file health:** recalculation speed, file size and stability are not assessed.
- **Applying corrections:** fixes are recommended, not applied; their effect needs re-testing after implementation.
- **Untested inputs:** only the scenarios listed in section 14 were run; other combinations of
  materials, countries and entries may reveal further issues.
