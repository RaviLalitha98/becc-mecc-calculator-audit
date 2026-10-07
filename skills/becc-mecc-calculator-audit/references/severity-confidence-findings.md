# Classification, severity, confidence and the findings record (stages 17, 18, 19)

## Class (what kind of problem)

| Class | Use when |
|---|---|
| Confirmed error | evidence shows the cell computes something other than what its label, header or unit says, ideally demonstrated numerically |
| Probable error requiring review | strong evidence (isolated pattern break pointing at a differently-labelled row) but intent not provable from the file |
| Unusual but potentially intentional logic | deviates from pattern but plausibly deliberate (special case, override, boundary row) |
| Maintainability risk | correct today, likely to break with growth/edits (fixed ranges, hand-picked SUMs, hard-coded indexes) |
| Version-control / workbook-evolution risk | differences between versions or twins, dead links, legacy sheets, inconsistent protection, save damage |

## Severity (impact if the finding is real)

- **Critical** — likely to materially change final BECC/MECC results: wrong EF or material
  mapping in a major material; wrong life-cycle link; major double counting; a major material
  category or module excluded from totals; broken formula feeding headline totals.
- **High** — significant calculation issue affecting a subset of results (one material, one
  module, one building part, one normalised metric).
- **Medium** — potentially incorrect or inconsistent calculation needing investigation;
  silent-zero behaviours; limited-scope range truncation.
- **Low** — minor inconsistency, maintainability or presentation issue.
- **Observation** — not necessarily wrong; document or review.

Assign severity from *what the cell feeds* (`trace.py dependents`) and how likely the triggering
condition is in real use. A defect in a rarely used row is lower than one in every project's
concrete line. Never from syntax alone.

## Confidence (how sure)

Confirmed (demonstrated by recalculation or unambiguous evidence) · High · Medium · Low.

## findings.json schema

```json
[
  {
    "id": "F-001",
    "title": "A1–A3 module total excludes the last material block",
    "class": "Probable error requiring review",
    "severity": "Critical",
    "confidence": "High",
    "workbook": "Calculator_v3.xlsx",
    "sheet": "Materials",
    "range": "H200",
    "category": "Aggregation",
    "lifecycle_module": "A1-A3",
    "material": "Insulation block (rows 176–195)",
    "current_logic": "=SUM(H10:H175)",
    "expected_logic": "=SUM(H10:H195) — matching sibling module totals I200:M200",
    "linked_sources": "H176:H195 ← G176:G195 × F176:F195",
    "evidence": "5 of 6 sibling totals in row 200 span rows 10:195; H200 spans 10:175. Rows 176–195 hold the Insulation block (labels in A176:A195).",
    "issue": "Module total omits the last material block.",
    "impact": "Summary!C4 and all normalised A1–A3 intensities understate whenever insulation is entered; the per-material summary still includes it, so the two summaries disagree.",
    "recommendation": "Extend the range to the block end and add a reconciliation check cell.",
    "verification": "recalc scenario T2b: 1,000 kg insulation in row 180 → I200 changes, H200 does not.",
    "proposed_formula": "=SUM(H10:H195)",
    "status": "open"
  }
]
```

Required: id, title, class, severity, confidence, sheet, range, category, current_logic, issue,
impact, recommendation. `findings_register.py` validates enumerations and produces
`findings_register.xlsx` (formulas stored as text), `findings_register.md` and severity × class
counts for the executive summary.

Preferred categories: Structure, Formula pattern, Hard-coded value, Reference/linkage,
Error/masking, Lookup/mapping, EF database, Life-cycle module, Units, Aggregation, Data
validation, Hidden logic, Protection, Version robustness, Excel compatibility, External link.

## Writing a finding

Lead with the consequence, then the evidence. Show current vs expected formula. Name the labels
of the cells involved so a reader can see *why* it is wrong without opening Excel. Give the
reproduction path (sheet → cell → what to compare). Keep speculation out of "evidence" — put it
in "issue" with a confidence that reflects it.

Recommended-correction block in the report:

**Current formula** `…`
**Proposed formula** `…`
**Reason** — exactly why the change is required and what else must change with it (sibling
cells, twin sheet, check cells).
