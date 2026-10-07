# Formula integrity, hard-codes, linkage and per-material consistency (stages 2, 3, 4, 12, 17)

## How the pattern detector thinks

Each formula is converted to relative R1C1: `=B12*C12` in D12 becomes `=R[0]C[-2]*R[0]C[-1]`,
identical in every row it was copied to. `formula_patterns.py` compares each cell with its
neighbours (vertically and horizontally, ±6 cells, tolerating one blank row):

- **SANDWICH** — the nearest neighbours either side agree, are well supported, and this cell differs.
- **MAJORITY** — ≥75% of the window agrees and this cell differs.
- **HARDCODE_*** — a number sits where neighbours have formulas.
- **SWAPPED PAIR** — adjacent cells with opposite ±1 offsets (two references exchanged).

For each hit it reconstructs the formula the pattern predicts *for this cell* and explains
the difference: `row offset +1`, `column offset`, `different sheet`, `range end differs`,
`same cells, different $-anchoring`, or `structurally different`.

Triage hints: `likely-defect`, `review`, `intentional-or-defect` (structurally different —
often a deliberate first/last row or special material), `robustness` (anchoring only),
`review-hardcode`.

## Verifying a candidate (for every one you report)

1. Read the current formula, the predicted formula and both neighbours.
2. Read the **row label** and **column header** of the cell *and of the referenced cell*.
   An off-by-one is a defect when the referenced row belongs to a different material, stage
   or building element — e.g. a "Concrete+Steel+Glass" intensity row pointing at the
   "+Aluminium" total row.
3. Check whether the pattern itself is right: if 11 cells point at row r and one at r+1, ask
   which row the label says should be used. The majority can be wrong.
4. Trace one level up and down (`trace.py`) to see what the cell feeds — that sets severity.
5. Where possible show the numeric effect (`recalc.py`). That makes it Confirmed.
6. Record: current, expected, reason, labels, impact path.

Structurally different formulas at block boundaries (the first row of a block that also sums
the block, a total row) are usually intentional — say so briefly rather than listing them as
defects.

## Errors and references (stage 2)

- `#REF!` in formula text (ERROR_TEXT) — always a defect or dead formula; find what it fed.
- Cached errors (`#N/A`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#SPILL!`) — on an empty template many
  `#DIV/0!` are expected. Judge whether the error would surface to users with real inputs and
  whether downstream text, rankings and charts break (e.g. "top 3 materials" text built with
  VLOOKUP(LARGE(...)) over a column of errors).
- Circular references: `calcPr.iterate` in the inventory; SELF_REFERENCE in the aggregation check.
- Names to `#REF!` or to external books; references to missing sheets; external workbook paths.
- Wrong row / column / sheet / material references; mixed anchoring (`$A1` vs `A$1`)
  inconsistent across a block — matters when the block is copied or extended.

## Hard-coded values (stage 3)

Sources: HARDCODE_* in `pattern_anomalies.md`, HARDCODE and MAGIC_NUMBER in `formula_scan.md`.
Classify each as: legitimate input · intentional override (documented?) · temporary patch ·
suspicious replacement of a formula · undocumented assumption.

Watch hard-coded emission factors, distances, transport factors, wastage %, conversion
factors (1000, 0.001), service lives, densities, recycling/recovery rates, grid electricity
and fuel factors and their year. A grid factor in a labelled input cell with source and year
is fine; the same number embedded in 1,000 formulas is a traceability risk.

MASKED_ERROR matters: `IFERROR(lookup, 0)` or `IFNA(…, "")` turns a missing EF, a misspelt
material or an unmatched country into a silent zero — an *omission* that lowers embodied
carbon without warning. Record which lookups are masked and what the user sees.

## Linkage tracing (stage 4)

For each headline output run `trace.py WB precedents "'Sheet'!Cell" --depth 6`. Confirm:
the inputs that influence it; the material records and EF columns that contribute; which
life-cycle stages contribute (none missing, none twice); intermediate calculations in order.

Targeted wrong-source checks: aluminium rows referencing steel tables; concrete referencing a
generic factor; A4 referencing A1–A3 totals; A5 wastage on the wrong quantity; C-stage using
production factors; Module D using the wrong recycling factor; totals omitting or repeating
stages; one component/scenario/variant reading another's inputs; a summary column reading
the other building part (super- vs sub-structure). Twin-sheet comparison
(`compare_versions.py WB WB --pair "A=B"`) catches these between mirrored sheets fast.

## Per-material formula architecture (stage 12)

For each material category write the expected chain, e.g.
`qty (I) × unit factor (K: kg→1, tonne→1000) × EF (M) = A1–A3 (N)` and
`qty × K × distance × vehicle EF = A4`. Then check, using `pattern_map.csv`, that every block
for that material uses the same chain, and that comparable materials (precast vs in-situ
concrete, rebar vs section steel) treat density, waste, transport, recycling, stage factors
and unit conversions consistently. Differences need a documented reason.

## Evidence record for each finding (stage 17)

workbook · worksheet · cell/range · current formula/value · linked source cells ·
observed pattern · expected pattern (if inferable) · category · explanation ·
potential impact · confidence · recommended verification/correction.
