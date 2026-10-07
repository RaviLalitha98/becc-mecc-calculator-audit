# Boundary, validation and version-regression testing (stages 13, 14, 28)

Static inspection finds candidates; recalculation proves impact. `recalc.py` writes inputs into
a copy, recalculates in headless LibreOffice and reads outputs. Expectations can be
`{"value": x, "tol": t}`, `{"same_as": "Sheet!A1"}`, `{"equals_product_of": [refs]}` or
`{"is_error": true}`.

## Designing scenarios

1. Identify the input cells for one material row (material, country/variant, quantity, unit,
   user EF override, transport mode/distance) from the input legend and data validation.
2. Pick values from the dropdown sources (inventory shows each rule's formula1) — values not in
   the list would be rejected by Excel even though the API accepts them; recalc reports this per
   input. For dependent dropdowns, read the key column of the EF table to find valid combinations.
3. Compute the expected result by hand from the EF table and put it in the scenario.
4. Read outputs at every step of the chain (row → block sum → module total → summary table →
   normalised value), so a failure shows *where* the chain breaks.
5. Set the area inputs (GFA etc.) so normalised outputs are not `#DIV/0!`.

## Standard tests (use those the workbook's scope allows)

| # | Test | Shows |
|---|---|---|
| 1 | Single material, known A1–A3 factor | basic chain; summary pick-up |
| 2 | Several materials in one element | aggregation, per-material totals, no cross-talk |
| 3 | Transport (known mass, distance, mode) | A4 units (t·km vs kg·km), mode factor lookup |
| 4 | Construction waste % / site energy | A5 treatment, no double counting with quantity |
| 5 | End-of-life C1–C4 | shares sum to 100%, correct factors |
| 6 | Module D | credit sign, separate reporting vs netting |
| 7 | Missing EF (material/country with no record) | warning vs silent zero vs fallback to generic |
| 8 | Zero quantity, zero distance, zero waste; invalid unit text; negative quantity | zeros propagate cleanly; invalid entries blocked or flagged |
| 9 | Material addition (row inserted inside/after a block, on a copy, in a spreadsheet app) | totals and lookups expand or miss the row |
| 10 | Version regression — same inputs through V1 and V2 | reconcile every module and material difference |

Worth adding: a quantity in the **last row of each block** and in rows near the end of the
sheet (catches totals with short ranges); user EF override of 0; the sub-structure equivalent of
each super-structure test; duplicated material names; recycling rate 0% and 100%.

For each, record whether the calculator gives a meaningful warning, returns zero incorrectly,
errors, silently substitutes a value, or misleads.

## Data validation review (stage 14)

From `inventory.json` (`sheets[].validations`): list each input region with its rule type,
bounds, list source, allowBlank and error style (only *stop* blocks entry; validation never
blocks paste). Then answer: can a user enter a physically invalid value
(negative mass, unit not in the conversion logic, waste > 100%, text in a number field) without
any warning, and what does the calculator do with it?

## Reporting test results

Table per scenario: inputs, expected, actual, PASS/FAIL, and which static finding it confirms
or refutes. Note the engine (LibreOffice version) and caveats; a LibreOffice-only failure in a
function LibreOffice lacks is not a calculator defect.
