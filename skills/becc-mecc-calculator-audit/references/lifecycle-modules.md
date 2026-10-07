# Life-cycle module audit and methodology review (stages 6, 16)

Module definitions follow EN 15978 / EN 15804 / ISO 21930 conventions. A calculator may cover
only some modules — report coverage explicitly ("A1–A5 only; B, C, D not modelled").

| Module | Typical formula | Quantity | Factor | Common faults |
|---|---|---|---|---|
| A1–A3 product | qty × unit conv. × EF | delivered or installed qty (state which) | EF A1–A3 per declared unit | wrong column returned; kg/t mismatch; EF already includes A4; separate A1, A2, A3 *plus* an A1–A3 total added together |
| A4 transport to site | mass × distance × mode factor (per t·km or kg·km) | mass incl. waste? | vehicle/ship factor | distance in miles/NM; per-t·km factor applied to kg; empty return double-applied or omitted; default marine distance by country applied silently |
| A5 construction | site energy/fuel × factor + waste (qty × waste% × (A1–A4 of wasted material + its disposal)) | wasted quantity | A1–A3 + C factors for waste | waste added to quantity *and* counted again in A5; project-level A5 replacing element A5 inconsistently between summaries |
| B1–B7 use | replacements = ceil(RSL_building / RSL_component) − 1; operational energy/water | | | replacement count off by one; B4 using A1–A3 without A4/A5/C; operational carbon mixed into embodied totals |
| C1 deconstruction | energy per m² or per t | | | missing; using A5 factor |
| C2 transport | mass × distance × factor | end-of-life mass | | reuse of A4 distance without justification |
| C3 waste processing | mass to recycling/recovery × factor | | | recycled share inconsistent with D |
| C4 disposal | mass to landfill × factor | | | landfill + recycling shares ≠ 100% |
| D beyond boundary | net flow × substitution factor | recovered − recycled input | credit factor | credit for recycled content already credited in A1–A3; D netted into the headline total when the methodology reports it separately |

## Checks for every module present

- Linked to the right quantity, the right factor and the right unit conversion.
- Applied only where appropriate.
- Aggregated once and only once into module totals and the grand total.
- Each module total is traceable to every material block: compare the set of rows feeding each
  module total (`aggregation_check` SIBLING_MISMATCH shows module totals with different row
  extents) and confirm summaries built different ways (per-material vs per-module) reconcile.

## Double-counting traps (test explicitly)

- EPD gives A1–A3 combined and the calculator also adds A1, A2, A3 separately.
- A4 embedded in a "delivered" factor and added again as transport.
- Recycled-content benefit in the production factor *and* again in Module D.
- Construction waste in A5 *and* in the C-stage quantity (or an inflated quantity *and* A5 waste).
- Project-level A5 entered while element-level A5 still sums into some totals.
- Summary totals built from both detail rows and subtotals (`NESTED_SUBTOTAL`).

## Methodology questions (keep separate from defects)

System boundary; Module D treatment; biogenic carbon (−1/+1, storage claims); carbonation of
concrete; recycled content and recycling credits (cut-off vs substitution); construction
wastage rates and their source; replacement cycles and reference service life; transport
assumptions and default distances; grid electricity factors and their year; allocation rules;
declared vs functional units; normalisation area (GFA vs CFA vs NLA vs sub-structure area);
reporting period; which statistic (median, mean, percentile) represents generic factors.

Where documentation exists, compare implementation with it and cite both. Where not, write:
"Methodology verification requires confirmation." Never silently impose an assumption — state
it and its effect.
