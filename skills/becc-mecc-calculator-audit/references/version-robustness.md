# Version robustness, cross-version consistency and semantic equivalence (stages 9, 10, 25, 26)

## Running the comparison

```
python scripts/compare_versions.py OLD.xlsx NEW.xlsx --out audit_out/version_diff
python scripts/compare_versions.py WB.xlsx WB.xlsx --pair "Super structure=Sub structure" --out audit_out/twin
```

The script aligns rows and columns (difflib over labels + R1C1 signatures), translates the old
formula's references into the new layout and reports only formulas that still differ. A formula
that moved from row 150 to 175 because 25 rows were inserted is *not* reported; one that now
reads a different logical cell is.

Change types: `FORMULA_TO_VALUE` (overwrite — top priority), `FUNCTION_CHANGE` (incl.
VLOOKUP→XLOOKUP), `REFERENCE_CHANGE`, `CLEARED`, `REMOVED`, `ADDED`, `DATA_CHANGE` (constants —
EF updates, with % change; float-representation noise is ignored), `VALUE_TO_FORMULA`,
`TYPE_CHANGE`, `LABEL_CHANGE`.

Workbook-level: the **saving application** (a NEW file written by openpyxl or another library
rather than Excel loses cross-sheet data validations, some charts, dynamic-array metadata and
cached results — report those as save damage, separately from content changes), sheets
added/removed/renamed, visibility, protection, validation counts, conditional formats, spill
cells, hidden rows/cols, charts, cached values stripped, calc settings, date system, names,
external links, newly used version-gated functions, VBA/Power Query.

Alignment is heuristic. When similarity is low, verify a few mappings by hand before trusting
REFERENCE_CHANGE lists. For twin sheets, many differences are legitimate (different material
blocks); look for the *isolated* ones (one subtotal range shorter than its twin).

## Three levels of comparison (stage 26)

1. **Syntax** — text differs? (`=B12*C12` vs `=[@Quantity]*[@EmissionFactor]`)
2. **Reference structure** — after alignment, same logical source cells / table / column?
3. **Calculation semantics** — same calculation meaning? This level matters most.

`VLOOKUP(A2,Factors!A:D,4,FALSE)` and `XLOOKUP(A2,Factors!A:A,Factors!D:D)` are equivalent only
if: both exact match; not-found handling equivalent (VLOOKUP → #N/A; XLOOKUP → #N/A unless
if_not_found is given — often 0/"" which masks gaps); duplicates resolve to the same row; and
column D is still the same field after any column insertion (VLOOKUP's hard-coded 4 does not
move; XLOOKUP's range does).

Check meaning, not names: if V1 is `Mass × A1-A3 EF` and V2 is `Mass × Carbon Factor`, open the
V2 "Carbon Factor" column and its source — has it become A1–A4 or A1–A5, a different statistic,
unit or geography?

## Classifying differences (stage 25)

| Class | Evidence |
|---|---|
| Documented intended change | documented in version notes; consistent across affected rows |
| Expected data update | DATA_CHANGE in EF tables with source/year updated |
| Formula implementation change | different formula, same semantics proven by tracing or scenario |
| Compatibility difference | result differs only because of Excel-version behaviour |
| Probable regression | unexplained change, inconsistent with notes, often isolated to one row/column/cell |
| Requires investigation | cannot yet be explained |

Trace every result difference backwards (scenario → module total → material block → row →
factor or formula) until one of these fits. Run the same scenarios through both versions
(`recalc.py --workbook OLD --workbook NEW`) and make sure every changed cell is exercised by at
least one scenario; errors can partly cancel, so check each change in isolation too.

An unsourced EF change is "Expected data update — requires source verification"; note
patterns such as the new value equalling another row's value (possible copy-paste).

## Fragility assessment (stage 9)

- **Robust** — structured tables or dynamic ranges; lookups by header; names that expand;
  totals built from tables; no hard-coded row lists.
- **Moderately fragile** — fixed ranges with generous headroom (`$D$2:$E$200` on a 60-row
  table); VLOOKUP with hard-coded column indexes; consistent but manual block structure.
- **Highly fragile** — fixed ranges that already stop at or before the last row; hand-picked
  multi-area SUM lists (`SUM(AF28:AF34,AF45:AF49,…)`); per-block totals with hard-coded heights;
  INDIRECT on sheet names; different extents for the same table in different formulas; module
  totals with different row extents.

Explain what breaks when: new materials are added; rows are inserted inside or *after* a block;
life-cycle modules are added (new columns shift VLOOKUP indexes); EF tables grow; sheets are
renamed (INDIRECT, external refs, VBA); columns move; a new database is introduced.

Simulation (optional, on a copy): insert a row inside a block and at its end in Excel or
LibreOffice (not openpyxl — it does not update references), then rerun `aggregation_check.py`
and `recalc.py` to show which totals miss the new row (Test 9).

## Fixed range vs structured table

`=SUM(H10:H150)` fails silently when rows are added below 150; `=SUM(Table1[Embodied Carbon])`
expands. Where the calculator uses fixed blocks, recommend tables or full-block totals plus a
check cell (`=SUM(all detail rows) - SUM(module totals)` should be 0) visible to users.
