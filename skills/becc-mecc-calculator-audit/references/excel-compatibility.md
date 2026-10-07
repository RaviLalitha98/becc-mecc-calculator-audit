# Excel version compatibility (stages 24, 27, 29)

Start from `inventory.md` → "Excel version compatibility (functions actually used)",
`_xlfn`/`_xlws` counts, calcPr, date system, file format, VBA/Power Query flags, and the
DYNAMIC_ARRAY / LEGACY_ARRAY / IMPLICIT_ISECT groups in `formula_scan.md`. Generate the matrix
from what is detected; never paste a generic table.

If supported Excel versions are not documented, say so and list it as a requirement to confirm.

## Function availability (first perpetual release; Microsoft 365 usually earlier)

| Introduced | Functions |
|---|---|
| 2010 | AGGREGATE, NETWORKDAYS.INTL |
| 2013 | IFNA, XOR, FORMULATEXT, SHEET(S), DAYS, NUMBERVALUE, CEILING/FLOOR.MATH |
| 2016 | FORECAST.ETS family |
| 2019 | CONCAT, TEXTJOIN, IFS, SWITCH, MAXIFS, MINIFS |
| 2021 | XLOOKUP, XMATCH, FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, RANDARRAY, LET; dynamic arrays |
| 2024 | LAMBDA & helpers, TEXTSPLIT/BEFORE/AFTER, VSTACK/HSTACK, TAKE/DROP, CHOOSECOLS/ROWS, TOCOL/TOROW, WRAPCOLS/ROWS, EXPAND, IMAGE |
| M365 only | GROUPBY, PIVOTBY, PERCENTOF, REGEX*, TRIMRANGE, TRANSLATE, PY, COPILOT |

Confirm edge cases against Microsoft's function reference before asserting them. Unsupported
functions return `#NAME?` in older versions; the file keeps the `_xlfn.` prefix and shows the
last cached value until the cell is recalculated or edited.

## What to check

- **Prefixes**: `_xlfn.X` means X is newer than the Excel 2007 function set — normal in modern
  files; it tells you the minimum version, not that something is broken. `_xlfn._xlws.` marks
  worksheet-only functions (FILTER, SORT). `_xlfn.SINGLE` = the `@` implicit-intersection
  operator; `_xlfn.ANCHORARRAY` = `A1#` spill references.
- **Dynamic arrays** (cells with the `cm` attribute): in Excel 2019 and earlier they behave as
  legacy arrays and do not spill; dependent dropdown lists built on spills (UNIQUE → list
  validation) truncate or break. Check `#SPILL!` risk (anything typed into the spill area).
- **Implicit intersection** (IMPLICIT_ISECT): a multi-cell range where a scalar is expected.
  Pre-dynamic Excel takes the same-row/column value; dynamic Excel shows `@` when opening old
  files — same result. A formula *re-entered* in dynamic Excel without `@` spills or errors.
- **Legacy CSE arrays** (LEGACY_ARRAY): verify the whole array block is intact (`array_ref`),
  no single cell inside was overwritten, and results match a dynamic-array rewrite.
- **Lookups**: VLOOKUP/HLOOKUP default to approximate; XLOOKUP defaults to exact; MATCH
  defaults to 1. Rewrites between them change semantics unless arguments are explicit.
- **Calculation settings** (`calcPr`): manual mode → stale values (`CACHED_MISMATCH`);
  `autoNoTable`; iteration on → circular references hidden, results depend on count/delta;
  `fullPrecision="0"` (precision as displayed) permanently rounds stored values.
- **Date system**: 1904 vs 1900 — mixing workbooks shifts dates by 1,462 days.
- **Rounding**: where ROUND/INT/TRUNC/CEILING/FLOOR occur (row, stage, subtotal or final);
  whether versions round at different stages; `=A1=B1` on calculated floats (FLOAT_EQUALITY) —
  suggest `ABS(A1-B1)<tol`.
- **Names**: scope, hidden names, OFFSET/INDIRECT dynamic names, names pointing to `#REF!` or
  external books.
- **Tables**: structured references, calculated-column consistency, totals rows; compare with
  fixed ranges for robustness.
- **External links**: absolute local paths (`/Users/...`, `C:\...`), network/SharePoint/OneDrive
  paths, whether any formula or name still uses them, cached vs live values, behaviour when
  missing (prompt, stale values, `#REF!` on edit).
- **Power Query / connections**: Get & Transform availability, connectors, credentials,
  refresh-on-open, cached data mistaken for current.
- **VBA**: `Declare` statements need `PtrSafe` (and `LongPtr`) on 64-bit Office; missing
  references; ActiveX (Windows only); version-specific object-model calls; hard-coded workbook
  and sheet names and file paths. Excel for the web does not run VBA; Mac has gaps.
- **File format / compatibility mode**: `.xls` (65,536 rows, no newer functions), `.xlsx`
  (no macros), `.xlsm`, `.xlsb`. List what would be lost if saved down.
- **Platform**: Excel for Mac / web / mobile differences where relevant.

## Matrix (build from detected features)

| Feature (detected) | Where | 2016 | 2019 | 2021 | 2024 | M365 | Risk |
|---|---|---|---|---|---|---|---|
| e.g. CONCAT ×5,190 | every EF lookup key | #NAME? | ✓ | ✓ | ✓ | ✓ | High if 2016 must be supported |

Risk = likelihood the supported user base hits it × what it breaks.

## Final compatibility conclusion (stage 29)

State explicitly:
- which Excel versions were *evaluated* and how (static inspection; recalculation in
  LibreOffice x.y; Excel itself only if actually available) — never claim Excel testing that did
  not happen;
- the minimum Excel version implied by the detected features and which versions are expected to work;
- which versions contain unsupported formulas, and in which cells/sheets;
- behaviours that may differ between versions (dynamic arrays, implicit intersection, dependent
  dropdowns, calc mode);
- platform caveats (Mac/web) if VBA, ActiveX or Power Query are present;
- the recommended minimum supported version and what would need to change to support older ones.
