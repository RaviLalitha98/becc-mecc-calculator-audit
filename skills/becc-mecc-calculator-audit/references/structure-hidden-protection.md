# Workbook structure, hidden logic and protection (stages 1, 15, 23)

## 1. Sheet classification

Using `inventory.md` and a look at each sheet's first rows, put every sheet in one class:

| Class | Typical signals |
|---|---|
| Input | visible, data-validation dropdowns, "for user to input" legends, few formulas |
| Calculation | thousands of formulas, row blocks per material, hidden helper columns |
| Emission-factor / database | hidden, many numbers, key column (often CONCAT of country+material), source/unit columns |
| Lookup / dropdown | hidden, UNIQUE/FILTER or static lists referenced by data validation |
| Summary / dashboard / charts | references to calc totals, charts, normalised results |
| Configuration | grid factor, year, building type, scenario switches |
| Version / history / documentation | change log, landing page, declaration, guidance |
| Imported / external | external link caches, Power Query output |
| Orphan / legacy | no incoming references and not an input/output/doc (e.g. "Sheet (2)" copies, old databases) |

Inventory's "sheets with no incoming references" list mixes legitimate inputs/outputs with
orphans — decide which is which. An orphan EF table that looks authoritative is a risk:
someone may update it believing it is live.

## 2. Calculation flow

Write the flow as a short chain per major result, naming sheets and columns, e.g.:

```
Calc!I (qty) × K (unit factor) × M (EF: user override H, else lookup F from Data_<material>!E)
  → N (A1–A3 per row) → block subtotal → module total (row 600)
  → Results!C3 → Project summary B31 → normalised intensity B48 (÷ GFA)
```

Record per step: the sheet, the formula family (from `pattern_map.csv`), the units, and
where the chain can be bypassed (user overrides, IF branches, fallbacks, project-level switches).

Flag:
- unclear paths (a result that cannot be traced statically — INDIRECT/OFFSET, VBA)
- duplicated calculations (the same module computed in two places with different formulas —
  e.g. a per-material summary and a per-module summary that should reconcile but do not)
- disconnected blocks (formulas whose outputs nothing reads — `trace.py dependents`)
- hidden calculation logic (hidden columns doing real work: unit factors, flags)
- hard-coded outputs (a summary cell typed as a number)

## 3. Hidden logic checklist (stage 15)

- Hidden and veryHidden sheets — veryHidden can only be unhidden via VBA; note them explicitly.
- Hidden rows/columns inside calculation regions (inventory counts; check whether totals include them).
- Defined names: workbook vs sheet scope, hidden names, names to `#REF!`, names to external books, dynamic names (OFFSET/INDIRECT).
- Helper cells with flags or switches that change calculation logic (e.g. "if project-level A5 entered, ignore element A5").
- Data-validation sources — dropdown lists derived from formulas (UNIQUE/FILTER) behave differently in older Excel.
- Conditional formatting that hides values (white font, custom number formats `;;;`).
- Comments/notes holding assumptions.
- Macros (VBA): if `vba_project` is true, extract with `olevba` (`pip install oletools --break-system-packages`; `olevba file.xlsm`) when authorised. Look for code that writes values into calculation cells, Workbook_Open events, recalculation control, hard-coded paths and sheet names, Windows API `Declare` without `PtrSafe`.
- Power Query / connections: query names, sources, refresh-on-open, whether outputs are cached.
- External links: path, whether any formula or name still uses them, cached values. A link part with no formula users is dead weight that triggers "update links" prompts; a name pointing into it is a live dependency.
- Charts referencing ranges that do not match the totals tables.

## 4. Protection (stage 23)

Record per component: password-to-open (encryption), workbook structure protection, sheet
protection (and which actions are allowed), locked/unlocked cells, VBA project lock.

Procedure with an authorised password:
1. Work on a copy. `common.decrypt_if_needed` decrypts encrypted files to a temp copy.
2. Sheet/structure protection does not block reading formulas with openpyxl, so static review
   needs no unprotecting. Never remove protection from any file you return.
3. If Excel-side inspection requires unprotecting, do it on a copy and say so in the report.

Report:
- successfully reviewed protected content;
- protected content that could not be inspected (locked VBA project, encrypted parts without password);
- whether protection prevents access to formulas, names, links, macros or logic.

Check whether a `<sheetProtection>` element is actually enforcing (`sheet="1"`) — an element
with only option flags protects nothing. Protection is not an error. It becomes an
auditability/maintainability finding when critical formulas are inaccessible to reviewers;
a protected sheet hides undocumented assumptions; users can edit inputs but not inspect
dependent calculations; protection differs unexpectedly between versions; or unlocked cells
exist inside calculation regions (users can overwrite formulas despite "protection") — check
`cell.protection.locked` for formula cells on protected sheets.

Never attempt password cracking, hash attacks or protection-stripping tools.
