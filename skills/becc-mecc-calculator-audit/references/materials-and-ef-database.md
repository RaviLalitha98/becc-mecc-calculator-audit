# Material mapping and emission-factor database (stages 5, 11)

## Mapping checklist

For each material category the calculator supports (concrete and components, cement,
admixtures, aggregates, precast/blended concretes, rebar, structural/section steel,
stainless, aluminium, glass, timber, gypsum/plaster, insulation, plastics, bricks, tiles,
paint, waterproofing, sealants, carpet, copper, MEP items, "others"):

- [ ] The selection maps to exactly one EF record — keys unique (`DUPLICATE_KEYS`).
- [ ] Key construction is consistent on both sides (e.g. `CONCAT(country, material)`); same spelling, no leading/trailing spaces (lookup_audit reports padded keys), same case conventions.
- [ ] Exact matching for text keys. Approximate VLOOKUP/MATCH (4th/3rd arg omitted or TRUE/1) on unsorted text returns a wrong row silently (`APPROXIMATE_MATCH`).
- [ ] The lookup range covers the whole table (`TRUNCATED`) and every formula addresses the same extent (`INCONSISTENT_RANGE`; `RANGE_EXTENT_VARIES` is benign headroom).
- [ ] The returned column is the intended one — read `returned_header` (A1–A3 GWP vs "Type" vs count vs density).
- [ ] Blank selections give blank/zero *with a visible cue*, not a default factor.
- [ ] Fallback/default factors (e.g. `IFERROR(country-specific, "Global" generic)`) are documented, visible to the user (a "value used" flag), and consistent across materials.
- [ ] Regional, recycled-content, product-specific (EPD) vs generic variants are distinguishable and chosen deliberately.
- [ ] The block label agrees with the table read (`material_mapping.csv` → `CHECK: label/target material mismatch`).
- [ ] User-defined factor overrides are flagged, their units stated, and supporting documents required.
- [ ] Unit dropdowns and quantity units match the factor's declared unit (a per-kg factor with a quantity entered in m³ is wrong by the density).

Lookup forms to read carefully: XLOOKUP (match_mode, search_mode, if_not_found), VLOOKUP/HLOOKUP
(hard-coded column index — breaks when a column is inserted into the table), INDEX/MATCH,
XMATCH, FILTER (multiple matches), SUMIFS/SUMPRODUCT (adds duplicates instead of picking one —
duplicates double the factor), nested IF/CHOOSE (hard-wired material lists), OFFSET/INDIRECT
(not traceable, volatile, sheet-name fragile), named ranges and structured references.

Duplicate behaviour: VLOOKUP/MATCH/XLOOKUP return the first match (XLOOKUP search_mode −1 the
last); SUMIFS adds all matches; FILTER returns all (spills). A duplicate key is harmless only
if the duplicate records are identical.

## Emission-factor database review

Use `ef_tables.csv` / the profile table in `lookup_audit.md`, then open the tables — header
detection is keyword-based (a "Type" column may hold "Country/Global", not a source).

| Check | What to look for |
|---|---|
| Duplicates | same key, different values; same material under two names |
| Units | kgCO2e/kg vs /t vs /m³ vs /m² vs /unit within one column; declared-unit column present? |
| Geography | country/region stated; generic "Global" rows; mixed geographies under one key |
| Provenance | database / EPD number / publisher, publication year, dataset version, validity |
| Life-cycle boundary | A1–A3 only? includes A4? mixed boundaries in one column |
| Product-specific vs generic | EPD values alongside database medians — flagged? |
| Statistic | median vs mean vs percentile — consistent across materials? sample count shown? |
| Monotonicity | does the factor fall as recycled content / SCM replacement rises? reversals need a source |
| Signs | negatives (timber biogenic carbon, Module D credits) — intended and declared? |
| Zeros | a 0 factor is almost always a gap, not a real value |
| Outliers | values >50× the column median, or a factor-of-1000 jump (kg vs t) |
| Age | old datasets for fast-changing materials (aluminium, steel, grid electricity) |

Do not label an EF incorrect without evidence. Use "Requires source verification" and say
which document would settle it. Negative A1–A3 values for timber are a methodology question
(EN 15804+A2 / ISO 21930 report −1 in A1–A3 and +1 at end of life); if the calculator has no
C-stage, carrying the negative alone understates whole-life carbon — record it as a methodology
issue unless the documentation says otherwise.
