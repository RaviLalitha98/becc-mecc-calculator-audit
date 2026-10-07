# Unit consistency (stage 7)

Write the unit next to every term of every module chain, before and after each formula. Most
unit faults are invisible in formula text and only show as a factor of 10ⁿ.

| Quantity | Units seen in BECC/MECC tools |
|---|---|
| Emission factor | kgCO2e/kg, kgCO2e/t, kgCO2e/m³, kgCO2e/m², kgCO2e/unit, kgCO2e/kWh, kgCO2e/L, kgCO2e/MJ, kgCO2e/t·km, kgCO2e/kg·km |
| Quantity | kg, t, m³, m², m, units, L, kWh, MJ |
| Transport | km, NM (×1.852 → km), miles (×1.609), t·km |
| Result | kgCO2e, tCO2e, kgCO2e/m² GFA |

Checklist:
- [ ] kg ↔ t (×/÷1000) applied exactly once. Unit-selection factor columns such as `IF(unit="kg",1,IF(unit="Tonne",1000,0))` return **0** for any unlisted unit → silent zero (test "tonnes", "t", "TONNE", "m3").
- [ ] The unit list offered matches the factor's declared unit (a kg/tonne dropdown on a material whose quantity is naturally m³).
- [ ] g ↔ kg, m ↔ km, NM ↔ km, kgCO2e ↔ tCO2e (results and chart labels consistent).
- [ ] volume × density → mass: density units (kg/m³), density per material (concrete ~2400, steel 7850, aluminium 2700, glass 2500, timber 400–700) and whether reinforcement is inside the concrete volume (double counting steel mass).
- [ ] transport: mass (t) × distance (km) × factor (kgCO2e/t·km). A per-kg·km factor applied to tonnes is ×1000 out.
- [ ] energy × factor: kWh × kgCO2e/kWh; diesel L × kgCO2e/L vs m³ (×1000); MJ ↔ kWh (÷3.6).
- [ ] percentages stored as 5 vs 0.05 (wastage, recycled content).
- [ ] normalisation area unit (m²) and which area is used.

Suspicious factors: results that differ by 10, 100, 1,000 or 1,000,000 from a hand
calculation, and magic numbers in `formula_scan.md` (1000, 0.001, 3.6, 1.852, 2400, 7850,
44/12, 100). Each needs a stated unit reason.

Report each suspicious calculation as: `formula` — units in → units out — expected units — effect.
