#!/usr/bin/env python3
"""Totals / subtotals audit (stage 8, parts of 6 and 9).

For every aggregation (SUM, SUMIF(S), SUMPRODUCT, SUBTOTAL, AGGREGATE and
array SUM(IF(...))) over a one-dimensional range:
  SIBLING_MISMATCH  sibling totals in the same row (or column) cover a different
                    extent — e.g. one stage total sums rows 27:577 while the
                    others sum 27:601
  EXCLUDES_ROWS     populated numeric/formula cells sit just beyond the range end
                    (or before its start) — "table grew, total did not"
  NESTED_SUBTOTAL   the range contains a cell that aggregates part of the same
                    range → double counting risk
  SELF_REFERENCE    the range contains the host cell (circular)
  CACHED_MISMATCH   for plain SUM(range): sum of cached values != cached total
                    → stale results (manual calc) or values changed after last calc
  SELECTIVE_SUM     SUM of many hand-picked areas of one column — verify the
                    selection and that it is updated when rows are added
  EXCLUDES_COLS?    populated cells directly right of a horizontal range
  HIDDEN_ROWS       range spans hidden rows
  INCLUDES_TEXT     header/label text inside the summed range

Outputs: aggregation_check.csv, aggregation_check.md
Usage: python aggregation_check.py WORKBOOK [--out DIR] [--lookahead 30]
"""
from __future__ import annotations

import argparse
import os
import re
from collections import Counter, defaultdict

from common import (load_cells, index_cells, parse_calls, refs_in_formula, area_bounds, a1,
                    header_for, row_label, write_csv, ensure_dir, short, Token)
from excel_compat import base_name
from inventory import workbook_xml_info, sheet_xml_info

AGG = {"SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUBTOTAL", "AGGREGATE"}
ORDER = ["SIBLING_MISMATCH", "EXCLUDES_ROWS", "NESTED_SUBTOTAL", "SELF_REFERENCE", "CACHED_MISMATCH",
         "SELECTIVE_SUM", "EXCLUDES_COLS?", "HIDDEN_ROWS", "INCLUDES_TEXT"]


def agg_ranges(formula, host_sheet):
    out, seen = [], set()
    for call in parse_calls(formula):
        if base_name(call.name) not in AGG:
            continue
        for toks in call.arg_tokens:
            for t in toks:
                if t.type == Token.OPERAND and t.subtype == Token.RANGE:
                    for op, book, sh, kind, data in refs_in_formula("=" + t.value):
                        if book or kind not in ("area", "cols", "rows"):
                            continue
                        b = area_bounds(kind, data)
                        key = (sh or host_sheet, *b)
                        if key not in seen:
                            seen.add(key)
                            out.append((base_name(call.name), sh or host_sheet, *b, t.value))
    return out


def is_agg_cell(cell):
    if cell is None or not cell.is_formula:
        return False
    calls = parse_calls(cell.formula)
    return bool(calls) and base_name(calls[0].name) in AGG


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--password")
    ap.add_argument("--out", default="audit_out")
    ap.add_argument("--lookahead", type=int, default=30)
    a = ap.parse_args()
    out = ensure_dir(a.out)
    wb_f, wb_v, real, cells = load_cells(a.workbook, a.password)
    idx = index_cells(cells)
    winfo = workbook_xml_info(real)
    hidden_rows = {s["name"]: set(sheet_xml_info(real, s["xml"])["hidden_rows"]) for s in winfo["sheets"]}

    def contributes(cell):
        return cell is not None and (cell.is_formula or cell.kind == "number")

    def rec(c, fn, rng, tsh, issues, detail):
        return {"sheet": c.sheet, "cell": c.addr, "function": fn, "range": rng, "target_sheet": tsh,
                "issues": ", ".join(sorted(set(issues))), "detail": "; ".join(detail),
                "formula": short(c.formula, 250), "cached": short(c.value, 40), "sig": c.r1c1,
                "row_label": short(row_label(idx[c.sheet], c.row), 60),
                "col_header": short(header_for(idx[c.sheet], c.row, c.col), 60)}

    findings = []
    per_row = defaultdict(list)
    per_col = defaultdict(list)
    examined = 0
    for c in cells:
        if not c.is_formula:
            continue
        rngs = agg_ranges(c.formula, c.sheet)
        percol = Counter((x[1], x[3]) for x in rngs if x[3] == x[5])
        selective = {k for k, n in percol.items() if n > 2}
        for tsh_, col_ in selective:
            findings.append(rec(c, "SUM", f"{percol[(tsh_, col_)]} areas", tsh_, ["SELECTIVE_SUM"],
                                [f"sums {percol[(tsh_, col_)]} hand-picked areas of one column — verify the selection matches "
                                 "the intended materials/rows and is updated when rows are added"]))
        for fn, tsh, r1, c1, r2, c2, text in rngs:
            if (tsh, c1) in selective and c1 == c2:
                continue
            examined += 1
            if r2 >= 1048576 or c2 >= 16384:
                continue
            sidx = idx.get(tsh, {})
            issues, detail = [], []
            if tsh == c.sheet and r1 <= c.row <= r2 and c1 <= c.col <= c2:
                issues.append("SELF_REFERENCE")
            if c1 == c2 and r2 > r1:
                col = c1
                beyond, r, gaps = [], r2 + 1, 0
                while r <= r2 + a.lookahead:
                    if tsh == c.sheet and r == c.row and col == c.col:
                        break
                    cell = sidx.get((r, col))
                    if contributes(cell):
                        if is_agg_cell(cell):
                            break
                        beyond.append(r); gaps = 0
                    elif cell is not None and cell.kind == "text" and str(cell.value).strip():
                        break
                    else:
                        gaps += 1
                        if gaps > 3:
                            break
                    r += 1
                before, r = [], r1 - 1
                while r >= max(1, r1 - 3):
                    cell = sidx.get((r, col))
                    if contributes(cell) and not (tsh == c.sheet and r == c.row and col == c.col):
                        before.append(r)
                    else:
                        break
                    r -= 1
                if beyond:
                    issues.append("EXCLUDES_ROWS")
                    nz = [x for x in beyond if sidx[(x, col)].value not in (0, None, "")]
                    detail.append(f"{len(beyond)} populated cell(s) after range end (rows {beyond[0]}–{beyond[-1]}; "
                                  f"{len(nz)} with non-zero cached value)")
                if before:
                    issues.append("EXCLUDES_ROWS")
                    detail.append(f"populated cell(s) just before range start (rows {before})")
                texts = [r for r in range(r1, r2 + 1) if sidx.get((r, col)) is not None and sidx[(r, col)].kind == "text"
                         and str(sidx[(r, col)].value).strip()]
                if texts:
                    issues.append("INCLUDES_TEXT"); detail.append(f"text in rows {texts[:5]}")
                hr = [r for r in range(r1, r2 + 1) if r in hidden_rows.get(tsh, set())]
                if hr:
                    issues.append("HIDDEN_ROWS"); detail.append(f"{len(hr)} hidden row(s) inside range")
                nested = []
                for r in range(r1, r2 + 1):
                    inner = sidx.get((r, col))
                    if inner is not None and inner.is_formula and (inner.sheet, r, col) != (c.sheet, c.row, c.col):
                        for _, ish, ir1, ic1, ir2, ic2, _t in agg_ranges(inner.formula, tsh):
                            if ish == tsh and ic1 <= col <= ic2 and not (ir2 < r1 or ir1 > r2):
                                nested.append(inner.addr); break
                if nested:
                    issues.append("NESTED_SUBTOTAL")
                    detail.append(f"cells inside range that aggregate the same range: {nested[:6]}")
                per_row[(c.sheet, c.row)].append((c.col, r1, r2, tsh, c))
            elif r1 == r2 and c2 > c1:
                per_col[(c.sheet, c.col)].append((c.row, c1, c2, tsh, c))
                beyond, k = [], c2 + 1
                while k <= c2 + 5:
                    cell = sidx.get((r1, k))
                    if contributes(cell) and not (tsh == c.sheet and r1 == c.row and k == c.col):
                        beyond.append(k); k += 1
                    else:
                        break
                if beyond and fn == "SUM":
                    issues.append("EXCLUDES_COLS?")
                    detail.append(f"populated cells directly right of range end: {[a1(r1, x) for x in beyond]}")
            if fn == "SUM" and re.fullmatch(r"=\s*SUM\([^(),]+\)\s*", c.formula or "") and isinstance(c.value, (int, float)):
                tot = 0.0
                for r in range(r1, r2 + 1):
                    for k in range(c1, c2 + 1):
                        v = sidx.get((r, k))
                        if v is not None and isinstance(v.value, (int, float)) and not isinstance(v.value, bool):
                            tot += v.value
                if abs(tot - c.value) > 1e-6 * max(1.0, abs(tot), abs(c.value)):
                    issues.append("CACHED_MISMATCH"); detail.append(f"cached total {c.value} vs recomputed {tot}")
            if issues:
                findings.append(rec(c, fn, text, tsh, issues, detail))

    def sibling(groups, axis):
        for key, lst in groups.items():
            if len(lst) < 3:
                continue
            ext = Counter((x[1], x[2]) for x in lst)
            (m1, m2), n = ext.most_common(1)[0]
            if n == len(lst) or n / len(lst) < 0.6:
                continue
            for pos, e1, e2, tsh, c in lst:
                if (e1, e2) != (m1, m2):
                    findings.append(rec(c, "(sibling)", f"{e1}:{e2}", tsh, ["SIBLING_MISMATCH"],
                                        [f"{axis}: {n}/{len(lst)} sibling totals span {m1}:{m2}; this one spans {e1}:{e2}"]))
    sibling(per_row, "rows covered, compared across the same row")
    sibling(per_col, "columns covered, compared down the same column")

    findings.sort(key=lambda f: (min(ORDER.index(i.strip()) for i in f["issues"].split(",")), f["sheet"]))
    write_csv(os.path.join(out, "aggregation_check.csv"), findings,
              ["sheet", "cell", "function", "range", "target_sheet", "issues", "detail", "formula", "cached", "row_label", "col_header"])
    cnt = Counter(i.strip() for f in findings for i in f["issues"].split(","))
    grouped = {}
    for f in findings:
        k = (f["sheet"], f["issues"], f["sig"])
        if k not in grouped:
            grouped[k] = dict(f, cells=[f["cell"]])
        elif f["cell"] not in grouped[k]["cells"]:
            grouped[k]["cells"].append(f["cell"])
    L = [f"# Aggregation / totals check — {os.path.basename(a.workbook)}\n",
         f"{examined} aggregation ranges examined. Issue counts: " + ", ".join(f"{k}: {v}" for k, v in cnt.most_common()) + "\n",
         f"{len(findings)} flagged cells in {len(grouped)} distinct pattern groups. INCLUDES_TEXT and EXCLUDES_ROWS often "
         "reflect a column header or the next block; SIBLING_MISMATCH and SELECTIVE_SUM are the high-value leads.\n"]
    for g in list(grouped.values())[:250]:
        cl = g["cells"]
        cells_txt = ", ".join(cl[:8]) + (f" … (+{len(cl) - 8})" if len(cl) > 8 else "")
        L.append(f"- **{g['sheet']}** [{cells_txt}] {g['issues']} — {g['detail']}  \n  e.g. {g['cell']}: `{g['formula']}` "
                 f"(cached {g['cached']}) | row: {g['row_label']} | col: {g['col_header']}")
    with open(os.path.join(out, "aggregation_check.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(L) + "\n")
    print(f"aggregation check → {out}/aggregation_check.md ; {dict(cnt)}")


if __name__ == "__main__":
    main()
