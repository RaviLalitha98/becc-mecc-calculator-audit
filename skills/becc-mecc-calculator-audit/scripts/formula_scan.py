#!/usr/bin/env python3
"""Formula integrity scan (audit stages 2, 3, 7, 15, 24).

Flags, grouped by R1C1 pattern so 5,000 copies of one formula appear once:
  ERROR_TEXT        formula text contains #REF!
  MISSING_SHEET     formula references a sheet name that does not exist
  EXTERNAL_REF      formula reads another workbook
  CACHED_ERROR      last-saved value is an Excel error (#N/A, #DIV/0!, ...)
  TEXT_IN_MATH      arithmetic operator applied to a reference that holds text
  HARDCODE          numeric constant in a formula-dominated column
  BLANK_PRECEDENT   formula refers to a single cell that is empty
  MASKED_ERROR      IFERROR/IFNA substituting 0 or "" (silent failure / omission)
  IMPLICIT_ISECT    multi-cell range where a scalar is expected in a non-array formula
  MAGIC_NUMBER      numeric literals inside formulas (unit conversions, factors, %)
  FLOAT_EQUALITY    direct '=' comparison between two references
  LEGACY_ARRAY      CSE array formulas
  DYNAMIC_ARRAY     dynamic-array formulas (cm attribute) — spill / version risk
  VOLATILE          OFFSET / INDIRECT / NOW / TODAY ...
  ROUNDING          ROUND/INT/TRUNC/... — where in the chain rounding occurs

Outputs: formula_scan.csv, formula_scan.md
Usage: python formula_scan.py WORKBOOK [--out DIR]
"""
from __future__ import annotations

import argparse
import os
import re
from collections import Counter, defaultdict

from common import (load_cells, index_cells, tokenize, Token, parse_calls, split_sheet, parse_address,
                    area_bounds, read_zip_text, header_for, row_label, a1, write_csv, ensure_dir, short)
from excel_compat import base_name, VOLATILE, ROUNDING_FUNCS
from inventory import workbook_xml_info

RANGE_OK = {"SUM", "SUMIF", "SUMIFS", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "COUNTBLANK", "AVERAGE",
            "AVERAGEIF", "AVERAGEIFS", "MAX", "MAXIFS", "MIN", "MINIFS", "VLOOKUP", "HLOOKUP", "LOOKUP",
            "MATCH", "INDEX", "XLOOKUP", "XMATCH", "SUMPRODUCT", "FILTER", "UNIQUE", "SORT", "SORTBY",
            "LARGE", "SMALL", "MEDIAN", "CONCAT", "TEXTJOIN", "AND", "OR", "PRODUCT", "SUBTOTAL",
            "AGGREGATE", "OFFSET", "ROWS", "COLUMNS", "ROW", "COLUMN", "STDEV", "STDEV.S", "STDEV.P",
            "RANK", "RANK.EQ", "PERCENTILE", "QUARTILE", "MODE", "TRANSPOSE", "MMULT", "FREQUENCY",
            "SUMSQ", "NPV", "IRR", "CHOOSE", "AREAS", "CELL", "ISREF", "TAKE", "DROP", "VSTACK", "HSTACK",
            "CHOOSECOLS", "CHOOSEROWS", "TOCOL", "TOROW", "LET", "MAP", "BYROW", "BYCOL", "REDUCE", "SCAN",
            "CORREL", "SLOPE", "INTERCEPT", "FORECAST", "TREND", "GROWTH", "LINEST", "GEOMEAN", "HARMEAN",
            "VAR", "VAR.S", "VAR.P", "DCOUNT", "DSUM", "DGET", "PERCENTILE.INC", "PERCENTILE.EXC",
            "QUARTILE.INC", "MAXA", "MINA", "AVERAGEA", "INDIRECT"}
BENIGN_LITERALS = {"0", "1", "-1", "0.0", "1.0"}
ORDER = ["ERROR_TEXT", "MISSING_SHEET", "EXTERNAL_REF", "CACHED_ERROR", "TEXT_IN_MATH", "HARDCODE",
         "BLANK_PRECEDENT", "MASKED_ERROR", "IMPLICIT_ISECT", "MAGIC_NUMBER", "FLOAT_EQUALITY",
         "LEGACY_ARRAY", "DYNAMIC_ARRAY", "VOLATILE", "ROUNDING"]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--password")
    ap.add_argument("--out", default="audit_out")
    a = ap.parse_args()
    out = ensure_dir(a.out)
    wb_f, wb_v, real, cells = load_cells(a.workbook, a.password)
    idx = index_cells(cells)
    winfo = workbook_xml_info(real)
    sheet_set = {s["name"] for s in winfo["sheets"]}

    dyn = defaultdict(set)
    for s in winfo["sheets"]:
        txt = read_zip_text(real, s["xml"]) or ""
        for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*\bcm="\d+"', txt):
            dyn[s["name"]].add(m.group(1))

    groups = defaultdict(lambda: {"cells": [], "details": set()})

    def flag(c, check, detail=""):
        g = groups[(check, c.sheet, c.r1c1 or c.formula or repr(c.value))]
        if not g["cells"] or g["cells"][-1] is not c:
            g["cells"].append(c)
        if detail:
            g["details"].add(detail)

    for c in cells:
        if not c.is_formula or c.kind == "datatable":
            continue
        f = c.formula
        toks = [t for t in (tokenize(f) or []) if t.type != Token.WSPACE]
        if "#REF!" in f:
            flag(c, "ERROR_TEXT", "#REF! inside formula text")
        if isinstance(c.value, str) and c.value.startswith("#"):
            flag(c, "CACHED_ERROR", c.value)
        calls = parse_calls(f)
        names = [base_name(x.name) for x in calls]
        for call in calls:
            bn = base_name(call.name)
            if bn in ("IFERROR", "IFNA") and len(call.args) == 2:
                fb = call.args[1].strip()
                if fb in ('0', '""', "0.0", '"-"', '"NA"', '"N/A"', ""):
                    inner = {base_name(x) for x in re.findall(r"([A-Z_.]+)\(", call.args[0].upper())}
                    kind = "lookup" if inner & {"VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "LOOKUP"} else "calc"
                    flag(c, "MASKED_ERROR", f"{bn}(…, {fb}) masks {kind} failures")
        for n in set(names) & VOLATILE:
            flag(c, "VOLATILE", n)
        for n in set(names) & ROUNDING_FUNCS:
            flag(c, "ROUNDING", n)
        skip = set()
        for call in calls:
            bn = base_name(call.name)
            if bn in ("VLOOKUP", "HLOOKUP") and len(call.args) >= 3:
                skip.add(call.args[2])
            if bn in ("INDEX", "MATCH", "XMATCH", "XLOOKUP", "CHOOSE", "LARGE", "SMALL", "SUBTOTAL",
                      "AGGREGATE", "OFFSET", "LEFT", "RIGHT", "MID", "ROUND", "ROUNDUP", "ROUNDDOWN", "TEXT"):
                skip.update(call.args[1:])
        lits = [t.value for t in toks if t.type == Token.OPERAND and t.subtype == Token.NUMBER]
        lits = [x for x in lits if x not in BENIGN_LITERALS and x not in skip]
        if lits:
            flag(c, "MAGIC_NUMBER", ", ".join(sorted(set(lits))))
        stack = []
        for i, t in enumerate(toks):
            if t.type == Token.FUNC and t.subtype == Token.OPEN:
                stack.append(base_name(t.value[:-1]))
            elif t.type == Token.PAREN and t.subtype == Token.OPEN:
                stack.append("(")
            elif t.subtype == Token.CLOSE and t.type in (Token.FUNC, Token.PAREN):
                if stack:
                    stack.pop()
            elif t.type == Token.OPERAND and t.subtype == Token.RANGE:
                book, sh, addr = split_sheet(t.value)
                kind, data = parse_address(addr)
                if book:
                    flag(c, "EXTERNAL_REF", t.value)
                    continue
                if sh is not None and sh not in sheet_set:
                    flag(c, "MISSING_SHEET", sh)
                tsheet = sh or c.sheet
                if kind == "cell":
                    tc = idx.get(tsheet, {}).get((data[0], data[1]))
                    if tc is None:
                        flag(c, "BLANK_PRECEDENT", t.value)
                    elif tc.kind == "text":
                        prev = toks[i - 1] if i else None
                        nxt = toks[i + 1] if i + 1 < len(toks) else None
                        if (prev is not None and prev.type == Token.OP_IN and prev.value in "*/+-^") or \
                           (nxt is not None and nxt.type == Token.OP_IN and nxt.value in "*/+-^"):
                            flag(c, "TEXT_IN_MATH", f"{t.value} holds text {short(tc.value, 30)!r}")
                if kind in ("area", "cols", "rows") and c.kind != "array" and a1(c.row, c.col) not in dyn.get(c.sheet, set()):
                    b = area_bounds(kind, data)
                    multi = b and (b[0] != b[2] or b[1] != b[3])
                    encl = next((x for x in reversed(stack) if x != "("), None)
                    if multi and (encl is None or encl not in RANGE_OK):
                        flag(c, "IMPLICIT_ISECT", f"{t.value} inside {encl or 'top level'}")
        for i in range(1, len(toks) - 1):
            t = toks[i]
            if t.type == Token.OP_IN and t.value == "=":
                l, r = toks[i - 1], toks[i + 1]
                if l.subtype == Token.RANGE and r.subtype == Token.RANGE:
                    flag(c, "FLOAT_EQUALITY", f"{l.value}={r.value}")
        if c.kind == "array":
            flag(c, "LEGACY_ARRAY" if a1(c.row, c.col) not in dyn.get(c.sheet, set()) else "DYNAMIC_ARRAY", c.array_ref or "")
        elif a1(c.row, c.col) in dyn.get(c.sheet, set()):
            flag(c, "DYNAMIC_ARRAY", "")

    for sheet, sidx in idx.items():
        bycol = defaultdict(list)
        for (r, col), c in sidx.items():
            bycol[col].append(c)
        for col, lst in bycol.items():
            lst.sort(key=lambda x: x.row)
            for i, c in enumerate(lst):
                if c.kind != "number":
                    continue
                win = [x for x in lst[max(0, i - 8): i + 9] if x is not c and abs(x.row - c.row) <= 10]
                nf = sum(1 for x in win if x.is_formula)
                if len(win) >= 4 and nf / len(win) >= 0.7:
                    flag(c, "HARDCODE", f"{c.value!r} among {nf}/{len(win)} formula neighbours")

    rows = []
    for (check, sheet, sig), g in groups.items():
        cs = g["cells"]; first = cs[0]; sidx = idx[sheet]
        rows.append({"check": check, "sheet": sheet, "count": len(cs),
                     "cells": ", ".join(x.addr for x in cs[:12]) + (f" … (+{len(cs) - 12})" if len(cs) > 12 else ""),
                     "example_formula": short(first.formula if first.is_formula else first.value, 250),
                     "example_cached": short(first.value, 60) if first.is_formula else "",
                     "detail": short("; ".join(sorted(g["details"])), 300),
                     "row_label": short(row_label(sidx, first.row), 60),
                     "col_header": short(header_for(sidx, first.row, first.col), 60)})
    rows.sort(key=lambda r: (ORDER.index(r["check"]) if r["check"] in ORDER else 99, r["sheet"], -r["count"]))
    write_csv(os.path.join(out, "formula_scan.csv"), rows,
              ["check", "sheet", "count", "cells", "example_formula", "example_cached", "detail", "row_label", "col_header"])

    L = [f"# Formula integrity scan — {os.path.basename(a.workbook)}\n",
         "Grouped by check → sheet → formula pattern. Counts are cells. All items are candidates for review.\n",
         "BLANK_PRECEDENT and #DIV/0! are expected wherever the workbook is an empty template (inputs not yet "
         "filled); they matter when the blank cell is NOT an input — e.g. a formula pointing one row past a data "
         "block, or at a deleted column.\n",
         "| Check | Groups | Cells |", "|---|---|---|"]
    agg = defaultdict(lambda: [0, 0])
    for r in rows:
        agg[r["check"]][0] += 1; agg[r["check"]][1] += r["count"]
    for k in ORDER:
        if k in agg:
            L.append(f"| {k} | {agg[k][0]} | {agg[k][1]} |")
    cur = None; shown = Counter()
    for r in rows:
        if r["check"] != cur:
            cur = r["check"]; L.append(f"\n## {cur}\n")
        shown[cur] += 1
        if shown[cur] > 60:
            if shown[cur] == 61:
                L.append("… more in formula_scan.csv")
            continue
        L.append(f"- **{r['sheet']}** ×{r['count']} [{r['cells']}] — {r['detail']}  \n  `{r['example_formula']}` "
                 f"(cached {r['example_cached']}) | row: {r['row_label']} | col: {r['col_header']}")
    with open(os.path.join(out, "formula_scan.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(L) + "\n")
    print(f"formula scan → {out}/formula_scan.md ; " + ", ".join(f"{k}:{v[1]}" for k, v in agg.items()))


if __name__ == "__main__":
    main()
