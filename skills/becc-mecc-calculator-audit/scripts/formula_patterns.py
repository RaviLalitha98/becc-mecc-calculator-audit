#!/usr/bin/env python3
"""Formula-pattern anomaly detection (audit stages 2, 3, 12).

Every formula is converted to relative R1C1 form, so a formula copied down a
column has the same signature in every row. A cell whose signature differs
from its neighbours is a *candidate* anomaly: an off-by-one reference, a
wrong column/sheet, a formula overwritten by a constant, or simply the edge of
a block. Nothing here is a confirmed error; every hit needs human review.

Detection rules (applied vertically and horizontally):
  SANDWICH   the nearest neighbours on both sides share signature S, this cell
             differs, and S is well supported in the window. Strongest signal.
  MAJORITY   >= --majority share of the cells in a +/- --window neighbourhood
             share S and this cell differs (and it is not a block edge).
  HARDCODE   a numeric constant sitting inside a run of formulas.
  SWAPPED PAIR  adjacent cells with opposite +1/-1 offsets (references exchanged).

For each hit the script reconstructs the formula the pattern would predict for
that cell, describes the differing references (row offset, column, sheet,
$-anchoring only), and shows whether referenced cells are blank.

Outputs: pattern_anomalies.csv, pattern_anomalies.md, pattern_map.csv
Usage: python formula_patterns.py WORKBOOK [--sheets "A,B"] [--out DIR]
"""
from __future__ import annotations

import argparse
import os
from collections import Counter, defaultdict

from common import (load_cells, index_cells, r1c1_to_a1, tokenize, Token, parse_address, split_sheet,
                    area_bounds, header_for, row_label, a1, write_csv, ensure_dir, short)
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

CONST = "<CONST>"


def signature(c):
    if c is None:
        return None
    if c.is_formula:
        return c.r1c1 or c.formula
    if c.kind == "number":
        return CONST
    return None  # text/bool/date are labels — they break runs


def neighbours(line, pos, window, gap):
    before, after = [], []
    keys = line["keys"]
    i = line["index"][pos]
    j, last = i - 1, pos
    while j >= 0 and len(before) < window:
        k = keys[j]
        if last - k > gap + 1:
            break
        s = line["sig"][k]
        if s is None:
            break
        before.append((k, s)); last = k; j -= 1
    j, last = i + 1, pos
    while j < len(keys) and len(after) < window:
        k = keys[j]
        if k - last > gap + 1:
            break
        s = line["sig"][k]
        if s is None:
            break
        after.append((k, s)); last = k; j += 1
    return before, after


def describe_diff(actual: str, expected: str, sheet: str, idx_all):
    ta = [t for t in (tokenize(actual) or []) if t.type != Token.WSPACE]
    te = [t for t in (tokenize(expected) or []) if t.type != Token.WSPACE]
    if len(ta) != len(te) or any((x.type, x.subtype) != (y.type, y.subtype) for x, y in zip(ta, te)):
        return "structurally different formula (different functions/operators/argument count)", "structural"
    notes, kinds = [], set()
    for x, y in zip(ta, te):
        if x.value == y.value:
            continue
        if x.type == Token.OPERAND and x.subtype == Token.RANGE:
            bx, sx, ax = split_sheet(x.value)
            by, sy, ay = split_sheet(y.value)
            kx, dx = parse_address(ax)
            ky, dy = parse_address(ay)
            if (sx or sheet) != (sy or sheet):
                notes.append(f"{x.value} → expected {y.value} (different sheet)"); kinds.add("sheet")
                continue
            bx_ = area_bounds(kx, dx); by_ = area_bounds(ky, dy)
            if bx_ and by_ and bx_ == by_:
                notes.append(f"{x.value} vs {y.value}: same cells, different $-anchoring"); kinds.add("anchoring")
                continue
            if bx_ and by_:
                dr = (bx_[0] - by_[0], bx_[2] - by_[2]); dc = (bx_[1] - by_[1], bx_[3] - by_[3])
                desc = f"{x.value} → expected {y.value}"
                if dc == (0, 0) and dr[0] == dr[1]:
                    desc += f" (row offset {dr[0]:+d})"; kinds.add("row_shift_1" if abs(dr[0]) == 1 else "row_shift")
                elif dr == (0, 0) and dc[0] == dc[1]:
                    desc += f" (column offset {dc[0]:+d})"; kinds.add("col_shift")
                elif dc == (0, 0) and dr[0] == 0:
                    desc += f" (range end differs by {dr[1]:+d} rows)"; kinds.add("range_end")
                else:
                    kinds.add("other_ref")
                tgt_sheet = sx or sheet
                if kx == "cell" and idx_all.get(tgt_sheet, {}).get((dx[0], dx[1])) is None:
                    desc += " [referenced cell is BLANK]"
                notes.append(desc)
            else:
                notes.append(f"{x.value} → expected {y.value}"); kinds.add("other_ref")
        elif x.type == Token.OPERAND and x.subtype == Token.NUMBER:
            notes.append(f"literal {x.value} vs {y.value}"); kinds.add("literal")
        else:
            notes.append(f"'{x.value}' vs '{y.value}'"); kinds.add("token")
    return "; ".join(notes) or "identical after normalisation", "+".join(sorted(kinds)) or "none"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--password")
    ap.add_argument("--sheets", help="comma-separated subset of sheets")
    ap.add_argument("--out", default="audit_out")
    ap.add_argument("--window", type=int, default=6)
    ap.add_argument("--majority", type=float, default=0.75)
    ap.add_argument("--gap", type=int, default=1, help="blank rows/cols tolerated inside a run")
    a = ap.parse_args()
    out = ensure_dir(a.out)
    sheets = [s.strip() for s in a.sheets.split(",")] if a.sheets else None
    wb_f, wb_v, real, cells = load_cells(a.workbook, a.password)
    idx_all = index_cells(cells)
    targets = sheets or list(idx_all.keys())

    hits = {}
    pattern_rows = []
    for sheet in targets:
        idx = idx_all.get(sheet, {})
        cols = defaultdict(dict); rows = defaultdict(dict)
        for (r, c), cell in idx.items():
            s = signature(cell)
            cols[c][r] = s; rows[r][c] = s
        for axis, lines in (("vertical", cols), ("horizontal", rows)):
            for line_id, sigmap in lines.items():
                keys = sorted(sigmap)
                line = {"keys": keys, "index": {k: i for i, k in enumerate(keys)}, "sig": sigmap}
                for pos in keys:
                    s = sigmap[pos]
                    if s is None:
                        continue
                    before, after = neighbours(line, pos, a.window, a.gap)
                    if not before and not after:
                        continue
                    rule = None; expected_sig = None; support = 0; total = 0
                    if before and after and before[0][1] == after[0][1] and before[0][1] != s and before[0][1] != CONST:
                        sup = sum(1 for _, x in before + after if x == before[0][1])
                        tot = len(before) + len(after)
                        # two coincidentally-equal neighbours in an otherwise varied line is noise
                        if sup >= 3 or sup / tot >= 0.5:
                            rule = "SANDWICH"; expected_sig = before[0][1]; support, total = sup, tot
                    if rule is None and before and after:
                        cnt = Counter(x for _, x in before + after)
                        top, n = cnt.most_common(1)[0]
                        total = len(before) + len(after)
                        if top != s and top != CONST and n >= 4 and n / total >= a.majority:
                            rule = "MAJORITY"; expected_sig = top; support = n
                    if rule is None:
                        continue
                    if s == CONST:
                        rule = "HARDCODE_" + rule
                    r, c = (pos, line_id) if axis == "vertical" else (line_id, pos)
                    cell = idx[(r, c)]
                    expected_a1 = "=" + r1c1_to_a1(expected_sig[1:], r, c) if expected_sig.startswith("=") else expected_sig
                    if cell.is_formula:
                        diff, dkind = describe_diff(cell.formula, expected_a1, sheet, idx_all)
                    else:
                        diff, dkind = f"constant {cell.value!r} where neighbours hold formulas", "hardcode"
                    nb = (lambda k: idx[(k, c)] if axis == "vertical" else idx[(r, k)])
                    record = {
                        "sheet": sheet, "cell": a1(r, c), "axis": axis, "rule": rule,
                        "support": f"{support}/{total}", "diff_kind": dkind,
                        "current": cell.formula if cell.is_formula else repr(cell.value),
                        "expected_pattern_a1": expected_a1, "difference": diff,
                        "cached_value": short(cell.value, 60) if cell.is_formula else "",
                        "row_label": short(row_label(idx, r), 60),
                        "col_header": short(header_for(idx, r, c), 60),
                        "neighbour_before": (nb(before[0][0]).formula or repr(nb(before[0][0]).value)) if before else "",
                        "neighbour_after": (nb(after[0][0]).formula or repr(nb(after[0][0]).value)) if after else "",
                    }
                    key = (sheet, r, c)
                    entry = hits.get(key)
                    if entry is None or (entry["rule"].endswith("MAJORITY") and rule.endswith("SANDWICH")):
                        if entry is not None:
                            record["axis"] = "both"
                        hits[key] = record
                    else:
                        entry["axis"] = "both"
        for c, sigmap in cols.items():
            fam = defaultdict(list)
            for r, s in sigmap.items():
                if s and s != CONST:
                    fam[s].append(r)
            for s, rs in fam.items():
                rs.sort()
                runs = []; start = prev = rs[0]
                for x in rs[1:]:
                    if x != prev + 1:
                        runs.append((start, prev)); start = x
                    prev = x
                runs.append((start, prev))
                pattern_rows.append({"sheet": sheet, "column": get_column_letter(c), "count": len(rs),
                                     "rows": ", ".join(f"{p}-{q}" if p != q else f"{p}" for p, q in runs[:20]) + (" …" if len(runs) > 20 else ""),
                                     "header": short(header_for(idx, rs[0], c), 50),
                                     "signature_r1c1": short(s, 300)})

    def triage(h):
        k = h["diff_kind"]
        if h["rule"].startswith("HARDCODE"):
            return "review-hardcode"
        if k == "anchoring":
            return "robustness"
        sup, tot = (int(x) for x in h["support"].split("/"))
        if k in ("row_shift_1", "sheet", "col_shift", "row_shift_1+sheet") and (h["rule"] == "SANDWICH" or sup / tot >= 0.6):
            return "likely-defect"
        if "structural" in k:
            return "intentional-or-defect"
        return "review"

    rows_out = sorted(hits.values(), key=lambda h: (h["sheet"], h["cell"]))
    for h in rows_out:
        h["triage_hint"] = triage(h)
    by_key = {}
    for h in rows_out:
        col, row = coordinate_from_string(h["cell"])
        by_key[(h["sheet"], row, column_index_from_string(col))] = h
    for (s, r, c), h in by_key.items():
        for dr, dc in ((1, 0), (0, 1)):
            o = by_key.get((s, r + dr, c + dc))
            if o and "offset +1" in h["difference"] and "offset -1" in o["difference"]:
                for x, y in ((h, o), (o, h)):
                    if "SWAPPED" not in x["difference"]:
                        x["difference"] += f" | SWAPPED PAIR with {y['cell']}"
                    x["triage_hint"] = "likely-defect"

    fields = ["sheet", "cell", "rule", "axis", "triage_hint", "diff_kind", "support", "current",
              "expected_pattern_a1", "difference", "cached_value", "row_label", "col_header",
              "neighbour_before", "neighbour_after"]
    write_csv(os.path.join(out, "pattern_anomalies.csv"), rows_out, fields)
    write_csv(os.path.join(out, "pattern_map.csv"), pattern_rows,
              ["sheet", "column", "count", "rows", "header", "signature_r1c1"])

    L = [f"# Formula pattern anomalies — {os.path.basename(a.workbook)}\n",
         "Candidates only. `likely-defect` = isolated break whose only difference is a shifted/other-sheet reference "
         "(or a swapped pair); `robustness` = same cells, different $-anchoring; `intentional-or-defect` = different "
         "formula structure (often a deliberate special case, sometimes an overwrite). Verify each against row labels "
         "and headers — the majority pattern can itself be wrong.\n"]
    cnt = Counter((h["sheet"], h["triage_hint"]) for h in rows_out)
    L.append("| Sheet | Triage | Count |\n|---|---|---|")
    for (s, t), n in sorted(cnt.items()):
        L.append(f"| {s} | {t} | {n} |")
    order = {"likely-defect": 0, "review-hardcode": 1, "review": 2, "intentional-or-defect": 3, "robustness": 4}
    for h in sorted(rows_out, key=lambda h: (order.get(h["triage_hint"], 9), h["sheet"]))[:400]:
        L.append(f"\n### {h['sheet']}!{h['cell']} — {h['triage_hint']} ({h['rule']}, {h['axis']}, support {h['support']})")
        L.append(f"- Row: {h['row_label']} | Column: {h['col_header']}")
        L.append(f"- Current: `{short(h['current'], 300)}`  (cached: {h['cached_value']})")
        L.append(f"- Pattern predicts: `{short(h['expected_pattern_a1'], 300)}`")
        L.append(f"- Difference: {short(h['difference'], 400)}")
    if len(rows_out) > 400:
        L.append(f"\n… {len(rows_out) - 400} more in pattern_anomalies.csv")
    with open(os.path.join(out, "pattern_anomalies.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(L) + "\n")
    print(f"{len(rows_out)} candidate anomalies → {out}/pattern_anomalies.md; "
          + str(dict(Counter(h['triage_hint'] for h in rows_out))))


if __name__ == "__main__":
    main()
