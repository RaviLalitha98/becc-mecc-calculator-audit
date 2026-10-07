#!/usr/bin/env python3
"""Dependency tracing (stages 4, 17).

  precedents  — walk upstream from an output cell to its inputs: prints a
                depth-limited tree (formula + cached value + row/column labels at
                every node) and a leaf summary: which constant/input cells and
                which sheets ultimately feed the result.
  dependents  — walk downstream: which formulas (and finally which outputs)
                use a given cell or range.

Static parse only: INDIRECT/OFFSET targets, external links and VBA-driven values
cannot be resolved and are reported as DYNAMIC. Ranges are expanded to their
populated cells.

Usage:
  python trace.py WORKBOOK precedents "'Results Tables'!B3" [--depth 6] [--out DIR]
  python trace.py WORKBOOK dependents "Data_Steel!E5" [--depth 4]
"""
from __future__ import annotations

import argparse
import os
import re
from collections import Counter, defaultdict

from common import (load_cells, index_cells, refs_in_formula, area_bounds, split_sheet, parse_address,
                    header_for, row_label, a1, ensure_dir, short, functions_in_formula)
from excel_compat import base_name


def parse_target(t: str):
    book, sh, addr = split_sheet(t.strip())
    kind, data = parse_address(addr.replace("$", ""))
    b = area_bounds(kind, data)
    if b is None:
        raise SystemExit(f"cannot parse reference {t}")
    return sh, b


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("mode", choices=["precedents", "dependents"])
    ap.add_argument("target")
    ap.add_argument("--password")
    ap.add_argument("--depth", type=int, default=6)
    ap.add_argument("--max-children", type=int, default=25)
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    wb_f, wb_v, real, cells = load_cells(a.workbook, a.password)
    idx = index_cells(cells)
    names = {n.upper(): dn.attr_text for n, dn in wb_f.defined_names.items()}
    sheet, (r1, c1, r2, c2) = parse_target(a.target)
    if sheet not in wb_f.sheetnames:
        raise SystemExit(f"sheet {sheet!r} not found; sheets: {wb_f.sheetnames}")
    lines = []

    def label(sh, r, c):
        s = idx.get(sh, {})
        return f"[{short(row_label(s, r), 45)} / {short(header_for(s, r, c), 35)}]"

    def node_text(sh, r, c):
        cell = idx.get(sh, {}).get((r, c))
        if cell is None:
            return f"'{sh}'!{a1(r, c)} = <blank>"
        if cell.is_formula:
            return f"'{sh}'!{a1(r, c)} `{short(cell.formula, 160)}` → {short(cell.value, 30)} {label(sh, r, c)}"
        return f"'{sh}'!{a1(r, c)} = {short(cell.value, 40)} (constant) {label(sh, r, c)}"

    def children(sh, r, c):
        cell = idx.get(sh, {}).get((r, c))
        if cell is None or not cell.is_formula:
            return [], []
        kids, dyn = [], []
        fns = {base_name(f) for f in functions_in_formula(cell.formula)}
        if fns & {"INDIRECT", "OFFSET"}:
            dyn.append(", ".join(sorted(fns & {"INDIRECT", "OFFSET"})))
        for op, book, tsh, kind, data in refs_in_formula(cell.formula):
            if book:
                dyn.append(f"external {op}"); continue
            if kind == "other":
                nm = names.get(op.upper())
                if nm:
                    for op2, b2, s2, k2, d2 in refs_in_formula("=" + nm):
                        bb = area_bounds(k2, d2)
                        if bb and not b2:
                            kids.append((s2 or sh, bb, f"name {op}"))
                        elif b2:
                            dyn.append(f"name {op} → external {op2}")
                continue
            bb = area_bounds(kind, data)
            if bb:
                kids.append((tsh or sh, bb, op))
        uniq, seen = [], set()
        for k in kids:
            if (k[0], k[1]) not in seen:
                seen.add((k[0], k[1])); uniq.append(k)
        return uniq, dyn

    leaves, sheets_touched, dynamic, visited = Counter(), Counter(), [], set()

    def walk_up(sh, r, c, depth, prefix):
        key = (sh, r, c)
        lines.append(prefix + node_text(sh, r, c) + (" (seen)" if key in visited else ""))
        if key in visited:
            return
        visited.add(key)
        kids, dyn = children(sh, r, c)
        for d in dyn:
            dynamic.append(f"'{sh}'!{a1(r, c)}: {d}")
            lines.append(prefix + "  ⚠ DYNAMIC: " + d)
        cell = idx.get(sh, {}).get((r, c))
        if cell is not None and not cell.is_formula:
            leaves[(sh, header_for(idx[sh], r, c))] += 1
        if depth >= a.depth:
            if kids:
                lines.append(prefix + f"  … depth limit ({len(kids)} refs not expanded)")
            return
        for tsh, (q1, k1, q2, k2), op in kids:
            sheets_touched[tsh] += 1
            sidx = idx.get(tsh, {})
            if q2 >= 1048576 or k2 >= 16384:
                lines.append(prefix + f"  ↳ {op}: whole row/column reference (not expanded)")
                continue
            pop = sorted((rr, cc) for (rr, cc) in sidx if q1 <= rr <= q2 and k1 <= cc <= k2)
            total = (q2 - q1 + 1) * (k2 - k1 + 1)
            if total > 1:
                lines.append(prefix + f"  ↳ range {op}: {len(pop)} populated of {total}")
            for n, (rr, cc) in enumerate(pop):
                if n >= a.max_children:
                    lines.append(prefix + f"    … {len(pop) - n} more cells in {op}")
                    break
                walk_up(tsh, rr, cc, depth + 1, prefix + "    ")

    if a.mode == "precedents":
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                walk_up(sheet, r, c, 0, "")
        lines.append("\n## Leaf (constant) inputs reached, by sheet / column header")
        for (sh, h), n in leaves.most_common(60):
            lines.append(f"- {sh} / {short(h, 60)}: {n} cell(s)")
        lines.append("\n## Sheets touched: " + ", ".join(f"{k} ({v})" for k, v in sheets_touched.most_common()))
        if dynamic:
            lines.append("\n## Untraceable (dynamic/external) links\n" + "\n".join("- " + d for d in dynamic[:50]))
    else:
        rev = defaultdict(list)
        for cell in cells:
            if not cell.is_formula:
                continue
            for op, book, tsh, kind, data in refs_in_formula(cell.formula):
                bb = area_bounds(kind, data)
                if bb and not book:
                    rev[tsh or cell.sheet].append((bb, cell))
        seen = set()

        def walk_down(sh, b, depth, prefix):
            q1, k1, q2, k2 = b
            users, uk = [], set()
            for (x1, y1, x2, y2), cell in rev.get(sh, []):
                if not (x2 < q1 or x1 > q2 or y2 < k1 or y1 > k2) and (cell.sheet, cell.row, cell.col) not in uk:
                    uk.add((cell.sheet, cell.row, cell.col)); users.append(cell)
            for u in users[: a.max_children]:
                key = (u.sheet, u.row, u.col)
                lines.append(prefix + node_text(u.sheet, u.row, u.col) + (" (seen)" if key in seen else ""))
                if key in seen or depth >= a.depth:
                    continue
                seen.add(key)
                walk_down(u.sheet, (u.row, u.col, u.row, u.col), depth + 1, prefix + "    ")
            if len(users) > a.max_children:
                lines.append(prefix + f"… {len(users) - a.max_children} more dependents")
        lines.append(f"Dependents of '{sheet}'!{a1(r1, c1)}" + (f":{a1(r2, c2)}" if (r1, c1) != (r2, c2) else ""))
        walk_down(sheet, (r1, c1, r2, c2), 0, "  ")
    text = "\n".join(lines)
    print(text)
    if a.out:
        ensure_dir(a.out)
        safe = re.sub(r"[^A-Za-z0-9]+", "_", a.target)[:60]
        with open(os.path.join(a.out, f"trace_{a.mode}_{safe}.md"), "w", encoding="utf-8") as fh:
            fh.write(f"# {a.mode} of {a.target}\n\n" + text + "\n")


if __name__ == "__main__":
    main()
