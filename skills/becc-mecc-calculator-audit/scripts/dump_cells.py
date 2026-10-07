#!/usr/bin/env python3
"""Dump every non-empty cell to CSV for grep/pandas work.

Columns: sheet, cell, row, col, kind, formula, r1c1, cached_value, array_ref
Usage: python dump_cells.py WORKBOOK [--out DIR] [--sheets "A,B"]
"""
import argparse
import os

from common import load_cells, write_csv, ensure_dir


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook"); ap.add_argument("--password"); ap.add_argument("--out", default="audit_out")
    ap.add_argument("--sheets")
    a = ap.parse_args()
    sheets = [s.strip() for s in a.sheets.split(",")] if a.sheets else None
    _, _, _, cells = load_cells(a.workbook, a.password, sheets)
    rows = [{"sheet": c.sheet, "cell": c.addr, "row": c.row, "col": c.col, "kind": c.kind, "formula": c.formula or "",
             "r1c1": c.r1c1 or "", "cached_value": c.value, "array_ref": c.array_ref or ""} for c in cells]
    ensure_dir(a.out)
    p = os.path.join(a.out, "cells.csv")
    write_csv(p, rows, ["sheet", "cell", "row", "col", "kind", "formula", "r1c1", "cached_value", "array_ref"])
    print(f"{len(rows)} cells → {p}")


if __name__ == "__main__":
    main()
