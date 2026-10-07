#!/usr/bin/env python3
"""Scenario testing by recalculation (stages 13, 14, 25, 28).

Opens a COPY of each workbook in headless LibreOffice (via UNO), writes test
inputs, recalculates, and reads outputs. Run the same scenario file against
two calculator versions to reconcile differences stage by stage.

The original workbook is never modified. LibreOffice is NOT Excel: read the
"engine caveats" at the top of the output (functions the installed LibreOffice
cannot evaluate). A result that differs between Excel and LibreOffice is a
compatibility observation, not proof of a calculator defect.

Scenario file (JSON) — see assets/scenarios_template.json:
{"scenarios": [{"name": "...", "description": "...",
   "inputs":  {"Sheet!E28": "Global", "Sheet!I28": 1000},
   "outputs": ["Sheet!N28", "Results!C3"],
   "expect":  {"Sheet!N28": {"equals_product_of": ["Sheet!I28", "Sheet!M28"], "tol": 1e-9},
               "Results!C3": {"same_as": "Sheet!N28"},
               "Results!C9": {"value": 0.0, "tol": 1e-6},
               "Results!C10": {"is_error": false}}}]}
Inputs accept numbers, text, booleans or null (clears the cell). Each input is
checked against the sheet's data-validation rule so the report shows whether
Excel would have rejected the entry.

Usage:
  python recalc.py --scenarios tests.json --workbook V1.xlsx [--workbook V2.xlsx] [--out DIR] [--save-copies]
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import socket
import subprocess
import tempfile
import time
from collections import defaultdict

from common import split_sheet, ensure_dir, write_csv, short, decrypt_if_needed, load_cells, functions_in_formula
from excel_compat import LIBREOFFICE_FIRST, base_name
from inventory import workbook_xml_info, sheet_xml_info


def lo_version():
    try:
        out = subprocess.run(["soffice", "--version"], capture_output=True, text=True, timeout=60).stdout
        m = re.search(r"(\d+\.\d+)", out)
        return m.group(1) if m else None
    except Exception:
        return None


def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p


class LO:
    def __init__(self):
        import uno
        self.port = free_port()
        self.profile = tempfile.mkdtemp(prefix="lo_profile_")
        self.proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
                                      f"-env:UserInstallation=file://{self.profile}",
                                      f"--accept=socket,host=127.0.0.1,port={self.port};urp;"],
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(60):
            try:
                self.ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={self.port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.5)
        else:
            raise SystemExit("could not start LibreOffice (soffice + python3-uno required)")
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

    def open(self, path):
        import uno
        from com.sun.star.beans import PropertyValue
        p = PropertyValue(); p.Name = "Hidden"; p.Value = True
        return self.desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(path)), "_blank", 0, (p,))

    def close(self):
        try:
            self.desktop.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=20)
        except Exception:
            self.proc.kill()
        shutil.rmtree(self.profile, ignore_errors=True)


def get_cell(doc, ref):
    book, sh, addr = split_sheet(ref.strip())
    return doc.Sheets.getByName(sh).getCellRangeByName(addr.replace("$", ""))


def read(cell):
    err = cell.getError()
    if err:
        return {"value": None, "display": cell.getString() or f"Err:{err}", "error": err}
    t = cell.getType().value
    if t == "FORMULA":
        s = cell.getString()
        rt = getattr(cell, "FormulaResultType2", None)
        if rt == 2:  # TEXT
            return {"value": s, "display": s, "error": 0}
        return {"value": cell.getValue(), "display": s, "error": 0}
    if t == "VALUE":
        return {"value": cell.getValue(), "display": cell.getString(), "error": 0}
    if t == "TEXT":
        return {"value": cell.getString(), "display": cell.getString(), "error": 0}
    return {"value": None, "display": "", "error": 0}


def write(cell, v):
    if v is None:
        cell.setFormula("")
    elif isinstance(v, bool):
        cell.setValue(1 if v else 0)
    elif isinstance(v, (int, float)):
        cell.setValue(float(v))
    elif isinstance(v, str) and v.startswith("="):
        raise SystemExit(f"formula inputs are not supported ({v}); edit formulas manually in a copy")
    else:
        cell.setString(str(v))


def validation_verdict(dv_rules, sheet, addr, value, doc):
    from openpyxl.utils.cell import range_boundaries, coordinate_from_string, column_index_from_string
    col, row = coordinate_from_string(addr.replace("$", ""))
    ci = column_index_from_string(col)
    for rule in dv_rules.get(sheet, []):
        for part in (rule.get("sqref") or "").split():
            try:
                c1, r1, c2, r2 = range_boundaries(part)
            except Exception:
                continue
            if c1 <= ci <= c2 and r1 <= row <= r2:
                t = rule.get("type")
                f1 = rule.get("formula1") or ""
                if t == "list":
                    if "(" in f1:
                        return f"list from formula `{f1[:60]}` — dependent/dynamic list, verify in Excel"
                    if f1.startswith('"'):
                        opts = [x.strip() for x in f1.strip('"').split(",")]
                    else:
                        try:
                            ref = f1.lstrip("=") if "!" in f1 else f"{sheet}!{f1.lstrip('=')}"
                            opts = [str(x).strip() for rowv in get_cell(doc, ref).getDataArray() for x in rowv if str(x).strip()]
                        except Exception:
                            return f"list rule ({f1}) — could not resolve list"
                    ok = value is None or str(value).strip() in opts or (isinstance(value, float) and str(int(value)) in opts)
                    return f"list `{f1[:40]}`: {'ACCEPTED' if ok else 'REJECTED by Excel'} ({len(opts)} options)"
                if t in ("decimal", "whole"):
                    if value is None:
                        return f"{t}: blank (allowBlank={rule.get('allowBlank')})"
                    if not isinstance(value, (int, float)):
                        return f"{t}: REJECTED by Excel (not a number)"
                    return f"{t} rule op={rule.get('operator') or 'between'} formula1={f1} — check bounds"
                return f"{t} rule present"
    return "no validation rule (any value accepted)"


def check_expect(name, exp, results):
    v = results.get(name, {}).get("value")
    tol = exp.get("tol", 1e-6)
    num = lambda x: isinstance(x, (int, float)) and not isinstance(x, bool)
    if "value" in exp:
        ok = (abs(v - exp["value"]) <= tol * max(1.0, abs(exp["value"]))) if num(v) and num(exp["value"]) else v == exp["value"]
        return ok, f"expected {exp['value']}, got {v}"
    if "same_as" in exp:
        o = results.get(exp["same_as"], {}).get("value")
        ok = (num(v) and num(o) and abs(v - o) <= tol * max(1.0, abs(o))) or v == o
        return ok, f"{v} vs {exp['same_as']}={o}"
    if "equals_product_of" in exp:
        prod = 1.0
        for r in exp["equals_product_of"]:
            x = results.get(r, {}).get("value")
            if not num(x):
                return False, f"factor {r} not numeric ({x})"
            prod *= x
        ok = num(v) and abs(v - prod) <= tol * max(1.0, abs(prod))
        return ok, f"{v} vs product {prod}"
    if "is_error" in exp:
        e = results.get(name, {}).get("error")
        return bool(e) == exp["is_error"], f"error={e}"
    return True, "no expectation"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--scenarios", required=True)
    ap.add_argument("--workbook", action="append", required=True)
    ap.add_argument("--password", action="append", default=[])
    ap.add_argument("--out", default="audit_out")
    ap.add_argument("--save-copies", action="store_true")
    a = ap.parse_args()
    out = ensure_dir(a.out)
    spec = json.load(open(a.scenarios, encoding="utf-8"))
    lov = lo_version()
    caveats = [f"LibreOffice {lov} used as calculation engine (not Excel)."]
    work = tempfile.mkdtemp(prefix="becc_recalc_")
    copies = []
    for i, wbp in enumerate(a.workbook):
        pw = a.password[i] if i < len(a.password) else None
        real = decrypt_if_needed(wbp, pw, work)
        dst = os.path.join(work, f"v{i + 1}_" + os.path.basename(wbp))
        shutil.copy2(real, dst)
        wi = workbook_xml_info(dst)
        dv = {s["name"]: sheet_xml_info(dst, s["xml"])["validations"] for s in wi["sheets"]}
        _, _, _, cl = load_cells(dst)
        used = {base_name(f) for c in cl if c.is_formula for f in functions_in_formula(c.formula)}
        for fn, first in LIBREOFFICE_FIRST.items():
            if fn in used and (first == "unsupported" or (lov and tuple(map(int, lov.split("."))) < tuple(map(int, first.split("."))))):
                caveats.append(f"{os.path.basename(wbp)} uses {fn}, unsupported in LibreOffice {lov} (needs {first}) — "
                               "those cells will error here but may work in Excel.")
        copies.append((wbp, dst, dv))
    lo = LO()
    rows, md = [], ["# Scenario recalculation results\n", "**Engine caveats:** " + " ".join(caveats) + "\n"]
    per_version = defaultdict(dict)
    try:
        for wbp, dst, dv in copies:
            for sc in spec["scenarios"]:
                doc = lo.open(dst)
                try:
                    verdicts = {}
                    for ref, v in sc.get("inputs", {}).items():
                        book, sh, addr = split_sheet(ref)
                        verdicts[ref] = validation_verdict(dv, sh, addr, v, doc)
                        write(get_cell(doc, ref), v)
                    doc.calculateAll()
                    exp = sc.get("expect", {})
                    wanted = list(dict.fromkeys(sc.get("outputs", []) + list(exp.keys())
                                                + [r for e in exp.values() for r in e.get("equals_product_of", [])]
                                                + [e["same_as"] for e in exp.values() if "same_as" in e]))
                    res = {}
                    for ref in wanted:
                        try:
                            res[ref] = read(get_cell(doc, ref))
                        except Exception as ex:
                            res[ref] = {"value": None, "display": f"READ FAIL {ex}", "error": -1}
                    per_version[sc["name"]][os.path.basename(wbp)] = res
                    md.append(f"\n## {sc['name']} — `{os.path.basename(wbp)}`\n")
                    if sc.get("description"):
                        md.append(sc["description"] + "\n")
                    md.append("| Input | Value | Data validation |\n|---|---|---|")
                    for ref, v in sc.get("inputs", {}).items():
                        md.append(f"| {ref} | {v!r} | {verdicts[ref]} |")
                    md.append("\n| Output | Value | Display | Check |\n|---|---|---|---|")
                    for ref in wanted:
                        r = res[ref]
                        chk = ""
                        if ref in exp:
                            ok, msg = check_expect(ref, exp[ref], res)
                            chk = ("PASS " if ok else "**FAIL** ") + msg
                        md.append(f"| {ref} | {short(r['value'], 40)} | {short(r['display'], 40)} | {chk} |")
                        rows.append({"workbook": os.path.basename(wbp), "scenario": sc["name"], "cell": ref,
                                     "value": r["value"], "display": r["display"], "error": r["error"], "check": chk})
                    if a.save_copies:
                        import uno
                        from com.sun.star.beans import PropertyValue
                        p = PropertyValue(); p.Name = "FilterName"; p.Value = "Calc MS Excel 2007 XML"
                        safe = re.sub(r"[^A-Za-z0-9]+", "_", sc["name"])[:40]
                        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(
                            os.path.join(out, f"recalc_{safe}_{os.path.basename(wbp)}"))), (p,))
                finally:
                    doc.close(True)
    finally:
        lo.close()
    if len(copies) > 1:
        md.append("\n## Cross-version reconciliation\n")
        names = [os.path.basename(c[0]) for c in copies]
        md.append("| Scenario | Cell | " + " | ".join(names) + " | Δ (last − first) |\n|---|---|" + "---|" * len(names) + "---|")
        for sc, byv in per_version.items():
            for ref in next(iter(byv.values())).keys():
                vals = [byv.get(n, {}).get(ref, {}).get("value") for n in names]
                delta = ""
                if all(isinstance(x, (int, float)) for x in vals):
                    delta = f"{vals[-1] - vals[0]:+.6g}" if abs(vals[-1] - vals[0]) > 1e-9 * max(1, abs(vals[0])) else "0"
                elif len(set(map(str, vals))) > 1:
                    delta = "differs"
                md.append(f"| {sc} | {ref} | " + " | ".join(short(v, 25) for v in vals) + f" | {delta} |")
    write_csv(os.path.join(out, "recalc_results.csv"), rows)
    with open(os.path.join(out, "recalc_results.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(md) + "\n")
    shutil.rmtree(work, ignore_errors=True)
    print(f"recalc → {out}/recalc_results.md")


if __name__ == "__main__":
    main()
