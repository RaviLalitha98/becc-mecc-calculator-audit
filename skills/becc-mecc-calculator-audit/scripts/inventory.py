#!/usr/bin/env python3
"""Workbook structure inventory (audit stages 1, 15, 23, 24).

Produces <out>/inventory.json and <out>/inventory.md describing:
  * file format, application/version that last saved it, date system,
    calculation settings (mode, iterative calc, precision-as-displayed)
  * every sheet: visibility (visible / hidden / veryHidden), protection,
    used range, formula/constant/array counts, hidden rows & columns,
    data-validation rules (incl. x14 extension rules openpyxl drops),
    conditional-format count, tables
  * defined names (scope, hidden, target, broken #REF!)
  * external links with their target paths, VBA, Power Query, connections,
    pivot caches, charts
  * sheet-to-sheet dependency matrix and sheets nothing points to
  * Excel functions used, `_xlfn.` usage and an Excel-version compatibility
    matrix built from the functions actually present

Usage: python inventory.py WORKBOOK [--password PW] [--out DIR]
"""
from __future__ import annotations

import argparse
import os
import posixpath
import re
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict

from common import (not_excel, load_cells, index_cells, refs_in_formula, functions_in_formula,
                    read_zip_text, zip_members, write_json, ensure_dir, short)
from excel_compat import (FIRST_VERSION, SPECIAL_XLFN, VERSION_RANK, VOLATILE,
                          LIBREOFFICE_FIRST, base_name)

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}


def _rels(path, member):
    txt = read_zip_text(path, member)
    if not txt:
        return {}
    root = ET.fromstring(txt)
    return {r.get("Id"): (r.get("Target"), r.get("TargetMode"), r.get("Type"))
            for r in root.findall("rel:Relationship", NS)}


def workbook_xml_info(path):
    info = {}
    txt = read_zip_text(path, "xl/workbook.xml")
    root = ET.fromstring(txt)
    fv = root.find("m:fileVersion", NS)
    info["fileVersion"] = dict(fv.attrib) if fv is not None else {}
    wp = root.find("m:workbookPr", NS)
    info["workbookPr"] = dict(wp.attrib) if wp is not None else {}
    info["date_system"] = "1904" if (wp is not None and wp.get("date1904") in ("1", "true")) else "1900"
    prot = root.find("m:workbookProtection", NS)
    info["workbookProtection"] = ({k: v for k, v in prot.attrib.items()
                                   if not any(s in k.lower() for s in ("hash", "salt", "password", "spincount"))}
                                  or {"present": True}) if prot is not None else None
    cp = root.find("m:calcPr", NS)
    info["calcPr"] = dict(cp.attrib) if cp is not None else {}
    rels = _rels(path, "xl/_rels/workbook.xml.rels")
    sheets = []
    for s in root.find("m:sheets", NS):
        rid = s.get(f"{{{NS['r']}}}id")
        target = rels.get(rid, (None,))[0]
        if target:
            target = target.lstrip("/")
            if not target.startswith("xl/"):
                target = posixpath.normpath(posixpath.join("xl", target))
        sheets.append({"name": s.get("name"), "state": s.get("state", "visible"), "xml": target})
    info["sheets"] = sheets
    names = []
    dn = root.find("m:definedNames", NS)
    if dn is not None:
        for d in dn:
            names.append({"name": d.get("name"), "localSheetId": d.get("localSheetId"),
                          "hidden": d.get("hidden") in ("1", "true"), "refers_to": d.text or ""})
    info["defined_names"] = names
    app = read_zip_text(path, "docProps/app.xml") or ""
    m1 = re.search(r"<Application>(.*?)</Application>", app)
    m2 = re.search(r"<AppVersion>(.*?)</AppVersion>", app)
    info["application"] = m1.group(1) if m1 else None
    info["app_version"] = m2.group(1) if m2 else None
    core = read_zip_text(path, "docProps/core.xml") or ""
    m3 = re.search(r"<dcterms:modified[^>]*>(.*?)</dcterms:modified>", core)
    m4 = re.search(r"<cp:lastModifiedBy>(.*?)</cp:lastModifiedBy>", core)
    info["modified"] = m3.group(1) if m3 else None
    info["last_modified_by_present"] = bool(m4)
    return info


def _unesc(s):
    return (s or "").replace("&quot;", '"').replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&apos;", "'")


def sheet_xml_info(path, xml_member):
    """Raw-XML facts openpyxl does not expose fully."""
    out = {"hidden_rows": [], "hidden_cols": [], "validations": [], "protection": None,
           "cond_formats": 0, "tables": [], "outline_rows": 0, "dynamic_cells": 0}
    txt = read_zip_text(path, xml_member) if xml_member else None
    if not txt:
        return out
    out["hidden_rows"] = [int(m.group(1)) for m in re.finditer(r'<row [^>]*?r="(\d+)"[^>]*?hidden="(?:1|true)"', txt)]
    for m in re.finditer(r'<col [^>]*?/>', txt):
        tag = m.group(0)
        if re.search(r'hidden="(?:1|true)"', tag):
            mn = int(re.search(r'min="(\d+)"', tag).group(1))
            mx = int(re.search(r'max="(\d+)"', tag).group(1))
            out["hidden_cols"].append((mn, mx))
    out["outline_rows"] = len(re.findall(r'<row [^>]*outlineLevel="', txt))
    out["dynamic_cells"] = len(re.findall(r'<c [^>]*\bcm="\d+"', txt))
    sp = re.search(r"<sheetProtection ([^>]*)/>", txt)
    if sp:
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', sp.group(1)))
        out["protection"] = {k: v for k, v in attrs.items()
                             if not any(s in k.lower() for s in ("hash", "salt", "password", "spincount"))}
    out["cond_formats"] = len(re.findall(r"<conditionalFormatting", txt)) + len(re.findall(r"<x14:conditionalFormatting", txt))
    for m in re.finditer(r"<dataValidation ([^>]*)>(.*?)</dataValidation>", txt, re.S):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        f1 = re.search(r"<formula1>(.*?)</formula1>", m.group(2), re.S)
        out["validations"].append({"sqref": attrs.get("sqref"), "type": attrs.get("type", "any"),
                                   "operator": attrs.get("operator"), "allowBlank": attrs.get("allowBlank"),
                                   "showErrorMessage": attrs.get("showErrorMessage"),
                                   "errorStyle": attrs.get("errorStyle", "stop"),
                                   "formula1": _unesc(f1.group(1)) if f1 else None, "ext": False})
    for m in re.finditer(r"<x14:dataValidation ([^>]*)>(.*?)</x14:dataValidation>", txt, re.S):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        f1 = re.search(r"<x14:formula1>\s*<xm:f>(.*?)</xm:f>", m.group(2), re.S)
        sq = re.search(r"<xm:sqref>(.*?)</xm:sqref>", m.group(2), re.S)
        out["validations"].append({"sqref": sq.group(1) if sq else None, "type": attrs.get("type", "any"),
                                   "operator": attrs.get("operator"), "allowBlank": attrs.get("allowBlank"),
                                   "showErrorMessage": attrs.get("showErrorMessage"),
                                   "errorStyle": attrs.get("errorStyle", "stop"),
                                   "formula1": _unesc(f1.group(1)) if f1 else None, "ext": True})
    rels_member = posixpath.join(posixpath.dirname(xml_member), "_rels", posixpath.basename(xml_member) + ".rels")
    for rid, (target, mode, typ) in _rels(path, rels_member).items():
        if typ and typ.endswith("/table"):
            tpath = posixpath.normpath(posixpath.join(posixpath.dirname(xml_member), target))
            ttxt = read_zip_text(path, tpath) or ""
            name = re.search(r'<table [^>]*\bname="([^"]*)"', ttxt)
            ref = re.search(r'<table [^>]*\bref="([^"]*)"', ttxt)
            out["tables"].append({"xml": tpath, "name": name.group(1) if name else None,
                                  "ref": ref.group(1) if ref else None,
                                  "calculated_columns": len(re.findall(r"<calculatedColumnFormula", ttxt)),
                                  "totals_row": 'totalsRowCount="1"' in ttxt})
    return out


def external_links(path):
    out = []
    for m in zip_members(path):
        mm = re.match(r"xl/externalLinks/(externalLink\d+)\.xml$", m)
        if not mm:
            continue
        rels = _rels(path, f"xl/externalLinks/_rels/{mm.group(1)}.xml.rels")
        txt = read_zip_text(path, m) or ""
        sheets = [_unesc(s) for s in re.findall(r'<sheetName val="([^"]*)"', txt)]
        cached = len(re.findall(r"<cell ", txt))
        for rid, (target, mode, typ) in rels.items():
            out.append({"part": m, "index": int(re.search(r"\d+", mm.group(1)).group(0)),
                        "target": target, "mode": mode, "sheets": sheets, "cached_cells": cached})
    return out


def package_features(path):
    mem = zip_members(path)
    feats = {
        "vba_project": any(m.lower().endswith("vbaproject.bin") for m in mem),
        "connections": "xl/connections.xml" in mem,
        "query_tables": [m for m in mem if m.startswith("xl/queryTables/")],
        "pivot_caches": [m for m in mem if m.startswith("xl/pivotCache/") and m.endswith(".xml") and "Definition" in m],
        "power_query": False,
        "activex_or_controls": [m for m in mem if m.startswith("xl/activeX/") or m.startswith("xl/ctrlProps/")],
        "charts": len([m for m in mem if re.match(r"xl/charts/chart\d+\.xml$", m)]),
        "chartex": len([m for m in mem if re.match(r"xl/charts/chartEx\d+\.xml$", m)]),
        "drawings": len([m for m in mem if re.match(r"xl/drawings/drawing\d+\.xml$", m)]),
        "comments": len([m for m in mem if re.match(r"xl/comments\d*\.xml", m) or m.startswith("xl/threadedComments/")]),
        "metadata_xml": "xl/metadata.xml" in mem,
    }
    for m in mem:
        if m.startswith("customXml/item") and m.endswith(".xml"):
            t = read_zip_text(path, m) or ""
            if "DataMashup" in t:
                feats["power_query"] = True
    if feats["connections"]:
        t = read_zip_text(path, "xl/connections.xml") or ""
        feats["connection_list"] = [{"name": n, "type": ty} for n, ty in
                                    re.findall(r'<connection [^>]*name="([^"]*)"[^>]*type="(\d+)"', t)]
    return feats


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook")
    ap.add_argument("--password")
    ap.add_argument("--out", default="audit_out")
    a = ap.parse_args()
    out = ensure_dir(a.out)

    wb_f, wb_v, real, cells = load_cells(a.workbook, a.password)
    winfo = workbook_xml_info(real)
    ext = external_links(real)
    feats = package_features(real)

    sheet_names = [s["name"] for s in winfo["sheets"]]
    sheets = []
    dep = defaultdict(Counter)
    fn_counter = Counter()
    fn_cells = defaultdict(list)
    xlfn_raw = Counter()
    ext_ref_cells = []
    volatile_cells = defaultdict(list)
    for c in cells:
        if not c.is_formula:
            continue
        for fn in functions_in_formula(c.formula):
            b = base_name(fn)
            fn_counter[b] += 1
            if len(fn_cells[b]) < 5:
                fn_cells[b].append(c.full)
            if fn.startswith("_XLFN") or fn.startswith("_XLWS"):
                xlfn_raw[fn] += 1
            if b in VOLATILE and len(volatile_cells[b]) < 10:
                volatile_cells[b].append(c.full)
        for txt in re.findall(r"_xlfn\.(SINGLE|ANCHORARRAY)", c.formula, re.I):
            xlfn_raw["_XLFN." + txt.upper()] += 1
        for op, book, sh, kind, data in refs_in_formula(c.formula):
            if book:
                ext_ref_cells.append({"cell": c.full, "formula": short(c.formula, 200)})
                dep[c.sheet]["[external " + book + "]" + (sh or "")] += 1
            elif sh is not None:
                dep[c.sheet][sh] += 1
            else:
                dep[c.sheet][c.sheet] += 1

    name_targets = Counter()
    broken_names = []
    ext_names = []
    for n in winfo["defined_names"]:
        rt = n["refers_to"]
        if "#REF!" in rt:
            broken_names.append(n)
        if re.search(r"\[\d+\]", rt):
            ext_names.append(n)
        for sh in sheet_names:
            if f"'{sh}'!" in rt or f"{sh}!" in rt:
                name_targets[sh] += 1

    for s in winfo["sheets"]:
        name = s["name"]
        ws = wb_f[name]
        x = sheet_xml_info(real, s["xml"])
        sc = [c for c in cells if c.sheet == name]
        kinds = Counter(c.kind for c in sc)
        for v in x["validations"]:
            f1 = v.get("formula1") or ""
            for sh in sheet_names:
                if f"'{sh}'!" in f1 or f"{sh}!" in f1:
                    name_targets[sh] += 1
        incoming = sum(cnt for frm, tos in dep.items() if frm != name for to, cnt in tos.items() if to == name)
        sheets.append({
            "name": name, "state": s["state"], "dimensions": ws.dimensions,
            "formulas": kinds.get("formula", 0), "array_formulas": kinds.get("array", 0),
            "data_tables": kinds.get("datatable", 0), "dynamic_array_cells": x["dynamic_cells"],
            "numbers": kinds.get("number", 0), "text": kinds.get("text", 0),
            "error_constants": kinds.get("error_const", 0),
            "cached_errors": sum(1 for c in sc if c.is_formula and isinstance(c.value, str) and c.value.startswith("#")),
            "protected": x["protection"] is not None and x["protection"].get("sheet", "1") in ("1", "true"),
            "protection_flags": x["protection"],
            "hidden_rows": len(x["hidden_rows"]), "hidden_rows_sample": x["hidden_rows"][:30],
            "hidden_cols": x["hidden_cols"],
            "validations": x["validations"], "conditional_formats": x["cond_formats"],
            "tables": x["tables"], "merged_ranges": len(ws.merged_cells.ranges),
            "incoming_refs": incoming, "incoming_name_or_validation_refs": name_targets.get(name, 0),
            "outgoing_sheets": sorted(k for k in dep.get(name, {}) if k != name),
        })

    used_versioned = {fn: FIRST_VERSION[fn] for fn in fn_counter if fn in FIRST_VERSION}
    matrix = []
    for fn, first in sorted(used_versioned.items(), key=lambda kv: VERSION_RANK[kv[1]]):
        row = {"feature": fn, "first_perpetual_version": first, "uses": fn_counter[fn],
               "example_cells": fn_cells[fn][:3], "libreoffice_first": LIBREOFFICE_FIRST.get(fn, "supported")}
        for v in ["2016", "2019", "2021", "2024", "M365"]:
            row[f"Excel {v}"] = "Yes" if VERSION_RANK[v] >= VERSION_RANK[first] else "NO (#NAME?)"
        matrix.append(row)
    min_version = max(used_versioned.values(), key=lambda v: VERSION_RANK[v]) if used_versioned else "pre-2010"

    result = {
        "workbook": os.path.abspath(a.workbook),
        "format": os.path.splitext(a.workbook)[1].lower(),
        "application": winfo["application"], "app_version": winfo["app_version"], "modified": winfo["modified"],
        "fileVersion": winfo["fileVersion"], "date_system": winfo["date_system"],
        "calcPr": winfo["calcPr"], "workbookProtection": winfo["workbookProtection"],
        "sheets": sheets, "defined_names": winfo["defined_names"], "broken_names": broken_names,
        "external_names": ext_names,
        "external_links": ext, "external_ref_cells": ext_ref_cells[:200],
        "external_ref_cell_count": len(ext_ref_cells),
        "package_features": feats,
        "functions_used": dict(fn_counter.most_common()),
        "xlfn_prefixed": dict(xlfn_raw), "volatile_functions": dict(volatile_cells),
        "compatibility_matrix": matrix, "minimum_excel_version_by_functions": min_version,
        "special_xlfn": {k: SPECIAL_XLFN[k.split(".")[-1]] for k in xlfn_raw if k.split(".")[-1] in SPECIAL_XLFN},
        "sheet_dependencies": {k: dict(v) for k, v in dep.items()},
    }
    write_json(os.path.join(out, "inventory.json"), result)

    L = [f"# Workbook inventory — {os.path.basename(a.workbook)}\n"]
    app = winfo["application"] or "?"
    L.append(f"- Format: `{result['format']}`; last saved by **{app}** {winfo['app_version'] or ''}; modified {winfo['modified']}; "
             f"fileVersion {winfo['fileVersion']}")
    if not_excel(app):
        L.append(f"  - ⚠ Not saved by Excel. Non-Excel writers (e.g. openpyxl) can drop x14 data validations, charts, "
                 f"dynamic-array metadata and cached values — compare with the previous Excel-saved version.")
    cp = winfo["calcPr"]
    L.append(f"- Calculation: mode=`{cp.get('calcMode', 'auto (default)')}`, fullCalcOnLoad={cp.get('fullCalcOnLoad', '0')}, "
             f"iterate={cp.get('iterate', '0')} (count {cp.get('iterateCount', '100')}, delta {cp.get('iterateDelta', '0.001')}), "
             f"fullPrecision={cp.get('fullPrecision', '1 (default)')}, calcId={cp.get('calcId')}")
    L.append(f"- Date system: {winfo['date_system']}")
    L.append(f"- Workbook structure protection: {winfo['workbookProtection']}")
    L.append(f"- VBA: {feats['vba_project']}; Power Query: {feats['power_query']}; connections: {feats['connections']}; "
             f"pivot caches: {len(feats['pivot_caches'])}; ActiveX/form controls: {len(feats['activex_or_controls'])}; "
             f"charts: {feats['charts']} (+{feats['chartex']} chartEx); dynamic-array metadata: {feats['metadata_xml']}")
    L.append(f"- External links: {len(ext)} link relationship(s); {len(ext_ref_cells)} formula cell(s) and "
             f"{len(ext_names)} defined name(s) reference external books")
    for e in ext:
        L.append(f"  - [{e['index']}] → `{e['target']}` ({e['mode']}), cached cells {e['cached_cells']}, sheets: {', '.join(e['sheets'][:15])}")
    L.append(f"- Minimum Excel version implied by functions: **{min_version}**")
    L.append("\n## Sheets\n")
    L.append("| # | Sheet | State | Prot | Used range | Formulas | Arrays | Dyn | Numbers | Cached errors | Hidden rows | Hidden cols | DV rules | In-refs | Out-sheets |")
    L.append("|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|")
    for i, s in enumerate(sheets, 1):
        L.append(f"| {i} | {s['name']} | {s['state']} | {'Y' if s['protected'] else ''} | {s['dimensions']} | {s['formulas']} | "
                 f"{s['array_formulas']} | {s['dynamic_array_cells']} | {s['numbers']} | {s['cached_errors']} | {s['hidden_rows']} | "
                 f"{len(s['hidden_cols'])} | {len(s['validations'])} | {s['incoming_refs']}+{s['incoming_name_or_validation_refs']} | "
                 f"{', '.join(s['outgoing_sheets'][:6])}{'…' if len(s['outgoing_sheets']) > 6 else ''} |")
    orphans = [s["name"] for s in sheets if s["incoming_refs"] == 0 and s["incoming_name_or_validation_refs"] == 0]
    L.append(f"\nSheets with no incoming formula/name/validation references (inputs, outputs, docs — or orphans): {orphans}")
    L.append("\n## Defined names\n")
    for n in winfo["defined_names"]:
        if n["name"].startswith("_xlchart"):
            continue
        scope = "workbook" if n["localSheetId"] is None else f"sheet #{n['localSheetId']}"
        flag = " **BROKEN**" if "#REF!" in n["refers_to"] else (" **EXTERNAL**" if n in ext_names else "")
        L.append(f"- `{n['name']}` ({scope}{', hidden' if n['hidden'] else ''}) = `{short(n['refers_to'], 120)}`{flag}")
    L.append("\n## Function usage\n")
    L.append(", ".join(f"{k}×{v}" for k, v in fn_counter.most_common()))
    if xlfn_raw:
        L.append(f"\n`_xlfn`/`_xlws` prefixed: {dict(xlfn_raw)}")
    if volatile_cells:
        L.append(f"\nVolatile functions: { {k: len(v) for k, v in volatile_cells.items()} }")
    L.append("\n## Excel version compatibility (functions actually used)\n")
    if matrix:
        L.append("| Function | First perpetual | Uses | 2016 | 2019 | 2021 | 2024 | M365 | LibreOffice | Example |")
        L.append("|---|---|---|---|---|---|---|---|---|---|")
        for r in matrix:
            L.append(f"| {r['feature']} | {r['first_perpetual_version']} | {r['uses']} | {r['Excel 2016']} | {r['Excel 2019']} | "
                     f"{r['Excel 2021']} | {r['Excel 2024']} | {r['Excel M365']} | {r['libreoffice_first']} | {r['example_cells'][0]} |")
    else:
        L.append("No version-gated functions detected.")
    L.append("\n## Data validation summary\n")
    for s in sheets:
        if s["validations"]:
            types = Counter(v["type"] for v in s["validations"])
            nostop = sum(1 for v in s["validations"] if v.get("errorStyle") in ("warning", "information")
                         or v.get("showErrorMessage") in (None, "0"))
            dyn = sum(1 for v in s["validations"] if v.get("formula1") and "(" in v["formula1"])
            L.append(f"- **{s['name']}**: {len(s['validations'])} rules {dict(types)}; {nostop} not enforcing a Stop error; "
                     f"{dyn} list sources built from formulas")
    L.append("\n## Sheet dependency edges (from → to: formula refs)\n")
    for frm, tos in result["sheet_dependencies"].items():
        others = {k: v for k, v in tos.items() if k != frm and v}
        if others:
            L.append(f"- {frm} → " + ", ".join(f"{k} ({v})" for k, v in sorted(others.items(), key=lambda kv: -kv[1])))
    with open(os.path.join(out, "inventory.md"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(L) + "\n")
    print(f"inventory written to {out}/inventory.md and inventory.json")


if __name__ == "__main__":
    main()
