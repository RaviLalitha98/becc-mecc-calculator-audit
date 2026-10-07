#!/usr/bin/env python3
"""Export the Markdown evaluation report to Word (.docx) and/or PDF.

The report is always written first as Markdown (assets/report_template.md);
this script turns that file into the other formats the user asked for.
Pages are A4 landscape so the wide findings tables stay readable.

Requirements:
  docx  python-docx               pip install python-docx
  pdf   markdown + xhtml2pdf      pip install markdown xhtml2pdf

Usage:
  python export_report.py audit_report.md --format docx pdf --out deliverables
  python export_report.py audit_report.md --format all
"""
from __future__ import annotations

import argparse
import os
import re
import sys

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(errors="replace")

INLINE = re.compile(r"(\*\*[^*]+\*\*|`[^`]+`|\*[^*\s][^*]*\*)")
TABLE_SEP = re.compile(r"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$")


CODE = re.compile(r"(```.*?```|`[^`\n]+`)", re.S)
FORMULA_STAR = re.compile(r"(?<=[\w)\]\"'])\*(?=[\w(\[\"'$])")


def protect_formula_stars(md: str) -> str:
    """Escape '*' used as multiplication (=I*M*K) so it is not read as italics; code is left alone."""
    return "".join(part if i % 2 else FORMULA_STAR.sub(r"\\*", part)
                   for i, part in enumerate(CODE.split(md)))


def split_row(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [c.strip() for c in re.split(r"(?<!\\)\|", line)]


def add_inline(par, text: str):
    """Add text to a python-docx paragraph, honouring **bold**, *italic* and `code`."""
    text = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1 (\2)", text).replace("\\*", "\x00")
    for part in INLINE.split(text):
        if not part:
            continue
        part = part.replace("\x00", "*")
        if part.startswith("**") and part.endswith("**"):
            par.add_run(part[2:-2]).bold = True
        elif part.startswith("`") and part.endswith("`"):
            run = par.add_run(part[1:-1])
            run.font.name = "Consolas"
        elif part.startswith("*") and part.endswith("*") and len(part) > 2:
            par.add_run(part[1:-1]).italic = True
        else:
            par.add_run(part)


def to_docx(md: str, path: str):
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.shared import Cm, Pt

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(1.8))
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)

    lines = protect_formula_stars(md).splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        s = line.strip()
        if not s:
            i += 1
            continue
        if s.startswith("```"):  # fenced code block
            i += 1
            code = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code.append(lines[i])
                i += 1
            p = doc.add_paragraph()
            run = p.add_run("\n".join(code))
            run.font.name = "Consolas"
            run.font.size = Pt(9)
            i += 1
            continue
        m = re.match(r"^(#{1,6})\s+(.*)", s)
        if m:
            # "#" becomes the document title (level 0); "##" Heading 1, "###" Heading 2, ...
            doc.add_heading(re.sub(r"[*`]", "", m.group(2)), level=min(len(m.group(1)) - 1, 4))
            i += 1
            continue
        if re.match(r"^(-{3,}|\*{3,}|_{3,})$", s):
            i += 1
            continue
        if s.startswith("|") and i + 1 < len(lines) and TABLE_SEP.match(lines[i + 1].strip()):
            header = split_row(s)
            rows = []
            i += 2
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(split_row(lines[i]))
                i += 1
            table = doc.add_table(rows=1, cols=len(header))
            table.style = "Table Grid"
            for c, h in enumerate(header):
                cell = table.rows[0].cells[c]
                cell.text = ""
                add_inline(cell.paragraphs[0], f"**{h}**" if h else "")
            for r in rows:
                cells = table.add_row().cells
                for c in range(len(header)):
                    add_inline(cells[c].paragraphs[0], r[c] if c < len(r) else "")
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.font.size = Pt(8)
            doc.add_paragraph()
            continue
        m = re.match(r"^[-*+]\s+(.*)", s)
        if m:
            add_inline(doc.add_paragraph(style="List Bullet"), m.group(1))
            i += 1
            continue
        m = re.match(r"^\d+[.)]\s+(.*)", s)
        if m:
            add_inline(doc.add_paragraph(style="List Number"), m.group(1))
            i += 1
            continue
        # plain paragraph: join following non-special lines
        para = [s]
        i += 1
        while i < len(lines):
            n = lines[i].strip()
            if not n or re.match(r"^(#|\||```|[-*+]\s|\d+[.)]\s|-{3,})", n):
                break
            para.append(n)
            i += 1
        add_inline(doc.add_paragraph(), " ".join(para))
    doc.save(path)


PDF_CSS = """
@page { size: a4 landscape; margin: 1.6cm; }
body { font-family: Helvetica; font-size: 9.5pt; line-height: 1.35; color: #222; }
h1 { font-size: 18pt; color: #1F3864; border-bottom: 1px solid #1F3864; padding-bottom: 4px; }
h2 { font-size: 13pt; color: #1F3864; margin-top: 14px; }
h3 { font-size: 11pt; color: #2F5496; }
table { border: 0.5px solid #999; margin: 6px 0 10px 0; }
th { background-color: #D9E2F3; font-weight: bold; padding: 3px; font-size: 8pt; text-align: left; }
td { padding: 3px; font-size: 8pt; vertical-align: top; }
code, pre { font-family: Courier; font-size: 8.5pt; background-color: #F2F2F2; }
pre { padding: 6px; }
"""


# The built-in PDF fonts cannot show →, −, ≥, ³ etc. Use a system TrueType font when one exists.
FONT_CANDIDATES = [  # (regular, bold)
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ("/usr/share/fonts/dejavu/DejaVuSans.ttf", "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
    ("/Library/Fonts/Arial Unicode.ttf", "/Library/Fonts/Arial Unicode.ttf"),
    ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
]


def register_font() -> str:
    """Register a Unicode TrueType font with xhtml2pdf; return the CSS that selects it."""
    from reportlab.lib.fonts import addMapping
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from xhtml2pdf.default import DEFAULT_FONT

    for regular, bold in FONT_CANDIDATES:
        if os.path.isfile(regular):
            bold = bold if os.path.isfile(bold) else regular
            pdfmetrics.registerFont(TTFont("Report", regular))
            pdfmetrics.registerFont(TTFont("Report-Bold", bold))
            addMapping("Report", 0, 0, "Report")
            addMapping("Report", 1, 0, "Report-Bold")
            addMapping("Report", 0, 1, "Report")
            addMapping("Report", 1, 1, "Report-Bold")
            DEFAULT_FONT["report"] = "Report"
            return "body, th, td { font-family: Report; }"
    print("No Unicode TrueType font found; symbols such as → may not display in the PDF.", file=sys.stderr)
    return ""


def size_columns(html: str) -> str:
    """Give each table column a width in proportion to its text, so long text columns are not squeezed."""
    def one(m):
        table = m.group(0)
        rows = [re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", r, re.S) for r in re.findall(r"<tr>(.*?)</tr>", table, re.S)]
        ncol = max((len(r) for r in rows), default=0)
        if ncol < 2:
            return table
        lengths = [0.0] * ncol
        for row in rows:
            for j, c in enumerate(row[:ncol]):
                lengths[j] += len(re.sub(r"<[^>]+>", "", c))
        header = [len(re.sub(r"<[^>]+>", "", c)) for c in rows[0]] + [0] * ncol
        weights = [min(max(x / len(rows), header[j] + 2, 6), 80) for j, x in enumerate(lengths)]
        widths = iter(f"{100 * w / sum(weights):.0f}%" for w in weights)
        return re.sub(r"<th(?=[\s>])([^>]*)>", lambda h: f'<th{h.group(1)} width="{next(widths, "")}">', table)
    return re.sub(r"<table>.*?</table>", one, html, flags=re.S)


def to_pdf(md: str, path: str):
    import markdown
    from xhtml2pdf import pisa

    body = markdown.markdown(protect_formula_stars(md), extensions=["tables", "fenced_code", "sane_lists"])
    body = size_columns(body)
    # Long formulas have no spaces to wrap at and would run off the page; the PDF engine only
    # breaks at real spaces, so add one after argument commas (Excel ignores these spaces).
    body = re.sub(r"[^\s<>]{60,}", lambda m: re.sub(r",(?=\S)", ", ", m.group(0)), body)
    css = PDF_CSS + register_font()
    html = f"<html><head><meta charset='utf-8'><style>{css}</style></head><body>{body}</body></html>"
    with open(path, "wb") as fh:
        result = pisa.CreatePDF(html, dest=fh, encoding="utf-8")
    if result.err:
        raise RuntimeError(f"PDF conversion reported {result.err} error(s)")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("report", help="Markdown report to export")
    ap.add_argument("--format", nargs="+", default=["all"], choices=["docx", "pdf", "all"])
    ap.add_argument("--out", default=None, help="output folder (default: next to the report)")
    a = ap.parse_args()

    md = open(a.report, encoding="utf-8").read()
    out = a.out or os.path.dirname(os.path.abspath(a.report))
    os.makedirs(out, exist_ok=True)
    stem = os.path.splitext(os.path.basename(a.report))[0]
    formats = ["docx", "pdf"] if "all" in a.format else a.format

    failed = False
    for fmt in formats:
        path = os.path.join(out, f"{stem}.{fmt}")
        try:
            (to_docx if fmt == "docx" else to_pdf)(md, path)
            print(f"Wrote {path}")
        except ImportError as e:
            pkg = "python-docx" if fmt == "docx" else "markdown xhtml2pdf"
            print(f"Cannot write {fmt}: {e}. Install with: pip install {pkg}", file=sys.stderr)
            failed = True
        except Exception as e:  # keep going so the other format can still be produced
            print(f"Failed to write {fmt}: {e}", file=sys.stderr)
            failed = True
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
