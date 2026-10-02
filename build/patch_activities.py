"""Retitle/recode the activities tree for C1158 (Markdown + in-place ReportLab PDF patch).

The PDFs have no build script in the WSQ repo, so their uncompressed Tj strings are
rewritten in place (pypdf clone_from + replace_contents). Every font that draws a
rewritten string is checked for the glyphs the new text needs.
"""
import io
import re
from pathlib import Path

import pypdf
from fontTools.ttLib import TTFont

ROOT = Path("activities")
KA = r"(?:[KA]\d (?:\||\\267) )+"

MD_SUBS = [
    (re.compile(r"^# Activities — Empowering Employee Health and Wellness at the Workplace$", re.M),
     "# Activities — Improve Employee Wellness and Work Resilience"),
    (re.compile(r"^Course: TGS-2024045222$", re.M), "Course: C1158"),
    (re.compile(r" — (?:[KA]\d · )+(LO\d)$", re.M), r" — \1"),
    (re.compile(r"^\*\*TSC mapping:\*\* (?:[KA]\d · )+(LO\d)", re.M), r"**Learning outcome:** \1"),
]

PDF_SUBS = [
    (rb"\(TGS-2024045222 \| Activity Resource\)", b"(C1158 | Activity Resource)"),
    (rb"\(Activities - Empowering Employee Health and\) Tj T\* \(Wellness at the Workplace\)",
     b"(Activities - Improve Employee Wellness) Tj T* (and Work Resilience)"),
    (rb"\(Course: TGS-2024045222\)", b"(Course: C1158)"),
    (rb"\(Course TGS-2024045222 \\267 " + KA.encode() + rb"(LO\d)\)", rb"(Course C1158 \\267 \1)"),
    (rb"\( - " + KA.encode() + rb"(LO\d)\)", rb"( - \1)"),
    (rb"\(TSC mapping: " + KA.encode() + rb"(LO\d) Objective:", rb"(Learning outcome: \1 Objective:"),
]
NEW_CHARS = set("C1158 | Activity Resource Improve Employee Wellness and Work Resilience Course: Learning outcome:")


def font_chars(page, name):
    """Characters the font can draw (Type1 standard fonts: all). ReportLab TTF subsets
    carry a (1,0) cmap whose codes equal the ASCII bytes in the Tj strings."""
    fobj = page["/Resources"]["/Font"][name].get_object()
    if fobj.get("/Subtype") == "/Type1":
        return None
    ff = fobj["/FontDescriptor"].get_object()["/FontFile2"].get_object().get_data()
    tt = TTFont(io.BytesIO(ff))
    return {chr(c) for t in tt["cmap"].tables for c, g in t.cmap.items() if g != ".notdef"}


def check_glyphs(page, data):
    """For every rewritten string, the font active at that point must hold its glyphs."""
    font = None
    for m in re.finditer(rb"/(\S+) [\d.]+ Tf|\(((?:[^()\\]|\\.)*)\) Tj", data):
        if m.group(1):
            font = "/" + m.group(1).decode()
            continue
        s = m.group(2).decode("latin-1")
        if not any(k in s for k in ("C1158", "Improve Employee", "and Work Resilience", "Learning outcome")):
            continue
        have = font_chars(page, font)
        if have is None:
            continue
        missing = {ch for ch in s if ch not in have and ch in NEW_CHARS and ch != " "}
        assert not missing, (font, s, missing)


for md in sorted(ROOT.rglob("*.md")):
    t = md.read_text()
    new = t
    for pat, rep in MD_SUBS:
        new = pat.sub(rep, new)
    if new != t:
        md.write_text(new)

changed = 0
for pdf in sorted(ROOT.rglob("*.pdf")):
    w = pypdf.PdfWriter(clone_from=str(pdf))
    hit = False
    for pg in w.pages:
        co = pg.get_contents()
        d = co.get_data()
        n = d
        for pat, rep in PDF_SUBS:
            n = re.sub(pat, rep, n)
        if n != d:
            check_glyphs(pg, n)
            co.set_data(n)
            pg.replace_contents(co)
            hit = True
    if hit:
        with open(pdf, "wb") as fh:
            w.write(fh)
        changed += 1
print("pdfs patched:", changed)
