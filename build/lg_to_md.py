"""Export the converted LG .docx to its Markdown mirror (headings, bullets, paragraphs, tables)."""
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

SRC = "courseware/LG-Improve Employee Wellness and Work Resilience.docx"
OUT = "LG-Improve Employee Wellness and Work Resilience (C1158).md"
doc = Document(SRC)
lines = ["# Learner Guide — Improve Employee Wellness and Work Resilience", "",
         "Course code: C1158 · Tertiary Infotech Academy Pte Ltd", ""]
started = False
for el in doc.element.body.iterchildren():
    tag = el.tag.split("}")[1]
    if tag == "p":
        p = Paragraph(el, doc)
        t = p.text.strip()
        st = p.style.name
        if st.startswith("Heading"):
            started = True                      # skip cover/TOC boilerplate before the first heading
            lvl = int(st.split()[-1]) if st.split()[-1].isdigit() else 1
            lines += ["#" * (lvl + 1) + " " + t, ""]
        elif started and t:
            lines += [("- " + t) if "List" in st else t, ""]
    elif tag == "tbl" and started:
        tb = Table(el, doc)
        rows = [[c.text.strip().replace("\n", " ").replace("|", "\\|") for c in r.cells] for r in tb.rows]
        lines.append("| " + " | ".join(rows[0]) + " |")
        lines.append("|" + "---|" * len(rows[0]))
        lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
        lines.append("")
open(OUT, "w").write("\n".join(lines).rstrip() + "\n")
print(OUT, len(lines))
