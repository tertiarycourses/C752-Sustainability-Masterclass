"""Write the LG Markdown mirror from the converted LG DOCX (body only: cover, version table and TOC skipped)."""
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
SRC = "courseware/LG-Sustainability Masterclass (C752).docx"
OUT = "LG-Sustainability Masterclass (C752).md"
d = docx.Document(SRC)
out = ["# Learner Guide — Sustainability Masterclass (C752)", "",
       "Tertiary Infotech Academy Pte Ltd · Version v1.0 · 1 October 2026", ""]
started = False
for el in d.element.body.iterchildren():
    tag = el.tag.split("}")[1]
    if tag == "p":
        p = Paragraph(el, d); t = p.text.strip(); st = p.style.name
        if st == "Heading 1": started = True
        if not started or not t: continue
        if st.startswith("Heading "):
            out += ["", "#" * (int(st[-1]) + 1) + " " + t, ""]
        elif "List" in st:
            out.append("- " + t)
        else:
            out += [t, ""]
    elif tag == "tbl" and started:
        rows = [[c.text.replace("\n", " ").replace("|", "\\|") for c in r.cells] for r in Table(el, d).rows]
        out.append("")
        out.append("| " + " | ".join(rows[0]) + " |")
        out.append("|" + "---|" * len(rows[0]))
        out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
        out.append("")
text = "\n".join(out)
while "\n\n\n" in text: text = text.replace("\n\n\n", "\n\n")
open(OUT, "w").write(text.rstrip() + "\n")
print(OUT, len(out))
