"""C752 Learner Guide: explicit rewrite of the WSQ LG (TGS-2023021099) as the non-WSQ twin."""
import shutil, docx
from docxutil import *

SRC = "source-wsq/courseware/LG-Carbon-Footprint-Management-for-Sustainability-v7.docx"
OUT = "courseware/LG-Sustainability Masterclass (C752).docx"
shutil.copy(SRC, OUT)
d = docx.Document(OUT)
retitle_all(d)

# Cover + version control record
set_text(para_by_text(d, "TGS Ref No:"), "Course Code: C752")
set_text(para_by_text(d, "Version v7"), "Version v1.0")
t0 = d.tables[0]
set_cell(t0.rows[1].cells[0], "v1.0")
set_cell(t0.rows[1].cells[1], "1 October 2026")
set_cell(t0.rows[1].cells[2], "Initial release of the Sustainability Masterclass learner guide: standards, Singapore "
         "context and nine real-use-case activities")
set_cell(t0.rows[1].cells[3], "Dr Alfred Ang")

# TOC
delete(para_by_text(d, "Assessment Preparation\t"))

# Course information
set_text(para_by_text(d, "Course code: TGS-2023021099"), "Course code: C752")
delete(para_by_text(d, "TSC: Carbon Footprint Management"))
set_text(para_by_text(d, "Duration: 16 training hours"),
         "Duration: 15 instructional hours over 2 days (9:30am – 5:30pm)")

# Assessment subsection (How to Use) and the Assessment Preparation section
delete_range(d, "Assessment", "Topic 1: Analyse Carbon Emission Hotspots")
delete_range(d, "Assessment Preparation", "Source Register")

# Competency codes -> the learning outcome each topic serves
for code, lo in (("K1 · K3 · K7 · A1", "LO1"), ("K4 · K5 · K8 · A2", "LO2"), ("K2 · K6 · K9 · A3", "LO3")):
    set_text(para_by_text(d, f"Competency alignment: {code}"), f"Learning outcome: {lo}")

# Footer
for s in d.sections:
    for part in (s.footer, s.header, s.first_page_footer, s.even_page_footer):
        for p in part.paragraphs:
            for r in p.runs:
                if "TGS-2023021099" in r.text:
                    r.text = r.text.replace("TGS-2023021099", "C752")
d.core_properties.title = "LG-Sustainability Masterclass (C752)"
d.save(OUT)
print("saved", OUT)
