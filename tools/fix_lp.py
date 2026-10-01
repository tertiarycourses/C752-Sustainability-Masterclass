"""C752 Lesson Plan: explicit rewrite of the WSQ LP (TGS-2023021099) as the non-WSQ twin."""
import re, shutil, docx
from docxutil import *

SRC = "source-wsq/courseware/LP-Carbon-Footprint-Management-for-Sustainability-v7.docx"
OUT = "courseware/LP-Sustainability Masterclass (C752).docx"
shutil.copy(SRC, OUT)
d = docx.Document(OUT)
retitle_all(d)

# Kept-slide map: source slides 3, 9, 11, 14, 15, 16, 86, 130, 131 (1-based) were dropped.
DROPPED = [3, 9, 11, 14, 15, 16, 86, 130, 131]
def m(n):
    n = int(n)
    assert n not in DROPPED, n
    return str(n - sum(1 for x in DROPPED if x < n))
def remap(text):
    return re.sub(r"\b\d+", lambda mo: m(mo.group(0)), text)

# Cover
set_text(para_by_text(d, "TGS Ref No:"), "Course Code: C752")
set_text(para_by_text(d, "Version v7"), "Version v1.0")

t0, t1, t2, t3 = d.tables

# Version control record -> single v1.0 row
set_cell(t0.rows[1].cells[0], "v1.0")
set_cell(t0.rows[1].cells[1], "1 October 2026")
set_cell(t0.rows[1].cells[2], "Initial release of the Sustainability Masterclass lesson plan: two-day schedule "
         "(9:30am–5:30pm, 15 instructional hours) mapped to the v1.0 124-slide deck")
set_cell(t0.rows[1].cells[3], "Dr Alfred Ang")

# TOC: the assessment entry goes
delete(para_by_text(d, "Assessment Schedule — Separate 2 Hours\t"))

# Overview
set_text(para_by_text(d, "Two training days of 8 hours each."),
         "Two days of 7.5 instructional hours each (15 hours in total), 9:30am – 5:30pm with a single "
         "30-minute lunch. Break durations are not counted as training time.")

s1 = [[c.text for c in r.cells] for r in t1.rows]
s2 = [[c.text for c in r.cells] for r in t2.rows]
LUNCH = ["Lunch", "—", "Not counted as training time", "—"]
SEG = {"K1 standards and inventory framing": "Standards and inventory framing",
       "K3 scopes, data and hotspot diagnosis": "Scopes, data and hotspot diagnosis",
       "K7 economics and investment tools": "Economics and investment tools",
       "K4 projections and scenarios": "Projections and scenarios",
       "K5 technologies and feasibility": "Technologies and feasibility",
       "K8 portfolio prioritisation": "Portfolio prioritisation",
       "K2 EMS and governance": "EMS and governance",
       "K6 and K9 plan architecture": "Reduction-plan architecture and action planning"}
def row(src, time, dur, slides=None):
    seg = SEG.get(src[2], src[2])
    return [time, dur, seg, remap(src[3]) if slides is None else slides, src[4], src[5]]

# Day 1: 480 teaching min (45/60/90/60/165/60) rescaled x450/480 onto 450 (5-min granularity);
# the 15-min tea and 60-min lunch become a single 30-min lunch.
rebuild_rows(t1, [
    (1, row(s1[1], "09:30–10:10", "0:40")),
    (2, row(s1[2], "10:10–11:05", "0:55")),
    (4, row(s1[4], "11:05–12:30", "1:25")),
    (5, ["12:30–13:00", "—"] + LUNCH),
    (6, row(s1[6], "13:00–13:55", "0:55")),
    (7, row(s1[7], "13:55–16:35", "2:40")),
    (8, row(s1[8], "16:35–17:30", "0:55")),
])
# Day 2: 480 teaching min (15/75/75/60/90/60/60/45) rescaled the same way; two teas and the
# 60-min lunch become a single 30-min lunch.
rebuild_rows(t2, [
    (1, row(s2[1], "09:30–09:45", "0:15", "—")),
    (2, row(s2[2], "09:45–10:55", "1:10")),
    (4, row(s2[4], "10:55–12:05", "1:10")),
    (5, row(s2[5], "12:05–13:00", "0:55")),
    (6, ["13:00–13:30", "—"] + LUNCH),
    (7, row(s2[7], "13:30–15:00", "1:30")),
    (8, row(s2[8], "15:00–15:55", "0:55")),
    (10, row(s2[10], "15:55–16:50", "0:55")),
    (11, row(s2[11], "16:50–17:30", "0:40")),
])

# Assessment section: heading + table
delete(para_by_text(d, "Assessment Schedule — Separate 2 Hours"))
t3._tbl.getparent().remove(t3._tbl)

# Resources
set_text(para_by_text(d, "v7 course slide deck"), "v1.0 course slide deck")
set_text(para_by_text(d, "v7 Learner Guide"), "v1.0 Learner Guide")

# Footer
for s in d.sections:
    for part in (s.footer, s.header, s.first_page_footer, s.even_page_footer):
        for p in part.paragraphs:
            for r in p.runs:
                if "TGS-2023021099" in r.text:
                    r.text = r.text.replace("TGS-2023021099", "C752")
d.core_properties.title = "LP-Sustainability Masterclass (C752)"
d.save(OUT)
print("saved", OUT)
