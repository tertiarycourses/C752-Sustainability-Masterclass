#!/bin/sh
# Rebuild the C752 deck from the WSQ source: convert_deck.py, then fix_deck.py.
set -e
cd "$(dirname "$0")/.."
OUT="courseware/Sustainability Masterclass (C752)-v1.0.pptx"
python3 ~/.claude/skills/wsq-to-non-wsq/scripts/convert_deck.py \
  --src "source-wsq/courseware/Carbon-Footprint-Management-for-Sustainability-v7.pptx" --out "$OUT" \
  --old-title "Carbon Footprint Management for Sustainability" --new-title "Sustainability Masterclass" \
  --old-code TGS-2023021099 --new-code C752 --tsc "PRE-ENV-4001-1.1" \
  --old-version v7 --new-version v1.0 \
  --sub "Version v7  ·  12 August 2026=>Version v1.0  ·  1 October 2026" \
  --sub "Carbon-Footprint-Management-for-Sustainability=>Sustainability-Masterclass" \
  --sub "TGS-2023021099 · 2 days · 16 training hours=>C752 · 2 days · 15 instructional hours" \
  --wsq-slides 2,8,10,13,14,15,85,129,130 --feedback-slide 131
python3 tools/fix_deck.py
