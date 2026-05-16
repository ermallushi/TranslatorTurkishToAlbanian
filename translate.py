#!/usr/bin/env python3
"""
Translate aMiraacle-01-Short.xls Turkish column to Albanian,
using translatedTRAL.csv as style/word reference.

Strategy:
1. Exact match lookup (case-insensitive, stripped)
2. Fuzzy match (token_sort_ratio >= 70) from the reference set
3. Mark unmatched lines as [NEEDS REVIEW]
"""

import csv
import xlrd
import xlwt
from rapidfuzz import process, fuzz

REFERENCE_CSV = "translatedTRAL.csv"
INPUT_XLS = "aMiraacle-01-Short.xls"
OUTPUT_XLS = "aMiraacle-01-Translated.xls"

FUZZY_THRESHOLD = 55  # minimum similarity score to accept a fuzzy match

print("Loading reference translations from CSV...")
exact_map = {}   # normalised_turkish -> albanian
turkish_list = []  # for fuzzy matching
albanian_list = []  # parallel list

with open(REFERENCE_CSV, encoding="utf-8-sig") as f:
    reader = csv.reader(f, delimiter="|")
    next(reader)  # skip header
    for row in reader:
        if len(row) < 5:
            continue
        tr_text = row[2].strip()
        al_text = row[4].strip()
        if not tr_text or not al_text:
            continue
        key = tr_text.lower()
        if key not in exact_map:
            exact_map[key] = al_text
            turkish_list.append(tr_text)
            albanian_list.append(al_text)

print(f"  Loaded {len(turkish_list):,} unique reference pairs.")

# Pre-normalised keys list for fuzzy matching
tr_keys_lower = [t.lower() for t in turkish_list]


def translate(turkish_text):
    """Return (albanian_text, match_type) for a Turkish sentence."""
    text = turkish_text.strip()
    if not text:
        return ("", "empty")

    # 1. Exact match
    key = text.lower()
    if key in exact_map:
        return (exact_map[key], "exact")

    # 2. Fuzzy match
    result = process.extractOne(
        key,
        tr_keys_lower,
        scorer=fuzz.token_sort_ratio,
        score_cutoff=FUZZY_THRESHOLD,
    )
    if result:
        matched_key, score, idx = result
        return (albanian_list[idx], f"fuzzy({score:.0f}%)")

    # 3. No match
    return (f"[NEEDS REVIEW] {text}", "no_match")


print("Reading input XLS...")
wb_in = xlrd.open_workbook(INPUT_XLS)
ws_in = wb_in.sheet_by_index(0)

# Build output workbook
wb_out = xlwt.Workbook(encoding="utf-8")
ws_out = wb_out.add_sheet("Sheet1")

# Styles
header_style = xlwt.easyxf(
    "font: bold true; pattern: pattern solid, fore_colour light_blue;"
)
meta_style = xlwt.easyxf("font: bold true;")
normal_style = xlwt.easyxf("")
needs_review_style = xlwt.easyxf("font: colour red;")

# Column widths (approximate)
col_widths = [3000, 5000, 12000, 12000, 8000, 12000]

stats = {"exact": 0, "fuzzy": 0, "no_match": 0, "empty": 0}

print("Translating rows...")
for row_idx in range(ws_in.nrows):
    row_vals = ws_in.row_values(row_idx)

    # Ensure we have 6 columns in output (add ALBANIAN at index 5)
    out_row = list(row_vals) + [""]

    if row_idx < 4:
        # Metadata rows (title, episode, type, blank)
        for col_idx, val in enumerate(out_row[:6]):
            style = meta_style if val else normal_style
            ws_out.write(row_idx, col_idx, val, style)
        continue

    if row_idx == 4:
        # Header row
        out_row[5] = "ALBANIAN"
        for col_idx, val in enumerate(out_row[:6]):
            ws_out.write(row_idx, col_idx, val, header_style)
        continue

    # Data rows: col 2 = TURKISH
    turkish_text = str(row_vals[2]).strip() if len(row_vals) > 2 else ""
    albanian_text, match_type = translate(turkish_text)
    out_row[5] = albanian_text

    stats[match_type.split("(")[0]] = stats.get(match_type.split("(")[0], 0) + 1

    for col_idx, val in enumerate(out_row[:6]):
        style = needs_review_style if match_type == "no_match" and col_idx == 5 else normal_style
        ws_out.write(row_idx, col_idx, val, style)

    if (row_idx - 4) % 50 == 0:
        print(f"  Row {row_idx - 4}/{ws_in.nrows - 5} ... latest: {turkish_text[:40]!r} -> {match_type}")

# Set column widths
for i, w in enumerate(col_widths):
    ws_out.col(i).width = w

wb_out.save(OUTPUT_XLS)
print(f"\nDone! Saved to {OUTPUT_XLS}")
print(f"Stats: exact={stats.get('exact',0)}, fuzzy={stats.get('fuzzy',0)}, no_match={stats.get('no_match',0)}, empty={stats.get('empty',0)}")
