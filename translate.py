#!/usr/bin/env python3
"""Translate aMiraacle-01-Short.xls using translation memory from translatedTRAL.csv."""

import csv
import re
from collections import Counter

import xlrd
import xlwt
from rapidfuzz import fuzz, process

REFERENCE_CSV = "translatedTRAL.csv"
INPUT_XLS = "aMiraacle-01-Short.xls"
OUTPUT_XLS = "aMiraacle-01-Translated.xls"

FUZZY_THRESHOLD = 88  # Chosen to keep only very close subtitle-memory matches.
MIN_TOKEN_OVERLAP = 0.67
MIN_LENGTH_RATIO = 0.75
MAX_LENGTH_RATIO = 1.35
SHORT_LINE_TOKEN_LIMIT = 8


def normalize(text):
    text = text.strip().lower()
    text = text.replace("’", "'").replace("`", "'")
    text = re.sub(r"\bulan\b", "lan", text)
    text = re.sub(r"[^\w\sçğıöşü]", " ", text, flags=re.UNICODE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def token_overlap_ratio(a, b):
    a_tokens = set(a.split())
    b_tokens = set(b.split())
    if not a_tokens or not b_tokens:
        return 0.0
    return len(a_tokens & b_tokens) / max(len(a_tokens), len(b_tokens))


def is_safe_fuzzy_match(source_norm, candidate_norm, score):
    if not source_norm or not candidate_norm:
        return False
    if score < FUZZY_THRESHOLD:
        return False
    overlap = token_overlap_ratio(source_norm, candidate_norm)
    if overlap < MIN_TOKEN_OVERLAP:
        return False
    source_tokens = set(source_norm.split())
    candidate_tokens = set(candidate_norm.split())
    # Short lines (<=8 tokens) are where fuzzy drift is riskiest, so we reject any
    # candidate that introduces extra tokens not present in the source line.
    if len(source_tokens) <= SHORT_LINE_TOKEN_LIMIT and (candidate_tokens - source_tokens):
        return False
    source_len = max(len(source_norm), 1)
    candidate_len = max(len(candidate_norm), 1)
    length_ratio = source_len / candidate_len
    return MIN_LENGTH_RATIO <= length_ratio <= MAX_LENGTH_RATIO

def build_reference_memory():
    """Build translation memory structures from CSV.

    Returns:
        tuple: (exact_map, normalized_choices, turkish_norm_list, albanian_norm_list)
    """
    print("Loading reference translations from CSV...")
    exact_map = {}  # raw lower -> albanian
    normalized_choices = {}  # normalized TR -> most common AL
    norm_pair_counter = Counter()

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
            raw_key = tr_text.lower()
            norm_key = normalize(tr_text)
            if raw_key not in exact_map:
                exact_map[raw_key] = al_text
            if norm_key:
                norm_pair_counter[(norm_key, al_text)] += 1

    for (norm_key, al_text), _count in norm_pair_counter.most_common():
        if norm_key not in normalized_choices:
            normalized_choices[norm_key] = al_text

    turkish_norm_list = list(normalized_choices.keys())
    albanian_norm_list = [normalized_choices[k] for k in turkish_norm_list]

    print(f"  Loaded {len(exact_map):,} raw exact pairs.")
    print(f"  Loaded {len(turkish_norm_list):,} normalized reference pairs.")
    return exact_map, normalized_choices, turkish_norm_list, albanian_norm_list


def translate(turkish_text, exact_map, normalized_choices, turkish_norm_list, albanian_norm_list):
    """Return (albanian_text, stats_bucket, match_label) for a Turkish sentence."""
    text = turkish_text.strip()
    if not text:
        return ("", "empty", "empty")

    # 1) Raw exact match
    key = text.lower()
    if key in exact_map:
        return (exact_map[key], "exact_raw", "exact_raw")

    # 2) Normalized exact match (punctuation/casing tolerant)
    norm_key = normalize(text)
    if norm_key in normalized_choices:
        return (normalized_choices[norm_key], "exact_norm", "exact_norm")

    # 3) Strict fuzzy match on normalized Turkish
    result = process.extractOne(
        norm_key,
        turkish_norm_list,
        scorer=fuzz.token_sort_ratio,
    )
    if result:
        matched_norm, score, matched_index = result
        if is_safe_fuzzy_match(norm_key, matched_norm, score):
            label = f"fuzzy_safe({score:.0f}%)"
            return (albanian_norm_list[matched_index], "fuzzy_safe", label)

    # 4) No safe match found
    return (f"[NEEDS REVIEW] {text}", "no_match", "no_match")


def main():
    """Run end-to-end XLS translation and write output with ALBANIAN column."""
    exact_map, normalized_choices, turkish_norm_list, albanian_norm_list = build_reference_memory()

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
    # TIME, CHARACTER, TURKISH, ENGLISH, SPANISH, ALBANIAN
    col_widths = [3000, 5000, 12000, 12000, 8000, 12000]

    stats = {"exact_raw": 0, "exact_norm": 0, "fuzzy_safe": 0, "no_match": 0, "empty": 0}

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
        albanian_text, match_bucket, match_label = translate(
            turkish_text, exact_map, normalized_choices, turkish_norm_list, albanian_norm_list
        )
        out_row[5] = albanian_text

        stats[match_bucket] = stats.get(match_bucket, 0) + 1

        for col_idx, val in enumerate(out_row[:6]):
            style = needs_review_style if match_bucket == "no_match" and col_idx == 5 else normal_style
            ws_out.write(row_idx, col_idx, val, style)

        if (row_idx - 4) % 50 == 0:
            print(
                f"  Row {row_idx - 4}/{ws_in.nrows - 5} ... latest: "
                f"{turkish_text[:40]!r} -> {match_label}"
            )

    # Set column widths
    for i, w in enumerate(col_widths):
        ws_out.col(i).width = w

    wb_out.save(OUTPUT_XLS)
    print(f"\nDone! Saved to {OUTPUT_XLS}")
    print(
        "Stats: "
        f"exact_raw={stats.get('exact_raw',0)}, "
        f"exact_norm={stats.get('exact_norm',0)}, "
        f"fuzzy_safe={stats.get('fuzzy_safe',0)}, "
        f"no_match={stats.get('no_match',0)}, "
        f"empty={stats.get('empty',0)}"
    )


if __name__ == "__main__":
    main()
