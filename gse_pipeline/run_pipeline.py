"""
run_pipeline.py
================
Entry point. Run with:

    python run_pipeline.py

Reads config.SOURCE_FILE, produces config.OUTPUT_FILE, and prints a summary
of every step so the whole process is visible and auditable - not just the
final numbers.
"""

import re
from collections import Counter

import openpyxl

import config
import pipeline
import write_output


def _spelling_occurrence_counts(original_rows):
    """Re-derive, from the ORIGINAL (pre-cleaning) text, exactly how many
    times each known misspelling occurred - used for the Change Log sheet
    and the console summary."""
    all_text = " ".join(str(r["Defects Description"] or "") for r in original_rows)
    counts = []
    for bad, good in config.SPELLING_MAP.items():
        n = len(re.findall(r"\b" + re.escape(bad) + r"\b", all_text, re.I))
        if n > 0:
            counts.append((bad, good, n))
    return counts


def _categorical_change_counts(original_rows):
    out = []
    for col, mapping in config.CATEGORICAL_MAP.items():
        c = Counter()
        for r in original_rows:
            v = r[col]
            if v is None:
                continue
            key = str(v).strip().lower()
            if key in mapping and mapping[key] != v:
                c[(v, mapping[key])] += 1
        for (before, after), cnt in c.items():
            out.append((col, before, after, cnt))
    return out


def main():
    print(f"Loading source workbook: {config.SOURCE_FILE}")
    wb = openpyxl.load_workbook(config.SOURCE_FILE, data_only=True)

    # --- 1. Read ---
    master_rows = pipeline.read_master(wb)
    print(f"  Master rows read: {len(master_rows)}")

    monthly_rows_by_month = {}
    for month in config.MONTH_ORDER:
        rows = pipeline.read_month(wb, month)
        monthly_rows_by_month[month] = rows
        print(f"  {month} rows read: {len(rows)}")

    # --- 2. Reconcile (find what's missing, append it) ---
    final_rows, stats = pipeline.reconcile(master_rows, monthly_rows_by_month)
    print(f"\nRows appended: {stats['appended_total']} "
          f"({dict(stats['appended_by_month'])})")
    print(f"Final merged row count: {len(final_rows)}")
    if stats["anomalies_master_has_more"]:
        print(f"  WARNING: {len(stats['anomalies_master_has_more'])} rows where Master has "
              f"MORE copies of a key than the source - review these manually.")

    # keep a copy of the pre-cleaning text for the change-log counts below
    original_rows_snapshot = [dict(r) for r in final_rows]

    # --- 3. Standardize text ---
    text_stats = pipeline.standardize_text(final_rows)
    print(f"\nDefects Description rows changed: {text_stats['desc_rows_fixed']}")
    print(f"Categorical field cells standardized: {text_stats['categorical_fixed']}")

    # --- 4. Flag Malay text (not translated) ---
    malay_count = pipeline.flag_malay(final_rows)
    print(f"Rows flagged for Malay content (not translated): {malay_count}")

    # --- 5. Convert Date to real date type ---
    date_failures = pipeline.convert_dates(final_rows)
    print(f"Dates converted to true date type. Parse failures: {date_failures}")

    # --- 6. Build change-log detail for the output workbook ---
    spelling_counts = _spelling_occurrence_counts(original_rows_snapshot)
    categorical_rows = _categorical_change_counts(original_rows_snapshot)
    format_fix_counts = [
        ("\"non negligence\" -> \"non-negligence\" (spacing)",
         sum(1 for r in original_rows_snapshot
             if r["Defects Description"] and re.search(r"\bnon\s+negligence\b", r["Defects Description"], re.I))),
        ("Trailing period(s) removed",
         sum(1 for r in original_rows_snapshot
             if r["Defects Description"] and str(r["Defects Description"]).strip().endswith("."))),
        ("Parenthesis spacing tightened",
         sum(1 for r in original_rows_snapshot
             if r["Defects Description"] and re.search(r"\(\s|\s\)", str(r["Defects Description"])))),
    ]

    # --- 7. Write output workbook ---
    write_output.write_workbook(
        final_rows, stats, text_stats, malay_count, date_failures,
        spelling_counts, categorical_rows, format_fix_counts,
        config.OUTPUT_FILE,
    )
    print(f"\nSaved: {config.OUTPUT_FILE}")


if __name__ == "__main__":
    main()
