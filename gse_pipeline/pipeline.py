"""
pipeline.py
===========
Core logic for the EDR GSE consolidation pipeline:

  1. read_master()     - load the existing merged Master sheet
  2. read_month()       - load one month's raw sheet, normalized to the
                           standard column schema
  3. reconcile()         - figure out exactly which rows from each month
                           are missing from Master, using a composite key
                           so legitimate repeat entries aren't lost or
                           double-counted
  4. standardize_text()  - fix known misspellings, punctuation, and
                           categorical-field inconsistencies
  5. flag_malay()        - mark (not translate) rows containing Malay text

Each function is independently testable and takes/returns plain Python
lists of dicts, so this logic isn't tied to openpyxl and could be swapped
to pandas later without rewriting the reconciliation rules.
"""

import re
import datetime as dt
from collections import Counter

import openpyxl

import config


# ---------------------------------------------------------------------------
# 1. READ MASTER
# ---------------------------------------------------------------------------
def read_master(wb):
    """Read the existing Master sheet into a list of standardized row dicts."""
    ws = wb[config.MASTER_SHEET]
    headers = [c.value for c in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}

    rows = []
    for r in range(2, ws.max_row + 1):
        date = ws.cell(row=r, column=col_idx["Date"]).value
        equip = ws.cell(row=r, column=col_idx["ADS Equipment No"]).value
        if date is None and equip is None:
            continue  # skip fully blank rows
        rows.append({
            "Date": date,
            "By Month": ws.cell(row=r, column=col_idx["By Month"]).value,
            "ADS Equipment No": equip,
            "Equipment Type": ws.cell(row=r, column=col_idx["Equipment Type"]).value,
            "Maintenance By": ws.cell(row=r, column=col_idx["Maintenance By"]).value,
            "Motorized/Non-Motorized": ws.cell(row=r, column=col_idx["Motorized/Non-Motorized"]).value,
            "Defects Description": ws.cell(row=r, column=col_idx["Defects Description"]).value,
            "Defects Categorization": ws.cell(row=r, column=col_idx["Defects Categorization"]).value,
            "Defects Specification": ws.cell(row=r, column=col_idx["Defects Specification"]).value,
        })
    return rows


# ---------------------------------------------------------------------------
# 2. READ ONE MONTH
# ---------------------------------------------------------------------------
def _resolve_column(headers_clean, standard_name):
    """Find which header in this sheet corresponds to a standard column,
    checking COLUMN_ALIASES for known alternate names (e.g. Ownership)."""
    candidates = config.COLUMN_ALIASES.get(standard_name, [standard_name])
    for name in candidates:
        if name in headers_clean:
            return headers_clean[name]
    return None


def read_month(wb, month_label):
    """Read one month's raw sheet into standardized row dicts."""
    sheet_name = config.MONTH_SHEETS[month_label]
    ws = wb[sheet_name]
    raw_headers = [c.value for c in ws[1]]
    headers_clean = {h.strip() if isinstance(h, str) else h: i + 1
                      for i, h in enumerate(raw_headers) if h}

    col = {name: _resolve_column(headers_clean, name) for name in config.STANDARD_COLUMNS}
    # Date and ADS Equipment No are mandatory; everything else is optional
    # (a missing optional column just means that field is None for the month).

    rows = []
    for r in range(2, ws.max_row + 1):
        date = ws.cell(row=r, column=col["Date"]).value
        equip = ws.cell(row=r, column=col["ADS Equipment No"]).value
        if date is None and equip is None:
            continue
        row = {"By Month": month_label}
        for name in config.STANDARD_COLUMNS:
            if name in ("Date", "By Month"):
                continue
            c = col[name]
            row[name] = ws.cell(row=r, column=c).value if c else None
        row["Date"] = date
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# HELPERS: normalization for matching (display values are never touched here)
# ---------------------------------------------------------------------------
def _norm_ws(s):
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s).strip())


def _norm_key_field(v):
    return _norm_ws(v).upper()


def clean_display(v):
    """Normalize a value for DISPLAY (not just matching): trims whitespace
    and converts numeric equipment numbers to text so the whole column has
    a consistent type."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    return _norm_ws(v)


def make_key(row):
    return tuple(_norm_key_field(row[f]) for f in config.MATCH_KEY_FIELDS)


# ---------------------------------------------------------------------------
# 3. RECONCILE
# ---------------------------------------------------------------------------
def reconcile(master_rows, monthly_rows_by_month):
    """
    For each month, compare per-key counts between Master and the monthly
    source. Uses a COUNT-based comparison (not a simple set difference) so
    that legitimate repeat entries (same equipment, different defect, same
    key by coincidence - or genuinely repeated) are handled correctly:

      - If a key appears more times in the monthly sheet than in Master,
        the EXTRA occurrences are appended (not all of them, and not none).
      - If a key appears the same or fewer times in monthly than Master,
        nothing is appended for that key.
      - If Master somehow has MORE copies of a key than the monthly source
        has, that's flagged as an anomaly for manual review (it did not
        happen in the original dataset, but the check stays in the
        pipeline for future data).

    Returns:
        final_rows: list of dicts (cleaned display values), each tagged
                    with '_appended' (bool) and 'By Month'
        stats: dict of counts useful for the audit trail / QA sheet
    """
    master_key_counts = Counter(make_key(r) for r in master_rows)

    to_append = []
    anomalies_master_has_more = []
    dup_within_month = []

    for month, rows in monthly_rows_by_month.items():
        month_key_counts = Counter(make_key(r) for r in rows)
        rows_by_key = {}
        for row in rows:
            rows_by_key.setdefault(make_key(row), []).append(row)

        for key, month_count in month_key_counts.items():
            master_count = master_key_counts.get(key, 0)
            if month_count > 1:
                dup_within_month.append((month, key, month_count))
            if month_count > master_count:
                extra_needed = month_count - master_count
                candidates = rows_by_key[key][-extra_needed:]
                to_append.extend(candidates)
            elif month_count < master_count:
                anomalies_master_has_more.append((month, key, month_count, master_count))

    def to_clean_row(r, appended):
        out = {k: clean_display(r[k]) for k in config.STANDARD_COLUMNS}
        out["_appended"] = appended
        return out

    def parse_date_sort_key(s):
        try:
            d, m, y = str(s).strip().split(".")
            return (int(y), int(m), int(d))
        except Exception:
            return (9999, 99, 99)

    master_clean = [to_clean_row(r, False) for r in master_rows]
    append_clean = [to_clean_row(r, True) for r in to_append]

    final_rows = []
    for month in config.MONTH_ORDER:
        existing_block = [r for r in master_clean if r["By Month"] == month]
        new_block = sorted(
            [r for r in append_clean if r["By Month"] == month],
            key=lambda r: parse_date_sort_key(r["Date"]),
        )
        final_rows.extend(existing_block)
        final_rows.extend(new_block)

    stats = {
        "master_total": len(master_rows),
        "monthly_totals": {m: len(rows) for m, rows in monthly_rows_by_month.items()},
        "appended_total": len(append_clean),
        "appended_by_month": Counter(r["By Month"] for r in append_clean),
        "anomalies_master_has_more": anomalies_master_has_more,
        "dup_within_month": dup_within_month,
        "equip_numeric_to_text": sum(1 for r in to_append if isinstance(r["ADS Equipment No"], (int, float))),
        "equip_whitespace_trimmed": sum(
            1 for r in to_append
            if isinstance(r["ADS Equipment No"], str) and r["ADS Equipment No"] != r["ADS Equipment No"].strip()
        ),
    }
    return final_rows, stats


# ---------------------------------------------------------------------------
# 4. STANDARDIZE TEXT
# ---------------------------------------------------------------------------
_WORD_RE = re.compile(r"[A-Za-z']+")
_SPELLING_MAP_UPPER = {k.upper(): v for k, v in config.SPELLING_MAP.items()}


def _apply_case(token, target_upper):
    if token.isupper():
        return target_upper
    if token.istitle():
        return target_upper.capitalize()
    if token.islower():
        return target_upper.lower()
    return target_upper


def _fix_description(desc):
    """Apply, in order: shift-key parenthesis-typo fix, spelling
    corrections, 'non negligence' hyphenation, trailing-period removal,
    parenthesis-spacing tightening, and case normalization.

    Order matters: spelling must be corrected BEFORE the hyphenation check,
    otherwise a row like "NON NEGLUGENCE" only gets the spelling fixed and
    never gets the space-to-hyphen conversion because the un-corrected
    phrase doesn't match the "non negligence" pattern being searched for
    inside the corrected pass. Trailing periods must be stripped in a loop
    (some rows had DOUBLE trailing periods, e.g. ". .", which a single
    non-repeating regex would only catch once).
    """
    if desc is None:
        return desc, False
    original = desc
    s = desc

    # keyboard slip: "(" mistyped as "9", ")" mistyped as "0", around NON/NEGLIGENCE
    s = re.sub(r"(?<![0-9])9(?=NON\b)", "(", s, flags=re.I)
    s = re.sub(r"(NEGLIGENCE)0(?!\d)", r"\1)", s, flags=re.I)

    # word-level spelling corrections
    def repl(m):
        w = m.group(0)
        key = w.upper()
        return _apply_case(w, _SPELLING_MAP_UPPER[key]) if key in _SPELLING_MAP_UPPER else w
    s = _WORD_RE.sub(repl, s)

    # "non negligence" -> "non-negligence" (after spelling is normalized)
    s = re.sub(
        r"\bnon\s+negligence\b",
        lambda m: ("non-negligence" if m.group(0).islower()
                    else "NON-NEGLIGENCE" if m.group(0).isupper()
                    else "Non-Negligence"),
        s, flags=re.I,
    )

    s = s.strip()
    # strip ALL trailing periods (handles the ". ." double-period cases)
    s = re.sub(r"(\s*\.)+$", "", s)
    # tighten "( X )" -> "(X)"
    s = re.sub(r"\(\s+", "(", s)
    s = re.sub(r"\s+\)", ")", s)
    # uppercase the small minority of mixed/lowercase rows for consistency
    if s != s.upper():
        s = s.upper()

    return s, (s != original)


def standardize_text(rows):
    """Mutates rows in place: standardizes Defects Description and the
    three categorical fields. Returns a stats dict for the audit trail."""
    desc_fixed = 0
    categorical_fixed = 0

    for row in rows:
        new_desc, changed = _fix_description(row["Defects Description"])
        if changed:
            desc_fixed += 1
        row["Defects Description"] = new_desc

        for col_name in ("Defects Categorization", "Defects Specification", "Equipment Type"):
            val = row[col_name]
            if val is None:
                continue
            mapping = config.CATEGORICAL_MAP.get(col_name, {})
            key = str(val).strip().lower()
            if key in mapping and mapping[key] != val:
                row[col_name] = mapping[key]
                categorical_fixed += 1

    return {"desc_rows_fixed": desc_fixed, "categorical_fixed": categorical_fixed}


# ---------------------------------------------------------------------------
# 5. FLAG MALAY TEXT (never translated automatically)
# ---------------------------------------------------------------------------
_MALAY_PATTERN = re.compile(r"\b(" + "|".join(config.MALAY_WORDS) + r")\b", re.I)


def flag_malay(rows):
    """Mutates rows in place, adding '_malay_flag'. Returns the count flagged."""
    flagged = 0
    for row in rows:
        desc = row["Defects Description"]
        is_malay = bool(desc and _MALAY_PATTERN.search(str(desc)))
        row["_malay_flag"] = is_malay
        if is_malay:
            flagged += 1
    return flagged


# ---------------------------------------------------------------------------
# 6. CONVERT DATE TEXT -> REAL DATE OBJECTS (so Power BI/Excel can use it)
# ---------------------------------------------------------------------------
def convert_dates(rows):
    """Mutates rows in place: 'DD.MM.YYYY' text -> datetime.date. Rows that
    fail to parse are left as-is and counted so nothing fails silently."""
    failures = 0
    for row in rows:
        s = row["Date"]
        try:
            d, m, y = str(s).strip().split(".")
            row["Date"] = dt.date(int(y), int(m), int(d))
        except Exception:
            failures += 1
    return failures
