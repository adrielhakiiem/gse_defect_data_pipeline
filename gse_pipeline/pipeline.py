import re
import datetime as dt
from collections import Counter

import openpyxl

import config

import shutil
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# 0. GET A READABLE COPY (handles the file being open/locked in Excel)
# ---------------------------------------------------------------------------
def get_readable_copy(source_path: Path) -> Path:
    """Copies the (possibly-locked) source into a scratch temp file so the
    pipeline can read it even while it's open in Excel elsewhere."""
    tmp_path = Path(tempfile.gettempdir()) / source_path.name
    shutil.copy2(source_path, tmp_path)
    return tmp_path

# ---------------------------------------------------------------------------
# 1. READ MASTER
# ---------------------------------------------------------------------------
def read_master(wb):
    """Read the existing Master sheet into a list of standardized row dicts."""
    ws = wb[config.MASTER_SHEET]
    headers = [c.value for c in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}
    missing = [name for name in config.STANDARD_COLUMNS if name not in col_idx]
    if missing:
        raise PipelineInputError(
            f"The Master sheet '{config.MASTER_SHEET}' is missing required column(s): {', '.join(missing)}."
        )

    rows = []
    for r in range(2, ws.max_row + 1):
        date = ws.cell(row=r, column=col_idx["Date"]).value
        equip = ws.cell(row=r, column=col_idx["ADS Equipment No"]).value
        if date is None and equip is None:
            continue  # skip fully blank rows
        rows.append({
            "Date": date,
            "By Month": _month_label(ws.cell(row=r, column=col_idx["By Month"]).value),
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


def read_month(wb, month_label, sheet_name):
    """Read one month's raw sheet into standardized row dicts."""
    ws = wb[sheet_name]
    raw_headers = [c.value for c in ws[1]]
    headers_clean = {h.strip() if isinstance(h, str) else h: i + 1
                      for i, h in enumerate(raw_headers) if h}

    col = {name: _resolve_column(headers_clean, name) for name in config.STANDARD_COLUMNS}
    # Date and ADS Equipment No are mandatory; everything else is optional
    # (a missing optional column just means that field is None for the month).
    required = [name for name in ("Date", "ADS Equipment No") if col[name] is None]
    if required:
        raise PipelineInputError(
            f"The monthly sheet '{sheet_name}' is missing required column(s): {', '.join(required)}."
        )

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


class PipelineInputError(Exception):
    """A source-workbook issue that should be explained without a traceback."""


def _month_label(value):
    """Keep By Month as a simple display month, even if source includes a year."""
    label = _norm_ws(value)
    match = re.match(r"[A-Za-z]+", label)
    return match.group(0).title() if match else label


def clean_display(v):
    """Normalize ordinary display values without turning blanks into text."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    normalized = _norm_ws(v)
    return normalized or None


def clean_equipment_number(value):
    """Standardize equipment IDs for reliable joins and Power BI filtering.

    IDs are identifiers rather than prose, so all whitespace and periods are
    removed and the result is uppercased (for example, ``APW 33.`` ->
    ``APW33``).
    """
    value = clean_display(value)
    if value is None:
        return None
    return re.sub(r"[.\s]+", "", value).upper()


def standardize_equipment_numbers(rows):
    """Mutate rows in place and return the number of changed IDs."""
    changed = 0
    for row in rows:
        original = row["ADS Equipment No"]
        cleaned = clean_equipment_number(original)
        if cleaned != original:
            changed += 1
        row["ADS Equipment No"] = cleaned
    return changed


def make_key(row):
    """Build a reconciliation key from canonical representations.

    Reconciliation happens before display cleaning, so key construction must
    independently handle equivalent equipment IDs and date representations.
    """
    return (
        _normalized_date_key(row["Date"]),
        clean_equipment_number(row["ADS Equipment No"]) or "",
        _norm_key_field(row["Defects Description"]),
    )


def _normalized_date_key(value):
    """Return one comparable calendar-date value for text and Excel dates."""
    if isinstance(value, (dt.date, dt.datetime)):
        return value.date().isoformat() if isinstance(value, dt.datetime) else value.isoformat()
    text = _norm_ws(value)
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return text.upper()


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
    
    # Flags only an appended row whose equipment + description exists in Master
    # but whose normalized calendar date is absent from those Master matches.
    def detect_date_typo_suspects(master_rows, to_append):
        master_lookup = {}
        for r in master_rows:
            k = (_norm_key_field(r["ADS Equipment No"]), _norm_key_field(r["Defects Description"]))
            master_lookup.setdefault(k, []).append(r)

        suspects = []
        for r in to_append:
            k = (_norm_key_field(r["ADS Equipment No"]), _norm_key_field(r["Defects Description"]))
            candidates = master_lookup.get(k, [])
            if not candidates or any(_normalized_date_key(c["Date"]) == _normalized_date_key(r["Date"])
                                     for c in candidates):
                continue
            c = candidates[0]
            suspects.append({
                "equipment": r["ADS Equipment No"],
                "description": r["Defects Description"],
                "master_date": c["Date"], "master_month": c["By Month"],
                "new_date": r["Date"], "new_month": r["By Month"],
            })
        return suspects
    
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

    date_typo_suspects = detect_date_typo_suspects(master_rows, to_append)

    def to_clean_row(r, appended):
        out = {k: clean_display(r[k]) for k in config.STANDARD_COLUMNS}
        out["_appended"] = appended
        return out

    def parse_date_sort_key(s):
        if isinstance(s, (dt.date, dt.datetime)):
            return (s.year, s.month, s.day)
        try:
            d, m, y = str(s).strip().split(".")
            return (int(y), int(m), int(d))
        except Exception:
            return (9999, 99, 99)

    master_clean = [to_clean_row(r, False) for r in master_rows]
    append_clean = [to_clean_row(r, True) for r in to_append]

    final_rows = sorted(master_clean + append_clean, key=lambda r: parse_date_sort_key(r["Date"]))

    stats = {
        "master_total": len(master_rows),
        "monthly_totals": {m: len(rows) for m, rows in monthly_rows_by_month.items()},
        "appended_total": len(append_clean),
        "appended_by_month": Counter(r["By Month"] for r in append_clean),
        "anomalies_master_has_more": anomalies_master_has_more,
        "dup_within_month": dup_within_month,
        "date_typo_suspects": date_typo_suspects,
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
    four categorical fields. Returns a stats dict for the audit trail."""
    desc_fixed = 0
    categorical_fixed = 0

    for row in rows:
        new_desc, changed = _fix_description(row["Defects Description"])
        if changed:
            desc_fixed += 1
        row["Defects Description"] = new_desc

        for col_name in ("Defects Categorization", "Defects Specification", "Equipment Type",
                 "Motorized/Non-Motorized"):
            val = row[col_name]
            if val is None:
                continue
            mapping = config.CATEGORICAL_MAP.get(col_name, {})
            key = str(val).strip().lower()
            if key in mapping and mapping[key] != val:
                row[col_name] = mapping[key]
                categorical_fixed += 1

    return {"desc_rows_fixed": desc_fixed, "categorical_fixed": categorical_fixed}


def duplicate_audit(rows):
    """Return duplicate groups by the requested business-key fields.

    This is deliberately separate from exact-row removal: repeated defect
    records are useful operational data and must remain in the export.
    """
    groups = {}
    for row in rows:
        key = tuple(row[field] for field in config.MATCH_KEY_FIELDS)
        groups.setdefault(key, 0)
        groups[key] += 1

    audit_rows = []
    for key, count in groups.items():
        if count > 1:
            audit_rows.append(dict(zip(config.MATCH_KEY_FIELDS, key), **{"Count": count}))
    return sorted(audit_rows, key=lambda row: (str(row["Date"]), row["ADS Equipment No"] or ""))


def remove_exact_duplicates(rows):
    """Remove only rows identical across every canonical data column.

    Rows sharing Date + equipment + description but differing in any other
    field are retained as separate records.
    """
    seen = set()
    unique_rows = []
    removed = 0
    for row in rows:
        key = tuple(row[field] for field in config.STANDARD_COLUMNS)
        if key in seen:
            removed += 1
            continue
        seen.add(key)
        unique_rows.append(row)
    return unique_rows, removed


def blank_counts(rows):
    """Count blank values in the key categorical fields for the QA report."""
    fields = (
        "Date",
        "ADS Equipment No",
        "Maintenance By",
        "Defects Categorization",
        "Defects Specification",
    )
    return {
        field: sum(row[field] is None or not str(row[field]).strip() for row in rows)
        for field in fields
    }


# ---------------------------------------------------------------------------
# DISABLED EQUIPMENT-CODE QA -------------------------------------------------
# Restore this optional, review-only check by uncommenting this block and the
# correspondingly marked blocks in config.py, write_output.py, and run_pipeline.py.
# _ALNUM_RE = re.compile(r"[^A-Z0-9]+")
#
# def _norm_equipment_code(value):
#     return _ALNUM_RE.sub("", str(value or "").upper())
#
# def _leading_code_like_segment(description):
#     if description is None:
#         return None
#     separator = re.search(r"[/-]", str(description))
#     if not separator:
#         return None
#     segment = str(description)[:separator.start()].strip()
#     return segment if re.search(r"\d", segment) else None
#
# def check_equipment_code_matches(rows):
#     """Classify standardized rows; this never changes source values."""
#     counts = Counter()
#     possible_mismatches, no_codes_present = [], []
#     for row in rows:
#         equipment_code = _norm_equipment_code(row["ADS Equipment No"])
#         description_code = _norm_equipment_code(row["Defects Description"])
#         if equipment_code and equipment_code in description_code:
#             counts["matched"] += 1
#             continue
#         leading_code = _leading_code_like_segment(row["Defects Description"])
#         if leading_code:
#             counts["possible_mismatch"] += 1
#             possible_mismatches.append({"row": row, "leading_code_found": leading_code,
#                                         "issue_type": "Possible mismatch"})
#         else:
#             counts["no_code_present"] += 1
#             no_codes_present.append({"row": row, "leading_code_found": "",
#                                      "issue_type": "No code in description"})
#     return {"matched": counts["matched"], "no_code_present": counts["no_code_present"],
#             "possible_mismatch": counts["possible_mismatch"],
#             "review_rows": possible_mismatches + no_codes_present}


# ---------------------------------------------------------------------------
# 6. FLAG MALAY TEXT (never translated automatically)
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
# 7. CONVERT DATE TEXT -> REAL DATE OBJECTS (so Power BI/Excel can use it)
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

# ---------------------------------------------------------------------------
# 8. PREVENT VALUES OTHER THAN BLANK/GSE/TCR IN MAINTENANCE BY
# ---------------------------------------------------------------------------
def fix_maintenance_by_contamination(rows):
    """
    Repairs the recurring 'Motorized'/'Non-Motorized' value leaking into
    Maintenance By (it belongs in the Motorized/Non-Motorized column, not
    here). Repair is evidence-based, not a guess:

      - Look at every OTHER row for the same ADS Equipment No with a valid
        Maintenance By (TCR/GSE).
      - If one value accounts for >= MAINTENANCE_BY_CONFIDENCE_THRESHOLD of
        that equipment's valid entries, use it.
      - Otherwise (no other rows, or a real split), leave it blank and flag
        for manual review - never silently guesses on genuine ambiguity.

    Mutates rows in place. Returns a stats dict for the audit trail.
    """
    valid_by_equip = {}
    for r in rows:
        val = _norm_key_field(r["Maintenance By"])
        if val in config.VALID_MAINTENANCE_BY:
            equip = _norm_key_field(r["ADS Equipment No"])
            valid_by_equip.setdefault(equip, Counter())[val] += 1

    auto_fixed, flagged = [], []
    for r in rows:
        val = _norm_key_field(r["Maintenance By"])
        if val not in config.MAINTENANCE_BY_CONTAMINATION_VALUES:
            continue

        equip = _norm_key_field(r["ADS Equipment No"])
        candidates = valid_by_equip.get(equip)
        inferred = None
        if candidates:
            total = sum(candidates.values())
            best_val, best_n = candidates.most_common(1)[0]
            if best_n / total >= config.MAINTENANCE_BY_CONFIDENCE_THRESHOLD:
                inferred = best_val

        if inferred:
            r["Maintenance By"] = inferred
            r["_maintenance_by_autofixed"] = True
            auto_fixed.append((r["ADS Equipment No"], r["Date"], inferred))
        else:
            r["Maintenance By"] = None
            r["_maintenance_by_flagged"] = True
            flagged.append((r["ADS Equipment No"], r["Date"]))

    return {"auto_fixed": auto_fixed, "flagged_for_review": flagged}
