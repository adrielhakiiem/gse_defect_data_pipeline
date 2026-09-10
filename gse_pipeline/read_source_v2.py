"""Read and normalize the 2026 CMMS EDR export."""

import argparse
import datetime as dt
import re
from collections import Counter
from pathlib import Path

import xlrd

import config


SOURCE_COLUMNS = (
    "Permit No.",
    "Work Status",
    "Work Order Description",
    "Work Order Type",
    "Received Date",
    "Asset Number",
    "Asset Description",
    "Action Taken",
    "Work Order Trade",
)

_PREFIX_RE = re.compile(
    r"^\s*([A-Za-z0-9.\-/ ]{2,15}?)[\-/]\s*(\S.*)?$"
)


class SourceInputError(Exception):
    """A source workbook cannot be processed as expected."""


def _clean_header(value):
    return str(value).strip() if value is not None else ""


def _blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def _read_source_workbook(workbook):
    """Return raw row dictionaries from an opened ``Worksheet`` workbook."""
    if "Worksheet" not in workbook.sheet_names():
        raise SourceInputError("Expected a sheet named 'Worksheet'.")

    sheet = workbook.sheet_by_name("Worksheet")
    if sheet.nrows == 0:
        return []
    headers = [_clean_header(sheet.cell_value(0, column))
               for column in range(sheet.ncols)]
    columns = {header: index for index, header in enumerate(headers) if header}
    missing = [name for name in SOURCE_COLUMNS if name not in columns]
    if missing:
        raise SourceInputError(
            "Missing required column(s): " + ", ".join(missing)
        )

    rows = []
    for row_index in range(1, sheet.nrows):
        values = [sheet.cell_value(row_index, column) for column in range(sheet.ncols)]
        if all(_blank(value) for value in values):
            continue
        rows.append({name: values[index] for name, index in columns.items()})
    return rows


def read_source(source_path=None):
    """Return raw row dictionaries from the single ``Worksheet`` sheet."""
    path = Path(source_path or config.SOURCE_FILE)
    workbook = xlrd.open_workbook(path, on_demand=True)
    try:
        return _read_source_workbook(workbook)
    finally:
        workbook.release_resources()


def _as_text(value):
    if _blank(value):
        return None
    return str(value).strip()


def _extract_code(description):
    text = _as_text(description)
    if not text:
        return None, None
    match = _PREFIX_RE.match(text)
    if not match or not match.group(2):
        return None, text
    code = match.group(1).strip(" .-/")
    body = match.group(2).strip()
    if not code or not body:
        return None, text
    return code, body


def _datetime_value(value, workbook_datemode):
    if isinstance(value, dt.datetime):
        return value
    if isinstance(value, dt.date):
        return dt.datetime.combine(value, dt.time())
    if isinstance(value, (int, float)):
        try:
            return xlrd.xldate_as_datetime(value, workbook_datemode)
        except (ValueError, OverflowError):
            return None
    text = _as_text(value)
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def normalize_rows(raw_rows, workbook_datemode=0):
    """Convert raw source rows into the v2 output schema."""
    normalized = []
    failed_examples = []
    for row in raw_rows:
        code, description = _extract_code(row.get("Work Order Description"))
        if code is None:
            code = _as_text(row.get("Asset Number"))
        if code is None:
            failed_examples.append(row.get("Work Order Description"))
        received_datetime = _datetime_value(row.get("Received Date"), workbook_datemode)
        date = received_datetime.date() if received_datetime else None
        normalized.append({
            "Permit No": row.get("Permit No."),
            "Work Status": row.get("Work Status"),
            "Equipment Code": code,
            "Asset Number": row.get("Asset Number"),
            "Asset Description": row.get("Asset Description"),
            "Defects Description": description,
            "Work Order Type": row.get("Work Order Type"),
            "Work Order Trade": row.get("Work Order Trade"),
            "Action Taken": row.get("Action Taken"),
            "Date": date,
            "By Month": date.strftime("%B") if date else None,
            "_received_datetime": received_datetime,
        })
    return normalized, failed_examples


def read_and_normalize(source_path=None):
    """Read the workbook and return normalized rows plus failed examples."""
    path = Path(source_path or config.SOURCE_FILE)
    workbook = xlrd.open_workbook(path, on_demand=True)
    try:
        raw_rows = _read_source_workbook(workbook)
        return normalize_rows(raw_rows, workbook.datemode)
    finally:
        workbook.release_resources()


def print_summary(rows, failed_examples):
    """Print the compact QA summary for a normalized source."""
    print(f"Total rows read: {len(rows)}")
    print(f"Equipment code extraction failed: {len(failed_examples)}")
    if failed_examples:
        print("Examples:")
        for example in failed_examples[:5]:
            print(f"  {example!r}")
    print("\nRows by Work Order Type:")
    for value, count in sorted(Counter(row["Work Order Type"] for row in rows).items(), key=lambda item: str(item[0])):
        print(f"  {value or '<blank>'}: {count}")
    print("\nRows by Work Status:")
    for value, count in sorted(Counter(row["Work Status"] for row in rows).items(), key=lambda item: str(item[0])):
        print(f"  {value or '<blank>'}: {count}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=config.SOURCE_FILE)
    args = parser.parse_args()
    print(f"Loading {args.source}...")
    try:
        rows, failed_examples = read_and_normalize(args.source)
    except (FileNotFoundError, PermissionError, OSError, xlrd.XLRDError,
            SourceInputError) as exc:
        print(f"ERROR: Could not read source workbook: {exc}")
        return 1
    print_summary(rows, failed_examples)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())