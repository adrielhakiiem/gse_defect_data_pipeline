from collections import Counter

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import config

OUT_COLS = ["No"] + config.STANDARD_COLUMNS

_HEADER_FONT = Font(name=config.FONT_NAME, bold=True, color="FFFFFF", size=10)
_HEADER_FILL = PatternFill("solid", fgColor=config.COLOR_HEADER_FILL)
_BODY_FONT = Font(name=config.FONT_NAME, size=10)
_APPENDED_FILL = PatternFill("solid", fgColor=config.COLOR_APPENDED_ROW)
_MALAY_FILL = PatternFill("solid", fgColor=config.COLOR_MALAY_FLAG)
_THIN = Side(style="thin", color="D9D9D9")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_TITLE_FONT = Font(name=config.FONT_NAME, bold=True, size=13)
_SECTION_FONT = Font(name=config.FONT_NAME, bold=True, size=11, color=config.COLOR_HEADER_FILL)
_NOTE_FONT = Font(name=config.FONT_NAME, size=10)


def _style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _BORDER
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 28


def _autosize(ws, ncols, sample_rows=500, max_width=60):
    for c in range(1, ncols + 1):
        letter = get_column_letter(c)
        max_len = len(str(ws.cell(row=1, column=c).value or ""))
        for r in range(2, min(ws.max_row, sample_rows) + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), max_width)


def _write_master_sheet(wb, rows):
    ws = wb.active
    ws.title = "Master Data (Merged)"
    ws.append(OUT_COLS)
    for i, row in enumerate(rows, start=1):
        vals = [i] + [row[c] for c in config.STANDARD_COLUMNS]
        ws.append(vals)
        r_idx = i + 1
        ws.cell(row=r_idx, column=2).number_format = "DD.MM.YYYY"
        for c in range(1, len(OUT_COLS) + 1):
            cell = ws.cell(row=r_idx, column=c)
            cell.font = _BODY_FONT
            cell.border = _BORDER
            if row.get("_malay_flag"):
                cell.fill = _MALAY_FILL
            elif row.get("_appended"):
                cell.fill = _APPENDED_FILL
    _style_header(ws, len(OUT_COLS))
    _autosize(ws, len(OUT_COLS))
    ws.column_dimensions["H"].width = 55
    ws.column_dimensions["B"].width = 12
    return ws


def _write_audit_sheet(wb, rows):
    ws = wb.create_sheet("Audit - Rows Added")
    cols = ["Date", "By Month (Source)"] + config.STANDARD_COLUMNS[2:]
    ws.append(cols)
    appended = [r for r in rows if r.get("_appended")]
    for row in appended:
        ws.append([row["Date"], row["By Month"]] + [row[c] for c in config.STANDARD_COLUMNS[2:]])
        ws.cell(row=ws.max_row, column=1).number_format = "DD.MM.YYYY"
        if row.get("_malay_flag"):
            for c in range(1, len(cols) + 1):
                ws.cell(row=ws.max_row, column=c).fill = _MALAY_FILL
    for r_idx in range(2, ws.max_row + 1):
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=r_idx, column=c)
            cell.font = _BODY_FONT
            cell.border = _BORDER
    _style_header(ws, len(cols))
    _autosize(ws, len(cols))
    ws.column_dimensions["G"].width = 55
    return ws, len(appended)


def _write_section(ws, r, title, lines):
    ws.cell(row=r, column=1, value=title).font = _SECTION_FONT
    r += 1
    for line in lines:
        ws.cell(row=r, column=1, value=line).font = _NOTE_FONT
        r += 1
    return r + 1


def _write_qa_sheet(wb, stats, text_stats, malay_count, date_failures, maint_stats):
    ws = wb.create_sheet("QA - Anomalies & Notes")
    ws.column_dimensions["A"].width = 105
    ws.cell(row=1, column=1, value="Data Quality & Reconciliation Notes").font = _TITLE_FONT
    r = 3

    monthly = stats["monthly_totals"]
    r = _write_section(ws, r, "1. Row count reconciliation", [
        f"Monthly source rows (header excluded): " + ", ".join(f"{m} {n}" for m, n in monthly.items())
        + f" = {sum(monthly.values())} total.",
        f"Original Master row count: {stats['master_total']}.",
        f"Rows appended: {stats['appended_total']} "
        + f"({', '.join(f'{m} {n}' for m, n in stats['appended_by_month'].items())}).",
        "Text-cleaning steps never change row counts - only cell content. Re-verify this after any pipeline change.",
    ])
    r = _write_section(ws, r, "2. Duplicate-key anomalies (Master has MORE copies than source)", [
        f"{len(stats['anomalies_master_has_more'])} found. Each one means Master contains a row that "
        "can't be traced back to the monthly source under its tagged month - investigate manually.",
    ])
    r = _write_section(ws, r, "3. Legitimate repeat entries preserved", [
        f"{len(stats['dup_within_month'])} composite keys (Date + Equipment No + Description) appear "
        "more than once within a single month's source - kept individually, not collapsed, since each "
        "represents a separate logged defect for the same equipment.",
    ])
    r = _write_section(ws, r, "4. ADS Equipment No formatting normalization", [
        f"{stats['equip_numeric_to_text']} rows: numeric equipment no. converted to text.",
        f"{stats['equip_whitespace_trimmed']} rows: leading/trailing whitespace trimmed.",
    ])
    r = _write_section(ws, r, "5. Text & categorical standardization", [
        f"{text_stats['desc_rows_fixed']} Defects Description rows had text standardized "
        "(spelling, punctuation, or case) - see the Change Log sheet for the full breakdown.",
        f"{text_stats['categorical_fixed']} categorical field cells standardized "
        "(Equipment Type / Defects Categorization / Defects Specification).",
    ])
    r = _write_section(ws, r, "6. Malay-language text - flagged, not translated", [
        f"{malay_count} rows contain Malay words in Defects Description, highlighted YELLOW "
        "in the Master and Audit sheets. Translation deliberately deferred to a manual pass.",
    ])
    r = _write_section(ws, r, "7. Date column converted to true date type", [
        f"Converted from text ('DD.MM.YYYY') to a real date value (still displays the same way) "
        f"so Power BI / Excel can sort, filter, and do time-intelligence correctly. "
        f"Parse failures: {date_failures}.",
    ])
    r = _write_section(ws, r, "8. Possible date-typo duplicates (same equipment + defect as Master, different date)", [
    f"{len(stats['date_typo_suspects'])} found - each pairs a Master row with an appended row that share "
    "equipment + defect description but disagree on date. Almost always a typo in one of the two dates. "
    "Verify and fix at the source before trusting the row count.",
    ])
    r = _write_section(ws, r, "9. Maintenance By contamination (auto-corrected)", [
    f"{len(maint_stats['auto_fixed'])} rows had 'Motorized'/'Non-Motorized' in Maintenance By, "
    "corrected using the equipment's other valid entries (>= 85% agreement required).",
    f"{len(maint_stats['flagged_for_review'])} left blank - no reliable reference or a genuine split; needs manual entry.",
    ])
    return ws


def _write_changelog_sheet(wb, spelling_counts, categorical_rows, format_fix_counts):
    ws = wb.create_sheet("Change Log - Standardization")
    ws.cell(row=1, column=1, value="Text & Data Standardization Change Log").font = _TITLE_FONT
    r = 3

    ws.cell(row=r, column=1, value="A. Formatting fixes applied to Defects Description").font = _SECTION_FONT
    r += 2
    for c, h in enumerate(["Fix", "Rows / Instances Affected"], start=1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.border = _BORDER
    r += 1
    for label, cnt in format_fix_counts:
        ws.cell(row=r, column=1, value=label).font = _NOTE_FONT
        ws.cell(row=r, column=2, value=cnt).font = _NOTE_FONT
        ws.cell(row=r, column=1).border = _BORDER
        ws.cell(row=r, column=2).border = _BORDER
        r += 1

    r += 2
    ws.cell(row=r, column=1, value="B. Individual spelling corrections").font = _SECTION_FONT
    r += 2
    for c, h in enumerate(["Misspelling", "Corrected To", "Occurrences"], start=1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.border = _BORDER
    r += 1
    for bad, good, cnt in sorted(spelling_counts, key=lambda x: -x[2]):
        ws.cell(row=r, column=1, value=bad).font = _NOTE_FONT
        ws.cell(row=r, column=2, value=good).font = _NOTE_FONT
        ws.cell(row=r, column=3, value=cnt).font = _NOTE_FONT
        for c in range(1, 4):
            ws.cell(row=r, column=c).border = _BORDER
        r += 1

    r += 2
    ws.cell(row=r, column=1, value="C. Categorical field standardization").font = _SECTION_FONT
    r += 2
    for c, h in enumerate(["Column", "Before", "After", "Rows Affected"], start=1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.border = _BORDER
    r += 1
    for col, before, after, cnt in categorical_rows:
        ws.cell(row=r, column=1, value=col).font = _NOTE_FONT
        ws.cell(row=r, column=2, value=before).font = _NOTE_FONT
        ws.cell(row=r, column=3, value=after).font = _NOTE_FONT
        ws.cell(row=r, column=4, value=cnt).font = _NOTE_FONT
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = _BORDER
        r += 1

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 16
    ws.freeze_panes = "A4"
    return ws


def write_workbook(rows, stats, text_stats, malay_count, date_failures,
                    spelling_counts, categorical_rows, format_fix_counts, maint_stats, output_path):
    wb = openpyxl.Workbook()
    _write_master_sheet(wb, rows)
    _write_audit_sheet(wb, rows)
    _write_qa_sheet(wb, stats, text_stats, malay_count, date_failures, maint_stats)
    _write_changelog_sheet(wb, spelling_counts, categorical_rows, format_fix_counts)
    wb.save(output_path)
