#!/usr/bin/env python3
"""
Milestone Three: builds a formatted .xlsx export of user-selected RETTD
search results, matching resources/NewOutputLayout.png's letterhead-style
layout. See notes/PROJECT_PLAN.md section 7 for the full design and the
decisions already recorded there (local .xlsx only via openpyxl; no Google
Sheets/cloud API, no Excel COM automation).

Deliberately has no Playwright import — rettd_search.py imports
natural_sort_key()/sort_rows() from here (not the other way around) so this
module, and anything that only needs to format/export data, never drags in
browser automation as a dependency.
"""

import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Underlying data field names, in export column order (matches the on-screen
# table's 8 columns, section 6.4 -- Buyer added to Milestone One's original
# 7). rettd_search.py imports this as M2_FIELDS.
FIELD_KEYS = [
    "Municipality",
    "Map",
    "Lot",
    "Seller",
    "Buyer",
    "Transfer Date",
    "Purchase Price",
    "DLN",
]

# Display labels for the exported sheet's header row. Deliberately differs
# from the on-screen table, which keeps "Transfer Date"/"Purchase Price"
# (section 6.4) for data-model consistency -- this exported sheet *is* the
# print-shaped artifact NewOutputLayout.png describes, which uses "Sale
# Date"/"Sale Price" (section 7.3's display-label decision).
SHEET_HEADERS = {
    "Municipality": "Municipality",
    "Map": "Map",
    "Lot": "Lot",
    "Seller": "Seller",
    "Buyer": "Buyer",
    "Transfer Date": "Sale Date",
    "Purchase Price": "Sale Price",
    "DLN": "DLN",
}

COLUMN_WIDTHS = {
    "Municipality": 16,
    "Map": 9,
    "Lot": 9,
    "Seller": 30,
    "Buyer": 30,
    "Transfer Date": 13,
    "Purchase Price": 15,
    "DLN": 15,
}

_MONTH_NUM = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
    "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}

HEADER_FILL = PatternFill(start_color="2F5D3A", end_color="2F5D3A", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
_THIN = Side(style="thin", color="000000")
THIN_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

HEADER_ROW = 9


def natural_sort_key(value: str):
    """Split into text/number chunks so '7' < '11' and '01' == '1' numerically,
    while non-numeric values (e.g. 'U11', '017-B') still compare sanely."""
    chunks = re.split(r"(\d+)", value or "")
    return tuple(int(c) if c.isdigit() else c.lower() for c in chunks)


def sort_rows(rows):
    """Natural sort by (Municipality, Map, Lot) -- the order established in
    Milestone One (section 5.4), shared by the CSV path, the on-screen
    table, and this export."""
    return sorted(
        rows,
        key=lambda r: (
            natural_sort_key(r["Municipality"]),
            natural_sort_key(r["Map"]),
            natural_sort_key(r["Lot"]),
        ),
    )


def _parse_price(value):
    """Parse a '1,186,919.00' / '0.00' style string into a float. Returns
    None if unparseable -- callers must not treat that as "skip the cell",
    since 0.0 is a real, meaningful value (section 7.5)."""
    if value is None:
        return None
    cleaned = str(value).replace("$", "").replace(",", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        return None


def _parse_transfer_date(value):
    """Parse the site's 'DD-Mon-YYYY' format via an explicit month table,
    not locale-dependent strptime('%b') (section 7.5)."""
    match = re.match(r"^(\d{1,2})-([A-Za-z]{3})-(\d{4})$", (value or "").strip())
    if not match:
        return None
    day, mon, year = match.groups()
    month = _MONTH_NUM.get(mon.title())
    if not month:
        return None
    try:
        return datetime(int(year), month, int(day))
    except ValueError:
        return None


def build_workbook(selected_rows, criteria):
    """Build (but don't save) the export workbook. `criteria` is a dict with
    optional 'property_type_label'/'county_label'/'municipality_label'
    display strings (section 6.4's criteria echo). `municipality_label`
    defaults to "All Municipalities" if omitted, matching the on-screen
    dropdown's own default option when no specific municipality is chosen."""
    rows_sorted = sort_rows(selected_rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "Commercial Sale Leads"

    last_col = get_column_letter(len(FIELD_KEYS))

    # Empty spacer row above the title (previously held the now-removed logo
    # image, section 7.4 update 2026-09-17) -- kept merged for a consistent
    # cell grid with the title/subtitle rows below it.
    ws.merge_cells(f"A1:{last_col}1")

    ws.merge_cells(f"A2:{last_col}2")
    ws["A2"] = "Commercial Sale Leads"
    ws["A2"].font = Font(size=18, bold=True)
    ws["A2"].alignment = Alignment(horizontal="center")

    ws.merge_cells(f"A3:{last_col}3")
    ws["A3"] = "Maine Department of Revenue - Real Estate Transfer Tax Database"
    ws["A3"].font = Font(size=14, bold=True, color="000000")
    ws["A3"].alignment = Alignment(horizontal="center")

    # Labels stay in column A ("reference categories"); values sit one cell
    # right in column B so the three values line up in their own vertical
    # column, separated from the labels (section 7.4 update 2026-09-17).
    criteria_rows = (
        ("A5", "B5", "Property Type:", criteria.get("property_type_label", "")),
        ("A6", "B6", "County:", criteria.get("county_label", "")),
        ("A7", "B7", "Municipality:", criteria.get("municipality_label", "All Municipalities")),
    )
    for label_ref, value_ref, label, value in criteria_rows:
        ws[label_ref] = label
        ws[label_ref].font = Font(size=10)
        ws[value_ref] = value
        ws[value_ref].font = Font(size=10)

    for col_idx, field in enumerate(FIELD_KEYS, start=1):
        cell = ws.cell(row=HEADER_ROW, column=col_idx, value=SHEET_HEADERS[field])
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = THIN_BORDER
        # Right-aligned to match the (numeric/date-valued, so already
        # right-aligned by Excel's default) data cells beneath them.
        right_aligned = field in ("DLN", "Transfer Date", "Purchase Price")
        cell.alignment = Alignment(horizontal="right" if right_aligned else "left")

    for row_offset, row in enumerate(rows_sorted, start=1):
        r = HEADER_ROW + row_offset
        for col_idx, field in enumerate(FIELD_KEYS, start=1):
            raw = row.get(field, "")
            cell = ws.cell(row=r, column=col_idx)
            cell.border = THIN_BORDER
            if field == "Purchase Price":
                price = _parse_price(raw)
                if price is not None:
                    cell.value = price
                    cell.number_format = "$#,##0"
                else:
                    cell.value = raw
            elif field == "Transfer Date":
                dt = _parse_transfer_date(raw)
                if dt is not None:
                    cell.value = dt
                    cell.number_format = "DD-MMM-YYYY"
                else:
                    cell.value = raw
            elif field in ("Map", "Lot", "DLN"):
                # Forced text -- otherwise Excel silently strips leading
                # zeros / coerces to a number (section 7.5, section 9).
                cell.value = str(raw)
                cell.number_format = "@"
                if field == "DLN":
                    cell.alignment = Alignment(horizontal="right")
            else:
                cell.value = raw

    for col_idx, field in enumerate(FIELD_KEYS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = COLUMN_WIDTHS[field]

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.freeze_panes = f"A{HEADER_ROW + 1}"
    ws.print_title_rows = f"{HEADER_ROW}:{HEADER_ROW}"

    return wb


def export_to_xlsx(selected_rows, criteria, path):
    build_workbook(selected_rows, criteria).save(path)
