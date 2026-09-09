#!/usr/bin/env python3
"""
RETTD Milestone One: search Maine Revenue Services' RETTD Public Search
(https://revenue.maine.gov) for the trial-search criteria and export a
sorted CSV of results.

The site is a FAST Enterprises GenTax app (server-rendered AJAX fragments,
not a JSON data API) — see notes/PROJECT_PLAN.md section 4 for recon
findings. This script drives the real form with Playwright and scrapes the
rendered results table.
"""

import csv
import re
import sys
from playwright.sync_api import sync_playwright

BASE_URL = "https://revenue.maine.gov"

# Trial search criteria (Milestone One §3 of the project plan).
# County/Property Type are selected by their <option> value codes, discovered
# during recon (§4) — County="PENOBSCOT" -> "ME019", Property
# Type="Commercial - Warehouse" -> "311".
COUNTY_CODE = "ME019"          # Penobscot
PROPERTY_TYPE_CODE = "311"     # Commercial - Warehouse
BEGIN_TRANSFER_DATE = "01-Jan-2026"
END_TRANSFER_DATE = "25-Aug-2026"

OUTPUT_CSV = "rettd_trial_search.csv"

# User-requested output fields, in the order they should appear in the CSV.
# DLN is included as a bonus traceability column (flagged as an open
# assumption in the project plan) since some transfers share Map/Lot/Date
# but differ by seller — without DLN they'd be indistinguishable.
OUTPUT_FIELDS = [
    "Municipality",
    "Map",
    "Lot",
    "Seller",
    "Transfer Date",
    "Purchase Price",
    "DLN",
]

# Maps the site's results-table header text to our output field names.
HEADER_TO_FIELD = {
    "Municipality": "Municipality",
    "Map": "Map",
    "Lot": "Lot",
    "Seller": "Seller",
    "Transfer Date": "Transfer Date",
    "Purchase Price": "Purchase Price",
    "DLN": "DLN",
}


def natural_sort_key(value: str):
    """Split into text/number chunks so '7' < '11' and '01' == '1' numerically,
    while non-numeric values (e.g. 'U11', '017-B') still compare sanely."""
    chunks = re.split(r"(\d+)", value or "")
    return tuple(int(c) if c.isdigit() else c.lower() for c in chunks)


def run_search():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        print("Navigating to Maine Tax Portal...", file=sys.stderr)
        page.goto(BASE_URL, timeout=30000, wait_until="networkidle")

        print("Opening RETTD Public Search...", file=sys.stderr)
        page.get_by_text("RETTD Public Search", exact=False).first.click()
        page.wait_for_selector(
            f"#Dd-d option[value='{COUNTY_CODE}']", timeout=20000, state="attached"
        )

        print("Filling search criteria...", file=sys.stderr)
        page.select_option("#Dd-d", COUNTY_CODE)  # County
        page.select_option("#Dd-f", PROPERTY_TYPE_CODE)  # Property Type
        page.fill("#Dd-l", BEGIN_TRANSFER_DATE)  # Begin Transfer Date
        page.fill("#Dd-m", END_TRANSFER_DATE)  # End Transfer Date

        # Echo-check: confirm the date fields actually hold what we typed
        # before submitting — a jQuery UI datepicker widget can silently
        # reformat or clear typed input.
        begin_val = page.eval_on_selector("#Dd-l", "el => el.value")
        end_val = page.eval_on_selector("#Dd-m", "el => el.value")
        assert begin_val == BEGIN_TRANSFER_DATE, f"Begin date echo mismatch: {begin_val!r}"
        assert end_val == END_TRANSFER_DATE, f"End date echo mismatch: {end_val!r}"

        # Blur the date field before clicking Search — clicking immediately
        # after page.fill() on the date widget silently swallows the click
        # (confirmed during recon; the datepicker/watermark plugin needs a
        # blur to commit the field first).
        page.get_by_text("Provide Property Information", exact=True).click()
        page.wait_for_timeout(300)

        print("Submitting search...", file=sys.stderr)
        page.click("#Dd-p", force=True)
        page.wait_for_selector("table tbody tr[data-row]", timeout=20000)
        page.wait_for_load_state("networkidle", timeout=20000)

        # Exhaust pagination: keep clicking "Scroll for More" while it's
        # visible/enabled. Once all results are loaded the link is hidden.
        while True:
            link = page.locator("a.ScrollForMoreLink").first
            if link.count() == 0 or not link.is_visible():
                break
            before = page.locator("table tbody tr[data-row]").count()
            link.click(force=True)
            page.wait_for_timeout(1500)
            page.wait_for_load_state("networkidle", timeout=20000)
            after = page.locator("table tbody tr[data-row]").count()
            print(f"Scroll for More: {before} -> {after} rows", file=sys.stderr)
            if after == before:
                break

        rows = extract_results(page)
        browser.close()
        return rows


def extract_results(page):
    # The page renders a second "clone" header (for a sticky-header effect)
    # alongside the real results table, so a bare "table thead th" selector
    # matches two header sets. Resolve everything relative to the *specific*
    # table that owns the data rows (identified by tbody tr[data-row]),
    # entirely in-browser, so header and body always come from the same table.
    data = page.evaluate(
        """() => {
            const dataRow = document.querySelector('tbody tr[data-row]');
            if (!dataRow) return null;
            const table = dataRow.closest('table');
            const headers = Array.from(table.querySelectorAll('thead th'))
                .map(th => th.innerText.trim());
            const rows = Array.from(table.querySelectorAll('tbody tr[data-row]')).map(tr =>
                Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim())
            );
            return {headers, rows};
        }"""
    )
    if not data:
        raise RuntimeError("No results table found on the page.")

    header_map = {h: i for i, h in enumerate(data["headers"]) if h}
    missing = [h for h in HEADER_TO_FIELD if h not in header_map]
    if missing:
        raise RuntimeError(f"Expected result columns not found: {missing}")

    print(f"Extracting {len(data['rows'])} result rows...", file=sys.stderr)

    rows = []
    for cells in data["rows"]:
        record = {
            field: cells[header_map[header]]
            for header, field in HEADER_TO_FIELD.items()
        }
        rows.append(record)

    return rows


def write_csv(rows, path):
    rows_sorted = sorted(
        rows,
        key=lambda r: (
            natural_sort_key(r["Municipality"]),
            natural_sort_key(r["Map"]),
            natural_sort_key(r["Lot"]),
        ),
    )
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_FIELDS)
        writer.writeheader()
        for r in rows_sorted:
            writer.writerow(r)
    return rows_sorted


def main():
    rows = run_search()
    if not rows:
        print("WARNING: search returned zero rows.", file=sys.stderr)
    rows_sorted = write_csv(rows, OUTPUT_CSV)
    print(f"Wrote {len(rows_sorted)} rows to {OUTPUT_CSV}", file=sys.stderr)


if __name__ == "__main__":
    main()
