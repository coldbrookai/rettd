#!/usr/bin/env python3
"""
RETTD search engine: drives Maine Revenue Services' RETTD Public Search
(https://revenue.maine.gov) with Playwright and scrapes the rendered
results table.

The site is a FAST Enterprises GenTax app (server-rendered AJAX fragments,
not a JSON data API) — see notes/PROJECT_PLAN.md section 2 for recon
findings. Running this file directly reproduces Milestone One's trial
search (unchanged behavior, still writes rettd_trial_search.csv). The
`search_m2()` entry point is Milestone Two's generalized search, used by
rettd_app.py — see notes/PROJECT_PLAN.md section 6. Milestone Three's xlsx
export lives in rettd_export.py, not here (section 7.3) — this module
imports natural_sort_key()/sort_rows() from there rather than the reverse,
so rettd_export.py never has to import Playwright.
"""

import csv
import json
import re
import sys
from datetime import date, timedelta

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from rettd_export import FIELD_KEYS as M2_FIELDS
from rettd_export import natural_sort_key, sort_rows  # noqa: F401 (re-exported)

BASE_URL = "https://revenue.maine.gov"

MONTH_ABBR = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]

# Trial search criteria (Milestone One section 3 of the project plan).
# County/Property Type are selected by their <option> value codes, discovered
# during recon (section 2) — County="PENOBSCOT" -> "ME019", Property
# Type="Commercial - Warehouse" -> "311".
COUNTY_CODE = "ME019"          # Penobscot
PROPERTY_TYPE_CODE = "311"     # Commercial - Warehouse

# State-wide search sentinel (Milestone Two, added 2026-09-17): the site's
# County <select> has no all-counties option of its own, but the site
# accepts a search with County left unselected and returns matches across
# all counties (confirmed live) -- search_m2() maps this sentinel to
# skipping County selection entirely, not to looping over every county.
STATEWIDE_COUNTY_CODE = "STATEWIDE"
BEGIN_TRANSFER_DATE = date(2026, 1, 1)
END_TRANSFER_DATE = date(2026, 8, 25)

OUTPUT_CSV = "rettd_trial_search.csv"

# User-requested output fields, in the order they should appear in the CSV.
# DLN is included as a bonus traceability column since some transfers share
# Map/Lot/Date but differ by seller — without DLN they'd be indistinguishable.
OUTPUT_FIELDS = [
    "Municipality",
    "Map",
    "Lot",
    "Seller",
    "Transfer Date",
    "Purchase Price",
    "DLN",
]

# M2_FIELDS (Buyer added to Milestone One's 7 — section 6.4) and the sort
# helpers below are imported from rettd_export.py, not defined here.

# Both lists map 1:1 to the site's own results-table header text, so these
# are identity maps (site header text == our field name in every case).
HEADER_TO_FIELD = {f: f for f in OUTPUT_FIELDS}
HEADER_TO_FIELD_M2 = {f: f for f in M2_FIELDS}

# Milestone Two's fixed price bounds (not user-editable — section 6.2).
M2_MIN_PRICE = 400000.00
M2_MAX_PRICE = 10000000.00

LOOK_BACK_DAYS = {"30d": 30, "60d": 60, "90d": 90, "180d": 180}
LOOK_BACK_YEARS = {"1y": 1, "5y": 5, "10y": 10}


def format_site_date(d: date) -> str:
    """The site's date fields expect 'DD-Mon-YYYY' (e.g. '05-Sep-2026').
    Uses an explicit month-abbreviation table rather than strftime('%b'),
    which is locale-dependent (see notes/PROJECT_PLAN.md section 7.5 for the
    same concern in the Milestone Three export)."""
    return f"{d.day:02d}-{MONTH_ABBR[d.month - 1]}-{d.year}"


def compute_begin_date(today: date, look_back: str) -> date:
    """Milestone Two's Look-Back Period control (section 6.2) sets Begin
    Transfer Date to today minus the chosen period."""
    if look_back in LOOK_BACK_DAYS:
        return today - timedelta(days=LOOK_BACK_DAYS[look_back])
    if look_back in LOOK_BACK_YEARS:
        years = LOOK_BACK_YEARS[look_back]
        try:
            return today.replace(year=today.year - years)
        except ValueError:
            # today is Feb 29 and (today.year - years) isn't a leap year.
            return today.replace(month=2, day=28, year=today.year - years)
    raise ValueError(f"Unknown look-back period: {look_back!r}")


def resolve_labels(page):
    """Map each form field's visible label text (via aria-labelledby) to its
    current element id. Building this fresh on every run means field ID
    drift (notes/PROJECT_PLAN.md section 2, "Field ID drift") no longer
    breaks the script — only a label-text change would, per section 6.3."""
    return page.evaluate(
        """() => {
            const map = {};
            document.querySelectorAll('[aria-labelledby]').forEach(el => {
                const text = el.getAttribute('aria-labelledby')
                    .split(/\\s+/)
                    .map(id => {
                        const n = document.getElementById(id);
                        return n ? n.innerText.trim() : '';
                    })
                    .filter(Boolean)
                    .join(' ');
                if (text && !(text in map) && el.id) map[text] = el.id;
            });
            return map;
        }"""
    )


def find_search_button(page):
    """The GenTax form renders a hidden duplicate 'Search' button in an
    off-screen panel alongside the real one — the same duplication pattern
    documented for the results table's sticky header (section 2). Find the
    one with an actual on-screen position rather than relying on an id,
    which has drifted before."""
    buttons = page.locator("button", has_text="Search")
    for i in range(buttons.count()):
        candidate = buttons.nth(i)
        box = candidate.bounding_box()
        if box and box["width"] > 0 and box["height"] > 0:
            return candidate
    raise RuntimeError("Could not find an on-screen Search button.")


_PAGE_OF_RE = re.compile(r"Page (\d+) of (\d+)")


def _pager_state(page):
    """Reads the results table's "Page X of Y" indicator. Returns (1, 1)
    when absent -- a result set of 50 rows or fewer has no pager at all
    (confirmed live, section 9's pagination note)."""
    text = page.evaluate(
        """() => {
            const el = Array.from(document.querySelectorAll('*')).find(
                e => e.innerText && /^Page \\d+ of \\d+$/.test(e.innerText.trim())
            );
            return el ? el.innerText.trim() : null;
        }"""
    )
    if not text:
        return 1, 1
    match = _PAGE_OF_RE.match(text)
    return int(match.group(1)), int(match.group(2))


def _click_next_page(page):
    """Advances the results table to its next page via the real Prev/Next
    pager (`a.TablePageLinkNext` / "Page X of Y", section 9) — confirmed
    live (2026-09-17) to be the site's actual multi-page mechanism, unlike
    the `a.ScrollForMoreLink` this scraper previously (and incorrectly)
    treated as pagination; that link stayed inert in every query tested,
    including ones that turned out to have hundreds of matching rows spread
    across many pages. `force=True` (used elsewhere in this file, e.g.
    find_search_button()) breaks this specific click — it skips the
    hit-testing/actionability check this pager's click handler apparently
    depends on — so this scrolls the link into view and clicks normally."""
    next_link = page.locator("a.TablePageLinkNext").first
    next_link.scroll_into_view_if_needed()
    next_link.click()


def open_datepicker_today(page, date_field_id):
    """Open a date field's picker via its 'Toggle Date Picker' sibling
    button and click 'Today' (confirmed working, section 2's "Date-picker
    'Today' button" note)."""
    page.locator(f"#ic_{date_field_id} button.ui-datepicker-trigger").click()
    page.wait_for_selector("#ui-datepicker-div", state="visible", timeout=5000)
    page.click("#ui-datepicker-div .ui-datepicker-current")
    page.wait_for_timeout(300)


def open_search_form(page, log=lambda msg: None):
    log("Navigating to Maine Tax Portal...")
    try:
        page.goto(BASE_URL, timeout=30000, wait_until="networkidle")
    except PlaywrightTimeoutError:
        raise RuntimeError(
            "Timed out reaching revenue.maine.gov — check your internet connection "
            "and try again."
        )
    except Exception as exc:
        # Playwright's own connection-failure errors (DNS, refused, etc.) are
        # not PlaywrightTimeoutError but shouldn't reach the user as a raw
        # driver exception either.
        raise RuntimeError(f"Could not reach revenue.maine.gov: {exc}")

    log("Opening RETTD Public Search...")
    try:
        page.get_by_text("RETTD Public Search", exact=False).first.click()
        page.wait_for_selector("select", timeout=20000, state="attached")
    except PlaywrightTimeoutError:
        raise RuntimeError(
            "Could not open RETTD Public Search — the site's page structure may "
            "have changed (see PROJECT_PLAN.md section 9)."
        )
    page.wait_for_timeout(300)


def _wait_for_municipality_cascade(page, muni_id, timeout=8000):
    """The site's Municipality combobox is populated by the site's own JS in
    response to a County change, but only if County received a real click
    (focus) before its value changed — Playwright's select_option() alone
    silently skips this (confirmed 2026-09-12, section 6.2). Waits for the
    cascade to actually land rather than assuming a fixed delay is enough."""
    page.wait_for_function(
        f"document.getElementById({json.dumps(muni_id)}).options.length > 1",
        timeout=timeout,
    )


def get_municipalities(county_code, log=lambda msg: None):
    """Fetch the live Municipality/Township list for a county, for populating
    the input page's cascading dropdown (section 6.2). Returns a list of
    {"value": ..., "label": ...} dicts (blank placeholder excluded). Opens
    its own short-lived browser session, like search_m2()."""
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        open_search_form(page, log)
        labels = resolve_labels(page)
        county_id = labels.get("County")
        muni_id = labels.get("Municipality/Township")
        if not county_id or not muni_id:
            browser.close()
            raise RuntimeError(
                "Could not find County/Municipality fields on the form — "
                "the site's form may have changed (see PROJECT_PLAN.md section 9)."
            )
        page.click(f"#{county_id}")
        page.select_option(f"#{county_id}", county_code)
        _wait_for_municipality_cascade(page, muni_id)
        options = page.eval_on_selector_all(
            f"#{muni_id} option",
            "els => els.map(o => ({value: o.value, label: o.textContent.trim()}))",
        )
        browser.close()
        return [o for o in options if o["value"]]


def execute_search(
    page,
    labels,
    county_code,
    property_type_code,
    header_to_field,
    municipality_value=None,
    min_price=None,
    max_price=None,
    begin_date=None,
    end_date=None,
    log=lambda msg: None,
):
    """Fill and submit the search form, then extract every result row across
    every page (see the Prev/Next pager handling below). `begin_date`/
    `end_date` are `date` objects; if `end_date` is None, the End Transfer
    Date field is set via the datepicker's Today button instead of being
    typed (Milestone Two behavior, section 6.2). `municipality_value` is an
    option value from get_municipalities(), not a free-text name.
    `county_code=None` leaves the County field unselected — confirmed live
    (2026-09-17) that the site itself accepts this and returns matches
    across all counties, contrary to this file's earlier assumption (from
    PROJECT_PLAN.md section 6.2's "County: Required") that County was
    site-mandated; that was a Milestone Two *UI* decision, not a site
    constraint. `header_to_field` is passed straight through to
    extract_results() for each page.

    Returns (rows, municipality_warning). `rows` is `[]` if the site
    reported no matching rows (no crash — section 6.5 flags exact
    zero-results presentation as still open, but this at least fails soft).
    `municipality_warning` is a user-facing string if the optional
    Municipality filter couldn't be applied, else None.
    """

    def fid(label):
        if label not in labels:
            raise RuntimeError(
                f"Could not find a form field labelled {label!r} — "
                "the site's form may have changed (see PROJECT_PLAN.md section 9)."
            )
        return labels[label]

    log("Filling search criteria...")
    if county_code:
        county_id = fid("County")
        # Must click (focus) before select_option() — see
        # _wait_for_municipality_cascade()'s docstring.
        page.click(f"#{county_id}")
        page.select_option(f"#{county_id}", county_code)
    page.select_option(f"#{fid('Property Type')}", property_type_code)

    municipality_warning = None
    if municipality_value:
        muni_id = labels.get("Municipality/Township")
        if muni_id:
            try:
                _wait_for_municipality_cascade(page, muni_id)
                page.select_option(f"#{muni_id}", value=municipality_value, timeout=3000)
            except Exception:
                municipality_warning = (
                    "Could not apply the Municipality filter for this county — "
                    "the search ran without it."
                )
        else:
            municipality_warning = "Municipality field not found on the form — search ran without it."

    if min_price is not None:
        min_id = fid("Minimum Purchase Price")
        min_str = f"{min_price:.2f}"
        page.fill(f"#{min_id}", min_str)
        got = page.eval_on_selector(f"#{min_id}", "el => el.value")
        assert got == min_str, f"Min price echo mismatch: {got!r}"
    if max_price is not None:
        max_id = fid("Maximum Purchase Price")
        max_str = f"{max_price:.2f}"
        page.fill(f"#{max_id}", max_str)
        got = page.eval_on_selector(f"#{max_id}", "el => el.value")
        assert got == max_str, f"Max price echo mismatch: {got!r}"

    begin_id = fid("Begin Transfer Date")
    begin_str = format_site_date(begin_date)
    page.fill(f"#{begin_id}", begin_str)
    got = page.eval_on_selector(f"#{begin_id}", "el => el.value")
    assert got == begin_str, f"Begin date echo mismatch: {got!r}"

    end_id = fid("End Transfer Date")
    if end_date is None:
        log("Setting End Transfer Date to today...")
        open_datepicker_today(page, end_id)
        expected_end = format_site_date(date.today())
    else:
        expected_end = format_site_date(end_date)
        page.fill(f"#{end_id}", expected_end)
    got = page.eval_on_selector(f"#{end_id}", "el => el.value")
    assert got == expected_end, f"End date echo mismatch: {got!r}"

    # Blur the date field before clicking Search — clicking immediately
    # after filling a date widget silently swallows the click (section 2's
    # "Gotcha"). Click a neutral heading element to commit focus away.
    page.get_by_text("Provide Property Information", exact=True).click()
    page.wait_for_timeout(300)

    log("Submitting search...")
    find_search_button(page).click(force=True)

    try:
        # 15s was measured (2026-09-12 stress test) to intermittently be too
        # short for heavy/broad queries — one legitimate load took 17s —
        # which caused false "no results" on searches that actually had
        # matches. 30s gives real headroom (see PROJECT_PLAN.md section 9).
        page.wait_for_selector("table tbody tr[data-row]", timeout=30000)
    except PlaywrightTimeoutError:
        log("No results found.")
        return [], municipality_warning

    page.wait_for_load_state("networkidle", timeout=20000)

    rows = extract_results(page, header_to_field)
    current_page, total_pages = _pager_state(page)
    while current_page < total_pages:
        log(f"Loading page {current_page + 1} of {total_pages}...")
        _click_next_page(page)
        page.wait_for_timeout(500)
        page.wait_for_load_state("networkidle", timeout=20000)
        rows.extend(extract_results(page, header_to_field))
        new_page, total_pages = _pager_state(page)
        if new_page <= current_page:
            log("Page number did not advance after clicking Next — stopping "
                "with what's loaded so far rather than looping forever.")
            break
        current_page = new_page

    return rows, municipality_warning


def extract_results(page, header_to_field):
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
    missing = [h for h in header_to_field if h not in header_map]
    if missing:
        raise RuntimeError(f"Expected result columns not found: {missing}")

    rows = []
    for cells in data["rows"]:
        record = {
            field: cells[header_map[header]]
            for header, field in header_to_field.items()
        }
        rows.append(record)

    return rows


def run_search():
    """Milestone One trial search — unchanged output/behavior."""
    log = lambda msg: print(msg, file=sys.stderr)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        open_search_form(page, log)
        labels = resolve_labels(page)
        rows, _ = execute_search(
            page,
            labels,
            county_code=COUNTY_CODE,
            property_type_code=PROPERTY_TYPE_CODE,
            header_to_field=HEADER_TO_FIELD,
            begin_date=BEGIN_TRANSFER_DATE,
            end_date=END_TRANSFER_DATE,
            log=log,
        )
        log(f"Extracted {len(rows)} result rows...")
        browser.close()
        return rows


def search_m2(property_type_code, county_code, look_back, municipality_value=None, log=lambda msg: None):
    """Milestone Two's generalized search (notes/PROJECT_PLAN.md section 6):
    fixed price bounds, Begin Transfer Date computed from `look_back`, End
    Transfer Date always today. `municipality_value` is an option value from
    get_municipalities(), not a free-text name. Returns
    (rows_sorted, municipality_warning); rows include the Buyer column
    (M2_FIELDS).

    `county_code == STATEWIDE_COUNTY_CODE` runs a single search with County
    left unselected (`execute_search(county_code=None, ...)`), rather than
    one search per county — confirmed live (2026-09-17, after an earlier
    16-county-loop version proved both wrong and painfully slow) that the
    site itself accepts a blank County and returns matches across all
    counties in the same ~1-20s as a single-county search. Municipality has
    no meaning without a County selected, so `municipality_value` is
    ignored in that case."""
    is_statewide = county_code == STATEWIDE_COUNTY_CODE
    site_county_code = None if is_statewide else county_code
    muni_value = None if is_statewide else municipality_value
    begin_date = compute_begin_date(date.today(), look_back)

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()

        open_search_form(page, log)
        labels = resolve_labels(page)
        rows, municipality_warning = execute_search(
            page,
            labels,
            county_code=site_county_code,
            property_type_code=property_type_code,
            header_to_field=HEADER_TO_FIELD_M2,
            municipality_value=muni_value,
            min_price=M2_MIN_PRICE,
            max_price=M2_MAX_PRICE,
            begin_date=begin_date,
            end_date=None,
            log=log,
        )
        browser.close()

    return sort_rows(rows), municipality_warning


def write_csv(rows, path):
    rows_sorted = sort_rows(rows)
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
