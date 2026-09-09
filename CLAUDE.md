# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

Tooling to search, collect, and present Maine real estate transfer records from
Maine Revenue Services' RETTD Public Search (revenue.maine.gov) in a more usable
form than the stock government search UI. First proven use case: commercial real
estate transfer activity, for real estate professionals tracking recent transfers.

Early stage, single-user prototype. The only working code is a Playwright scraper
(`workarea/rettd_search.py`); no UI exists yet. See `notes/PROJECT_PLAN.md`
for the full milestone plan (Milestone One — search/retrieve — is complete;
Milestone Two — a local pywebview desktop landing page with an input form and
results table, replacing an earlier dashboard concept — is in planning;
Milestone Three — downloadable results file — and Milestone Four — hardening —
are not started). `rettd_search.py`'s field IDs were found stale (2026-09-09)
and have been refreshed — see `notes/PROJECT_PLAN.md` §2's "Field ID drift"
note for the current mapping and why this can recur.

## Commands

```bash
cd workarea
python3 -m venv venv                 # one-time setup
source venv/bin/activate
pip install -r requirements.txt      # installs playwright
playwright install chromium          # Playwright's sandboxed browser, not a system package

python rettd_search.py               # runs the trial search, writes rettd_trial_search.csv
```

There is no test suite, linter, or build step — this is a single scraper script,
run manually and verified against `resources/searchOutput.png` (a known-correct
result screenshot for the hardcoded trial-search criteria).

## Repo layout and what's actually tracked in git

Only `index.html` is committed to git so far. `.gitignore` excludes `venv/`,
`notes/`, `resources/`, `*.csv`, and `README.*` — so `README.md`,
`notes/PROJECT_PLAN.md`, and everything under `resources/` (design references,
screenshots) exist on disk but are intentionally not version-controlled. Don't
assume `git status`/`git log` reflects the project's real state or history;
read those files directly instead.

- `workarea/` — the actual Python/Playwright code and its venv.
- `notes/PROJECT_PLAN.md` — authoritative record of site recon findings, the
  milestone roadmap, and design decisions already made (and why). Read this
  before changing scraping approach or re-investigating something it already
  answers.
- `resources/MVC_DESIGN.md` — the green/cream color palette, typography, and
  component spec to use for any future UI work.
- `resources/NewDataEntry.png` / `resources/NewOutputLayout.png` (sourced
  from `resources/SampleForms.ods`) — the Milestone Two input-form and
  results-output layout mockups; `resources/MVClgo01.png` — the logo lockup
  used on the output page.
- `resources/ScreenshotRETTDsearch.png`, `ScreenshotTargetDB.png`,
  `ScreenshotResults.png`, `searchOutput.png` — screenshots of the target
  site used during Milestone One recon.

## Architecture of the target site (why the scraper works the way it does)

RETTD is a **FAST Enterprises GenTax** application (identifiable by the `/_/`
URL segment and `/_/EventOccurred` AJAX endpoint) — a proprietary server-driven
UI framework, not Angular/React. This shapes the whole scraping approach:

- **No JSON data API.** Every interaction (dropdown change, Search click) POSTs
  to `/_/EventOccurred` and gets back a JSON envelope whose `html` field is a
  full re-rendered HTML fragment, including results as a real `<table>`.
  Replicating this protocol directly would be far more fragile than driving a
  real browser — hence Playwright (Chromium, headless), not an HTTP client.
- Form fields are genuine, accessible HTML (`<select>` for County/Municipality/
  Property Type with stable `<option value>` codes; date fields are plain text
  inputs enhanced with a jQuery UI datepicker) — `select_option()`/`fill()`
  work directly, no custom-widget handling needed.
- **Date field gotcha:** clicking Search immediately after `fill()`-ing a date
  field silently does nothing — the datepicker plugin needs a blur event to
  commit the typed value first. The script clicks a neutral heading element
  before clicking Search, and echo-checks the field values before submitting.
- **Pagination** ("Scroll for More") is real but client-triggered — the script
  loops clicking it while visible, tracking row count before/after, and stops
  when the link hides or the count stops changing. This path is implemented
  but not yet exercised at scale (only 7-row result sets tested so far).
- **Results extraction is header-driven, not position-driven:** the DOM has two
  header regions (a sticky-header clone), so results are resolved from the
  specific `<table>` that owns `tbody tr[data-row]`, and columns are read by
  mapping header text → index — never by hardcoded column position. This
  matters because column order/count could shift without notice.
- The scraper is coupled to current GenTax markup (form field IDs, and exact
  results-table header text). A Maine Revenue Services site update could
  break it silently — there's no versioned API contract, and this has
  already happened once: the form field IDs `rettd_search.py` used
  (`#Dd-k`/`#Dd-l` for Begin/End Transfer Date, `#Dd-o` for Search) drifted
  within about two weeks of Milestone One (a new DLN input field was added,
  shifting every ID after it). The script now uses the current IDs
  (`#Dd-l`/`#Dd-m` for Begin/End Transfer Date, `#Dd-p` for Search — full
  mapping in `notes/PROJECT_PLAN.md` §2) but nothing prevents this from
  recurring. Prefer resolving fields by their `aria-labelledby` label text
  over hardcoding `Dd-*` IDs going forward (planned for Milestone Two).

## Key domain notes

- Map/Lot values are alphanumeric and inconsistently zero-padded (e.g. Map:
  `U11`, `139`, `R18`; Lot: `A`, `04`, `017-B`). Sort with `natural_sort_key()`
  in `rettd_search.py` (splits digit/non-digit chunks, coerces digits to
  `int`) — a plain string sort produces wrong ordering.
  Output rows keep the original string values verbatim; only the sort key is
  transformed.
- `DLN` is carried through as an output column because Map/Lot/Transfer Date
  alone can collide across distinct transfers (confirmed case: two same-day
  same-map/lot rows with different sellers/prices) — it's the disambiguator.
- `0.00` purchase price is a real value (non-arm's-length transfer) and must
  not be filtered out.
- County/Property Type option codes and date formatting are isolated at the
  top of `rettd_search.py`, which will help when Milestone Two wires up real
  user inputs — but Milestone Two also changes several of Milestone One's
  fixed assumptions (Property Type/County become mandatory and
  category-filtered, price bounds become fixed values, End Transfer Date is
  set by clicking the date picker's "Today" button instead of being typed),
  so it's not pure wiring. See `notes/PROJECT_PLAN.md` §6 for the full spec.
