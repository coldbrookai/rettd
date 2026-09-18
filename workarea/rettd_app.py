#!/usr/bin/env python3
"""
RETTD desktop app: a pywebview window around ../index.html that runs
Milestone Two searches through rettd_search.search_m2() and Milestone
Three's export through rettd_export.build_workbook(), returning results to
the page over the JS<->Python bridge.

See notes/PROJECT_PLAN.md sections 6 and 7 for the design this implements.
"""

import sys
from datetime import date
from pathlib import Path

import webview

import rettd_export as export
import rettd_search as core

INDEX_HTML = Path(__file__).resolve().parent.parent / "index.html"


def _suggested_filename(criteria):
    parts = [criteria.get("property_type_label", ""), criteria.get("county_label", "")]
    base = "_".join(p.replace(" ", "-") for p in parts if p) or "RETTD-Export"
    return f"{base}_{date.today():%Y-%m-%d}.xlsx"


class Api:
    def __init__(self):
        self.window = None
        self._last_rows_by_dln = {}

    def get_municipalities(self, county_code):
        try:
            options = core.get_municipalities(
                county_code, log=lambda msg: print(msg, file=sys.stderr)
            )
        except Exception as exc:
            return {"ok": False, "error": str(exc)}
        return {"ok": True, "options": options}

    def search(self, criteria):
        try:
            rows, warning = core.search_m2(
                property_type_code=criteria["property_type_code"],
                county_code=criteria["county_code"],
                look_back=criteria["look_back"],
                municipality_value=criteria.get("municipality_value"),
                log=lambda msg: print(msg, file=sys.stderr),
            )
        except Exception as exc:
            return {"ok": False, "error": str(exc)}
        # Retained so export_selected() can look rows up by DLN without a
        # second round trip of full row data through the DOM (section 7.6).
        self._last_rows_by_dln = {r["DLN"]: r for r in rows}
        return {"ok": True, "rows": rows, "warning": warning}

    def export_selected(self, payload):
        dlns = payload.get("dlns") or []
        criteria = payload.get("criteria") or {}
        selected_rows = [
            self._last_rows_by_dln[d] for d in dlns if d in self._last_rows_by_dln
        ]
        if not selected_rows:
            return {"ok": False, "error": "No matching rows to export — try searching again."}

        try:
            workbook = export.build_workbook(selected_rows, criteria)
        except Exception as exc:
            return {"ok": False, "error": f"Could not prepare the export: {exc}"}

        result = self.window.create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename=_suggested_filename(criteria),
            file_types=("Excel Workbook (*.xlsx)",),
        )
        if not result:
            # Cancel is a silent no-op, not an error (section 7.6 step 5).
            return {"ok": True, "cancelled": True}
        path = result[0] if isinstance(result, (list, tuple)) else result
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        try:
            workbook.save(path)
        except Exception as exc:
            return {"ok": False, "error": f"Could not save the file: {exc}"}

        return {"ok": True, "cancelled": False, "path": path, "count": len(selected_rows)}


def main():
    api = Api()
    window = webview.create_window(
        "RETTD — Commercial Sale Leads",
        url=str(INDEX_HTML),
        js_api=api,
        width=1100,
        height=800,
        min_size=(700, 500),
        # The results table and output header (section 6.4) are laid out
        # for a full desktop-sized window -- start maximized rather than at
        # the width/height above, which are only a fallback for platforms
        # where maximized on launch isn't honored.
        maximized=True,
    )
    api.window = window
    webview.start()


if __name__ == "__main__":
    main()
