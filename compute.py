"""
Cambri Demand Plan — GitHub Actions compute script
----------------------------------------------------
Downloads the SharePoint Excel file, reads the manually-maintained
'Actuals' summary sheet, and writes data.json to the repo root.

Run:
    pip install requests openpyxl
    python compute.py
"""

import datetime
import json
from io import BytesIO

import requests
from openpyxl import load_workbook

# ── CONFIG ────────────────────────────────────────────────────────────────────
SHAREPOINT_URL = (
    "https://wiseandsallycom-my.sharepoint.com/:x:/g/personal/"
    "csongor_doma_cambri_io/"
    "IQDKVYMS8mOJS5WxpQi7eylEAWeqX1IwmUSiOAj2kDGFOns"
    "?e=Zh70hH&download=1"
)
TARGET_YEAR = 2026
OUTPUT_FILE = "data.json"

# ── FIXED ANNUAL TARGETS ──────────────────────────────────────────────────────
TARGETS = {
    "marketing": {"s1": 130,  "s2": 104, "poc": 36.4, "wins": 18.2, "value": 2275000},
    "bdr":       {"s1": 313,  "s2": 72,  "poc": 19.4, "wins": 9.7,  "value": 730000 },
    "partner":   {"s1": 50,   "s2": 22,  "poc": 11,   "wins": 5.5,  "value": 550000 },
}

# Maps values found in the Source and Metric columns of the Actuals sheet
SRC_MAP = {
    "marketing":  "marketing",
    "bdr":        "bdr",
    "partner":    "partner",
}

METRIC_MAP = {
    "s1":                   "s1",
    "opps created":         "s1",
    "opportunities created":"s1",
    "s2":                   "s2",
    "advanced to s2":       "s2",
    "poc":                  "poc",
    "wins #":               "wins",
    "wins#":                "wins",
    "wins number":          "wins",
    "wins €":               "value",
    "wins€":                "value",
    "value":                "value",
    "wins value":           "value",
    "bookings":             "value",
}


def download_workbook():
    session = requests.Session()
    session.headers["User-Agent"] = (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0 Safari/537.36"
    )
    resp = session.get(SHAREPOINT_URL, allow_redirects=True, timeout=60)
    resp.raise_for_status()
    return load_workbook(BytesIO(resp.content), data_only=True, read_only=True)


def read_actuals_sheet(wb):
    """
    Read the 'Actuals' summary sheet.

    Expected layout:
        Row 1: title banner (skip)
        Row 2: month numbers 1–12 in cols C–N (skip, used for position)
        Row 3: headers — col A=Source, col B=Metric, cols C–N=Jan–Dec
        Row 4+: data rows

    Source column may be blank (carry-forward from previous row).
    """
    ws = wb["Actuals"]
    all_rows = list(ws.iter_rows(values_only=True))

    # Find the header row: first row where col A or col B contains "Source" or "Metric"
    header_idx = None
    for i, row in enumerate(all_rows):
        vals = [str(v).strip().lower() if v is not None else "" for v in row[:3]]
        if "source" in vals or "metric" in vals:
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Could not find header row in 'Actuals' sheet")

    data_rows = all_rows[header_idx + 1:]

    # Empty result: all None (= no data yet for that month)
    actuals = {
        src: {m: [None] * 12 for m in ("s1", "s2", "poc", "wins", "value")}
        for src in ("marketing", "bdr", "partner")
    }

    current_src = None

    for row in data_rows:
        # Skip fully empty rows
        if all(v is None for v in row):
            continue

        src_raw    = str(row[0]).strip().lower() if row[0] is not None else ""
        metric_raw = str(row[1]).strip().lower() if row[1] is not None else ""

        # Carry-forward source
        if src_raw in SRC_MAP:
            current_src = SRC_MAP[src_raw]
        # If col A is blank, keep current_src

        if current_src is None:
            continue

        metric = METRIC_MAP.get(metric_raw)
        if metric is None:
            continue

        # Columns C–N (index 2–13) = Jan–Dec
        for mo_idx in range(12):
            col_idx = 2 + mo_idx
            if col_idx >= len(row):
                break
            val = row[col_idx]
            if val is not None:
                try:
                    actuals[current_src][metric][mo_idx] = float(val)
                except (TypeError, ValueError):
                    pass  # leave as None

    return actuals


def main():
    print("[compute.py] Downloading SharePoint file…")
    wb = download_workbook()
    print("[compute.py] Reading Actuals sheet…")
    actuals = read_actuals_sheet(wb)
    wb.close()

    payload = {
        "targets":      TARGETS,
        "actuals":      actuals,
        "last_fetched": datetime.datetime.utcnow().isoformat() + "Z",
        "source":       "sharepoint_actuals_sheet",
    }

    with open(OUTPUT_FILE, "w") as f:
        json.dump(payload, f, indent=2)

    print(f"[compute.py] ✓ Wrote {OUTPUT_FILE}")

    # Print summary for verification
    for src in ("marketing", "bdr", "partner"):
        s1 = actuals[src]["s1"][:3]
        print(f"  {src:12s} S1 Jan-Mar: {s1}")


if __name__ == "__main__":
    main()


