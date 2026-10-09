"""
verify_db_against_nse_financial_results.py
===============================================================================
Cross-checks processed/event_dates.db's quarterly_result_dates against each
stock's own NSE-sourced <SYMBOL>/financial_results.csv (broadCastDate +
fromDate/toDate period, NSE's own primary record of when each result was
actually declared). This file only holds NSE's most recent ~5 quarters per
stock (that's what NSE's API returns), so it can only authoritatively confirm
recent quarters -- but for those, it's ground truth, not a guess.

Maps each financial_results.csv row's (fromDate, toDate) period to this
project's q_code (FY<YY>_Q<N>, where Q1=Apr-Jun, Q2=Jul-Sep, Q3=Oct-Dec,
Q4=Jan-Mar) and compares the real broadCastDate against what's stored.
"""
import os
import pathlib
import sqlite3
import sys
from datetime import datetime

sys.stdout.reconfigure(errors="replace")

ROOT = pathlib.Path("D:/behaviour analysis")
DB_PATH = ROOT / "processed" / "event_dates.db"


def period_to_qcode(from_date, to_date):
    # from_date/to_date like '01-OCT-2024'; Indian FY: Apr-Jun=Q1 .. Jan-Mar=Q4
    to_dt = datetime.strptime(to_date, "%d-%b-%Y")
    month, year = to_dt.month, to_dt.year
    if month in (4, 5, 6):
        q, fy_end = 1, year + 1
    elif month in (7, 8, 9):
        q, fy_end = 2, year + 1
    elif month in (10, 11, 12):
        q, fy_end = 3, year + 1
    else:  # Jan-Mar
        q, fy_end = 4, year
    return f"FY{str(fy_end)[2:]}_Q{q}"


def main():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    checked = 0
    matches = 0
    mismatches = []
    no_db_row = []

    for entry in sorted(os.scandir(ROOT), key=lambda e: e.name):
        if not entry.is_dir():
            continue
        sym = entry.name.strip().upper()
        fr_path = ROOT / entry.name / "financial_results.csv"
        if not fr_path.exists() or fr_path.stat().st_size < 10:
            continue
        try:
            import pandas as pd
            df = pd.read_csv(fr_path, dtype=str).fillna("")
        except Exception:
            continue
        if "broadCastDate" not in df.columns:
            continue

        for _, r in df.iterrows():
            bdate = r.get("broadCastDate", "")
            fdate = r.get("fromDate", "")
            tdate = r.get("toDate", "")
            if not (bdate and fdate and tdate):
                continue
            try:
                nse_date = datetime.strptime(bdate, "%d-%b-%Y").strftime("%Y-%m-%d")
                q_code = period_to_qcode(fdate, tdate)
            except ValueError:
                continue

            row = c.execute(
                "SELECT result_date FROM quarterly_result_dates WHERE q_code=? AND symbol=?",
                (q_code, sym),
            ).fetchone()
            checked += 1
            if row is None:
                no_db_row.append((sym, q_code, nse_date))
                continue
            db_date = row[0]
            if db_date == nse_date:
                matches += 1
            else:
                mismatches.append((sym, q_code, db_date, nse_date))

    print(f"Checked {checked} NSE-sourced (symbol, quarter) broadcast dates against the DB")
    print(f"Matches: {matches}")
    print(f"Mismatches: {len(mismatches)}")
    for m in mismatches:
        print("  MISMATCH:", m)
    print(f"No corresponding DB row (symbol/quarter not in our 18-quarter universe): {len(no_db_row)}")
    for n in no_db_row[:20]:
        print("  NO ROW:", n)
    if len(no_db_row) > 20:
        print(f"  ... and {len(no_db_row) - 20} more")

    conn.close()


if __name__ == "__main__":
    main()
