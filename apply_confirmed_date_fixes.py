"""
apply_confirmed_date_fixes.py
===============================================================================
Applies the 66 web-source-confirmed quarterly result-date corrections found
by 4 parallel verification agents (web-searched against NSE board-meeting
filings, company press releases, and dated news articles) across FY23_Q2
through FY26_Q1.

DATE-ONLY fix, per explicit instruction: corrects result_date, entry_date,
exit_date (entry/exit recomputed via the live NSE trading calendar, preserving
each row's existing lead/hold window -- NOT resetting the window). Does NOT
touch entry_px/exit_px/actual_pnl/actual_return/outcome -- those were computed
against the old (wrong) dates and need their own recompute pass later. Every
row this touches gets a new `date_fix_note` field recording the correction so
that follow-up pass has a clear list of what to redo.

Applies to both processed/event_dates.db and dashboard_data/event_dashboard_data.json
(the live site's actual source).
"""
import json
import sqlite3
from datetime import datetime
import nse_trading_calendar as ntc

ROOT = r"D:\behaviour analysis"
DB_PATH = ROOT + r"\processed\event_dates.db"
DASHBOARD_PATH = ROOT + r"\dashboard_data\event_dashboard_data.json"

# (q_code, symbol, new_result_date_iso, source_note)
FIXES = [
    ("FY23_Q2", "TCS", "2022-10-10", "TCS board meeting + press release"),
    ("FY23_Q2", "SBIN", "2022-11-05", "SBI board meeting"),
    ("FY23_Q2", "AXISBANK", "2022-10-20", "Axis Bank board meeting"),
    ("FY23_Q2", "LT", "2022-10-31", "L&T results + earnings call"),
    ("FY23_Q2", "MARUTI", "2022-10-28", "Maruti board approval"),
    ("FY23_Q2", "HDFCBANK", "2022-10-15", "HDFC Bank board meeting (Saturday)"),
    ("FY23_Q2", "WIPRO", "2022-10-12", "Wipro official Q2 FY23 result"),
    ("FY23_Q2", "NESTLEIND", "2022-10-19", "Nestle India board meeting"),

    ("FY23_Q3", "TCS", "2023-01-09", "TCS board meeting"),
    ("FY23_Q3", "SBIN", "2023-02-03", "SBI results day"),
    ("FY23_Q3", "WIPRO", "2023-01-13", "Wipro official Q3 FY23 result"),

    ("FY23_Q4", "TCS", "2023-04-12", "TCS board meeting"),
    ("FY23_Q4", "AXISBANK", "2023-04-27", "Axis Bank board meeting"),
    ("FY23_Q4", "LT", "2023-05-10", "L&T board meeting"),
    ("FY23_Q4", "MARUTI", "2023-04-26", "Maruti board meeting"),
    ("FY23_Q4", "HINDUNILVR", "2023-04-27", "HUL results day"),
    ("FY23_Q4", "POWERGRID", "2023-05-19", "PowerGrid board meeting"),
    ("FY23_Q4", "INFY", "2023-04-13", "Infosys results day"),
    ("FY23_Q4", "HDFCBANK", "2023-04-15", "HDFC Bank results day"),

    ("FY24_Q1", "WIPRO", "2023-07-13", "businesstoday.in Q1FY24 IT results calendar"),
    ("FY24_Q1", "NESTLEIND", "2023-07-27", "Nestle India official board outcome"),
    ("FY24_Q1", "BHARTIARTL", "2023-08-03", "Business Standard"),
    ("FY24_Q1", "TITAN", "2023-08-02", "Business Standard / indiaretailing"),

    ("FY24_Q2", "WIPRO", "2023-10-18", "businesstoday.in Wipro Q2FY24"),
    ("FY24_Q2", "NESTLEIND", "2023-10-19", "5paisa Nestle India Q2FY24"),
    ("FY24_Q2", "HDFCBANK", "2023-10-16", "Business Standard"),
    ("FY24_Q2", "MARUTI", "2023-10-27", "Maruti official press release"),
    ("FY24_Q2", "AXISBANK", "2023-10-25", "Business Standard"),
    ("FY24_Q2", "CIPLA", "2023-10-27", "NSE filing + 5paisa"),

    ("FY24_Q3", "TATAMOTORS", "2024-02-02", "5paisa Tata Motors Q3FY24"),
    ("FY24_Q3", "AXISBANK", "2024-01-23", "NSE BM intimation filing"),
    ("FY24_Q3", "LT", "2024-01-30", "Business Standard"),
    ("FY24_Q3", "NESTLEIND", "2024-02-07", "Nestle India official board outcome"),

    ("FY24_Q4", "INFY", "2024-04-18", "Infosys official press release"),
    ("FY24_Q4", "HDFCBANK", "2024-04-20", "multiple sources (Saturday filing)"),
    ("FY24_Q4", "AXISBANK", "2024-04-24", "Axis Bank press release"),
    ("FY24_Q4", "MARUTI", "2024-04-26", "Business Standard"),
    ("FY24_Q4", "NTPC", "2024-05-24", "NTPC board meeting (Friday)"),
    ("FY24_Q4", "POWERGRID", "2024-05-22", "Business Standard"),
    ("FY24_Q4", "CIPLA", "2024-05-10", "NSE disclosure filing (Friday)"),
    ("FY24_Q4", "GRASIM", "2024-05-22", "Business Standard"),

    ("FY25_Q1", "HDFCBANK", "2024-07-20", "icicidirect (Saturday filing)"),

    ("FY25_Q3", "TATAMOTORS", "2025-01-29", "NSE filing + Jan 2025 news"),
    ("FY25_Q3", "HDFCBANK", "2025-01-22", "NSE BM intimation"),
    ("FY25_Q3", "LT", "2025-01-30", "NSE BM outcome + Business Standard"),
    ("FY25_Q3", "BPCL", "2025-01-22", "Business Standard"),

    ("FY25_Q4", "NESTLEIND", "2025-04-24", "Business Standard + NSE dividend intimation"),
    ("FY25_Q4", "POWERGRID", "2025-05-19", "PowerGrid press release"),
    ("FY25_Q4", "WIPRO", "2025-04-16", "Wipro official press release"),
    ("FY25_Q4", "SBILIFE", "2025-04-24", "NSE integrated filing"),
    ("FY25_Q4", "ETERNAL", "2025-05-01", "Business Standard (was wrong quarter's date)"),
    ("FY25_Q4", "INDIGO", "2025-05-21", "Business Standard + NSE BM intimation (was wrong quarter's date)"),

    ("FY26_Q1", "ICICIBANK", "2025-07-19", "NSE BM outcome + Business Standard"),
    ("FY26_Q1", "BAJFINANCE", "2025-07-24", "NSE prior intimation"),
    ("FY26_Q1", "ASIANPAINT", "2025-07-29", "company BM intimation"),
    ("FY26_Q1", "BHARTIARTL", "2025-08-05", "NSE press release PDF"),
    ("FY26_Q1", "ITC", "2025-08-01", "ITC LODR filing"),
    ("FY26_Q1", "SUNPHARMA", "2025-07-31", "Sun Pharma BM outcome PDF"),
    ("FY26_Q1", "GRASIM", "2025-08-08", "earnings call confirmation"),
    ("FY26_Q1", "HINDALCO", "2025-08-12", "NSE concall invitation + BM intimation"),
    ("FY26_Q1", "ADANIENT", "2025-07-31", "NSE BM intimation PDF"),
    ("FY26_Q1", "DRREDDY", "2025-07-23", "Dr Reddy's press release"),
    ("FY26_Q1", "HINDUNILVR", "2025-07-31", "HUL BM filing"),
    ("FY26_Q1", "BAJAJ-AUTO", "2025-08-06", "NSE outcome filing"),
    ("FY26_Q1", "JIOFIN", "2025-07-17", "NSE BM intimation"),
    ("FY26_Q1", "TRENT", "2025-08-06", "NSE BM outcome (was duplicated from FY25_Q4)"),
]


def fmt_ddmonyyyy(iso_date):
    return datetime.strptime(iso_date, "%Y-%m-%d").strftime("%d-%b-%Y")


def detect_format(sample):
    if not sample or sample in ("Yet to come", "Awaiting NSE Announcement"):
        return "ddmon"
    if len(sample) == 10 and sample[4] == "-" and sample[7] == "-":
        return "iso"
    return "ddmon"


def main():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    db_updated = 0
    for q_code, symbol, new_date, source in FIXES:
        row = c.execute(
            "SELECT result_date, entry_date, exit_date, window FROM quarterly_result_dates WHERE q_code=? AND symbol=?",
            (q_code, symbol),
        ).fetchone()
        if row is None:
            print(f"DB: no row for {q_code}/{symbol}, skipping")
            continue
        old_result_date, old_entry, old_exit, window = row
        lead, hold = 2, 4
        if window and " to " in window:
            try:
                lead = int(window.split("T-")[1].split(" ")[0])
                hold = int(window.split("T+")[1])
            except Exception:
                pass
        holidays = ntc.load_trading_holidays()
        result_dt = datetime.strptime(new_date, "%Y-%m-%d").date()
        new_entry = ntc.add_trading_days(result_dt, -lead, holidays).strftime("%Y-%m-%d")
        new_exit = ntc.add_trading_days(result_dt, hold, holidays).strftime("%Y-%m-%d")
        c.execute(
            """UPDATE quarterly_result_dates
               SET result_date=?, entry_date=?, exit_date=?,
                   date_source=?, verified_at=?
               WHERE q_code=? AND symbol=?""",
            (new_date, new_entry, new_exit,
             f"web-verified correction ({source}); was {old_result_date}",
             datetime.now().isoformat(timespec="seconds"),
             q_code, symbol),
        )
        db_updated += 1
    conn.commit()
    conn.close()
    print(f"DB: updated {db_updated} rows")

    with open(DASHBOARD_PATH, encoding="utf-8") as f:
        dashboard = json.load(f)

    holidays = ntc.load_trading_holidays()
    json_updated = 0
    not_found = []
    for q_code, symbol, new_date, source in FIXES:
        q = next((q for q in dashboard["quarters"] if q["q_code"] == q_code), None)
        if q is None:
            not_found.append((q_code, symbol, "quarter not found"))
            continue
        s = next((s for s in q["stocks"] if s["symbol"] == symbol), None)
        if s is None:
            not_found.append((q_code, symbol, "symbol not found in quarter"))
            continue

        old_result_date = s.get("result_date")
        result_fmt = detect_format(old_result_date)

        lead = s.get("lead", 2)
        hold = s.get("hold", 4)
        result_dt = datetime.strptime(new_date, "%Y-%m-%d").date()
        new_entry_dt = ntc.add_trading_days(result_dt, -lead, holidays)
        new_exit_dt = ntc.add_trading_days(result_dt, hold, holidays)

        if result_fmt == "iso":
            new_result_str = new_date
            new_entry_str = new_entry_dt.strftime("%Y-%m-%d")
            new_exit_str = new_exit_dt.strftime("%Y-%m-%d")
        else:
            new_result_str = fmt_ddmonyyyy(new_date)
            new_entry_str = new_entry_dt.strftime("%d-%b-%Y (%a)")
            new_exit_str = new_exit_dt.strftime("%d-%b-%Y (%a)")

        old_entry = s.get("entry_date")
        old_exit = s.get("exit_date")

        s["result_date"] = new_result_str
        if "result_declaration_date" in s:
            s["result_declaration_date"] = new_result_str
        s["entry_date"] = new_entry_str
        s["exit_date"] = new_exit_str
        s["date_fix_note"] = (
            f"Date corrected via web-source verification on 09-Oct-2026 "
            f"(was result_date={old_result_date}, entry_date={old_entry}, exit_date={old_exit}; "
            f"source: {source}). Price/outcome fields NOT yet recomputed for the corrected "
            f"date -- entry_px/exit_px/actual_pnl/actual_return/outcome below still reflect "
            f"the OLD (wrong) date and need a dedicated recompute pass."
        )
        json_updated += 1

    with open(DASHBOARD_PATH, "w", encoding="utf-8") as f:
        json.dump(dashboard, f, indent=2, ensure_ascii=False)
        f.write("\n")

    print(f"dashboard_data/event_dashboard_data.json: updated {json_updated} rows")
    if not_found:
        print(f"NOT FOUND ({len(not_found)}):")
        for nf in not_found:
            print(" ", nf)


if __name__ == "__main__":
    main()
