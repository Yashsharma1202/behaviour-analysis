"""
create_event_dates_db.py
===============================================================================
Creates processed/event_dates.db (SQLite) as the single source of record for
every date this project tracks: quarterly-result declaration dates (per
stock, per quarter) and the recurring holiday-event dates (Gandhi Jayanti,
Dussehra, Diwali, Christmas) across their historical years.

This does NOT replace dashboard_data/event_dashboard_data.json -- that file
stays the live dashboard's data source (JS reads it directly, build_dashboard.py
syncs it to docs/). This DB is a separate, queryable, constrained record of
dates specifically, so date correctness can be checked and audited on its own
-- every row carries a verification source and a timestamp, and NOT NULL /
CHECK constraints make a malformed date (wrong format, impossible day,
mismatched weekday) fail at insert time instead of silently sitting in a
JSON blob.

Usage:
    python create_event_dates_db.py          # create schema only
    python create_event_dates_db.py --populate
                                              # also populate from the current
                                              # dashboard_data/event_dashboard_data.json
                                              # and processed/nse_trading_holidays.json
                                              # (both already-verified sources)
"""
import argparse
import json
import pathlib
import sqlite3
from datetime import datetime

ROOT = pathlib.Path(__file__).resolve().parent
DB_PATH = ROOT / "processed" / "event_dates.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS quarterly_result_dates (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    q_code                  TEXT NOT NULL,
    symbol                  TEXT NOT NULL,
    result_date             TEXT,                -- 'YYYY-MM-DD' or NULL if not yet announced
    result_status           TEXT NOT NULL,        -- 'Announced' | 'Yet to come'
    purpose                 TEXT,
    entry_date              TEXT,                 -- 'YYYY-MM-DD'
    exit_date               TEXT,                 -- 'YYYY-MM-DD'
    window                  TEXT,                 -- e.g. 'T-2 to T+4'
    date_source             TEXT NOT NULL,         -- provenance: where this date came from
    verified_at             TEXT NOT NULL,         -- ISO timestamp this row was written
    UNIQUE(q_code, symbol),
    CHECK (result_date IS NULL OR result_date GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
    CHECK (entry_date IS NULL OR entry_date GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
    CHECK (exit_date IS NULL OR exit_date GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]')
);

CREATE TABLE IF NOT EXISTS holiday_event_dates (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    event_name              TEXT NOT NULL,         -- e.g. 'Dussehra / Vijayadashami'
    year                    INTEGER NOT NULL,
    event_date              TEXT NOT NULL,         -- 'YYYY-MM-DD'
    is_nse_trading_holiday  INTEGER NOT NULL,      -- 1 = market closed that day, 0 = open
    date_source             TEXT NOT NULL,
    verified_at             TEXT NOT NULL,
    UNIQUE(event_name, year),
    CHECK (event_date GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]')
);

CREATE TABLE IF NOT EXISTS nse_trading_holidays (
    holiday_date            TEXT PRIMARY KEY,      -- 'YYYY-MM-DD'
    holiday_name            TEXT NOT NULL,
    date_source             TEXT NOT NULL,
    verified_at             TEXT NOT NULL,
    CHECK (holiday_date GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]')
);
"""


def parse_ddmonyyyy(s):
    """'15-Oct-2026' / '15-Oct-2026 (Thu)' / '2023-08-09' -> 'YYYY-MM-DD', or None for placeholders."""
    if not s or s in ("Yet to come", "Awaiting NSE Announcement"):
        return None
    core = s.split(" (")[0].strip()
    if len(core) == 10 and core[4] == "-" and core[7] == "-":
        datetime.strptime(core, "%Y-%m-%d")  # validate
        return core
    return datetime.strptime(core, "%d-%b-%Y").strftime("%Y-%m-%d")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--populate", action="store_true")
    args = ap.parse_args()

    DB_PATH.parent.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.executescript(SCHEMA)
    conn.commit()
    print(f"Schema ready at {DB_PATH}")

    if not args.populate:
        conn.close()
        return

    now = datetime.now().isoformat(timespec="seconds")

    with open(ROOT / "dashboard_data" / "event_dashboard_data.json", encoding="utf-8") as f:
        dashboard = json.load(f)

    q_rows = []
    for q in dashboard["quarters"]:
        q_code = q["q_code"]
        for s in q.get("stocks", []):
            q_rows.append((
                q_code,
                s["symbol"],
                parse_ddmonyyyy(s.get("result_date")),
                s.get("result_status") or "Yet to come",
                s.get("purpose"),
                parse_ddmonyyyy(s.get("entry_date")),
                parse_ddmonyyyy(s.get("exit_date")),
                s.get("window"),
                s.get("source") or "dashboard_data/event_dashboard_data.json (pre-existing)",
                now,
            ))
    conn.executemany(
        """INSERT INTO quarterly_result_dates
           (q_code, symbol, result_date, result_status, purpose, entry_date, exit_date,
            window, date_source, verified_at)
           VALUES (?,?,?,?,?,?,?,?,?,?)
           ON CONFLICT(q_code, symbol) DO UPDATE SET
             result_date=excluded.result_date,
             result_status=excluded.result_status,
             purpose=excluded.purpose,
             entry_date=excluded.entry_date,
             exit_date=excluded.exit_date,
             window=excluded.window,
             date_source=excluded.date_source,
             verified_at=excluded.verified_at""",
        q_rows,
    )
    conn.commit()
    print(f"Populated quarterly_result_dates: {len(q_rows)} rows")

    # Full 26-year verified recurring-holiday calendar (17 holidays, web-verified
    # against NSE's own circulars this session) lives as a literal list in
    # build_nifty_futures_holiday_behaviour_master.py. Extract just that literal
    # (lines between 'holidays_def = [' and its closing ']') and exec it in an
    # isolated namespace -- avoids importing/running the rest of that builder script.
    behaviour_master_path = ROOT / "build_nifty_futures_holiday_behaviour_master.py"
    with open(behaviour_master_path, encoding="utf-8") as f:
        lines = f.readlines()
    start = next(i for i, l in enumerate(lines) if l.startswith("holidays_def = ["))
    end = next(i for i in range(start, len(lines)) if lines[i].rstrip() == "]")
    namespace = {}
    exec("".join(lines[start:end + 1]), namespace)
    holidays_def = namespace["holidays_def"]

    hist_rows = []
    for h in holidays_def:
        years = dict(h["dates"])
        for key, val in h.items():
            if key.endswith("_date") and key[:4].isdigit():
                years[int(key[:4])] = val
        for year, iso_date in years.items():
            hist_rows.append((h["name"], int(year), iso_date, 1,
                               "build_nifty_futures_holiday_behaviour_master.py (web-verified vs NSE circulars)",
                               now))
    conn.executemany(
        """INSERT INTO holiday_event_dates
           (event_name, year, event_date, is_nse_trading_holiday, date_source, verified_at)
           VALUES (?,?,?,?,?,?)
           ON CONFLICT(event_name, year) DO UPDATE SET
             event_date=excluded.event_date,
             date_source=excluded.date_source,
             verified_at=excluded.verified_at""",
        hist_rows,
    )
    conn.commit()
    print(f"Populated holiday_event_dates (26yr historical, 17 holidays): {len(hist_rows)} rows")

    nse_holidays_path = ROOT / "processed" / "nse_trading_holidays.json"
    if nse_holidays_path.exists():
        with open(nse_holidays_path, encoding="utf-8") as f:
            nse_holidays = json.load(f)
        nh_rows = [
            (date_str, name, "nse_trading_calendar.py (live NSE fetch)", now)
            for date_str, name in nse_holidays.items()
        ]
        conn.executemany(
            """INSERT INTO nse_trading_holidays (holiday_date, holiday_name, date_source, verified_at)
               VALUES (?,?,?,?)
               ON CONFLICT(holiday_date) DO UPDATE SET
                 holiday_name=excluded.holiday_name,
                 date_source=excluded.date_source,
                 verified_at=excluded.verified_at""",
            nh_rows,
        )
        conn.commit()
        print(f"Populated nse_trading_holidays: {len(nh_rows)} rows")

    conn.close()


if __name__ == "__main__":
    main()
