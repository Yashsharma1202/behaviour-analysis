"""
fetch_futures_from_quant_db.py
===============================================================================
Live source for futures data: a read-only PostgreSQL connection to quant_db
(credentials in quant_db_config.json, gitignored -- never commit it).

quant_db's tables (Index_Futures, Stock_Futures_A_F/G_L/M_S/NUM/T_Z) are
confirmed the SAME upstream source the existing spot_parquet/ and
options_parquet/ caches were originally dumped from -- identical column set
(Ticker, Date, Time, Open, High, Low, Close, Volume, OpenInterest, Instrument,
Type, Strike, ExpiryDate, ContractType, DaysToExpiry, MetaData, id), same
1-minute bar granularity, same per-date-file convention. That dump pipeline
stopped after 2026-07-06; quant_db itself kept being fed daily and is current
through "yesterday" (checked 2026-10-07: data through 2026-10-06).

This script backfills futures_parquet/<YYYY-MM-DD>.parquet for every trading
day in quant_db newer than the latest local file, unioning all 6 futures
tables per day -- mirroring spot_parquet/options_parquet's own convention so
it's additive, not a competing format.

Usage:
    python fetch_futures_from_quant_db.py              # backfill missing days
    python fetch_futures_from_quant_db.py --since 2026-09-01   # force a start date
===============================================================================
"""
import argparse
import json
import pathlib
import sys
import time

import pandas as pd
import psycopg2

sys.stdout.reconfigure(encoding='utf-8')

ROOT = pathlib.Path(__file__).resolve().parent
OUT_DIR = ROOT / 'futures_parquet'
OUT_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PATH = ROOT / 'quant_db_config.json'

FUTURES_TABLES = [
    'Index_Futures',
    'Stock_Futures_A_F', 'Stock_Futures_G_L', 'Stock_Futures_M_S',
    'Stock_Futures_NUM', 'Stock_Futures_T_Z',
]

COLUMNS = [
    'Ticker', 'Date', 'Time', 'Open', 'High', 'Low', 'Close', 'Volume',
    'OpenInterest', 'Instrument', 'Type', 'Strike', 'ExpiryDate',
    'ContractType', 'DaysToExpiry', 'MetaData', 'id',
]


def load_config():
    if not CONFIG_PATH.exists():
        raise SystemExit(
            f"{CONFIG_PATH.name} not found. Create it (gitignored) with "
            f'{{"host":..., "port":..., "dbname":..., "user":..., "password":...}}.'
        )
    return json.loads(CONFIG_PATH.read_text(encoding='utf-8'))


def connect():
    cfg = load_config()
    return psycopg2.connect(
        host=cfg['host'], port=cfg['port'], dbname=cfg['dbname'],
        user=cfg['user'], password=cfg['password'], connect_timeout=15,
    )


def latest_local_date():
    files = sorted(OUT_DIR.glob('*.parquet'))
    return files[-1].stem if files else None


def trading_days_after(conn, since_date_str):
    cur = conn.cursor()
    cur.execute(
        'SELECT DISTINCT "Date" FROM "Index_Futures" WHERE "Date" > %s ORDER BY "Date"',
        (since_date_str,),
    )
    return [r[0].isoformat() for r in cur.fetchall()]


_COL_LIST = ', '.join(f'"{c}"' for c in COLUMNS)


def fetch_day(conn, date_str):
    frames = []
    for tbl in FUTURES_TABLES:
        q = f'SELECT {_COL_LIST} FROM "{tbl}" WHERE "Date" = %s'
        df = pd.read_sql(q, conn, params=(date_str,))
        if not df.empty:
            frames.append(df)
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--since', help='force backfill start date (YYYY-MM-DD), exclusive')
    args = ap.parse_args()

    since = args.since or latest_local_date()
    print(f"futures_parquet/ latest local file: {since or '(none -- first run)'}")

    conn = connect()
    print("Connected to quant_db.")

    if since is None:
        print("No --since given and no local files exist -- refusing to backfill the full "
              "2019-present history in one run (far too large). Pass --since explicitly.")
        conn.close()
        return

    missing = trading_days_after(conn, since)
    if not missing:
        print("Already up to date -- nothing to backfill.")
        conn.close()
        return

    print(f"Backfilling {len(missing)} trading day(s): {missing[0]} .. {missing[-1]}")
    t0 = time.time()
    written = 0
    for i, d in enumerate(missing, 1):
        out_path = OUT_DIR / f"{d}.parquet"
        if out_path.exists():
            continue
        df = fetch_day(conn, d)
        if df is None or df.empty:
            print(f"  [{i}/{len(missing)}] {d}: no rows (holiday?) -- skipped")
            continue
        df.to_parquet(out_path, index=False)
        written += 1
        print(f"  [{i}/{len(missing)}] {d}: {len(df):,} rows -> {out_path.name}")

    conn.close()
    elapsed = time.time() - t0
    print(f"\nDone. {written} file(s) written in {elapsed:.1f}s.")


if __name__ == '__main__':
    main()
