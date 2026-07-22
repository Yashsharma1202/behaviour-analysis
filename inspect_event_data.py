"""
inspect_event_data.py
===============================================================================
STEP 1 for the event behaviour-analysis. Inspects the TWO input CSVs BEFORE any
analysis is written, so we confirm the exact schema together first.

  1. PRICE FILE  : daily OHLC   -> date, symbol, open, high, low, close, volume
  2. EVENT FILE  : event calendar -> symbol, event_date, event_type

It prints, for each file: size, rows, columns + dtypes + null counts, first /
last / random sample rows, detected key columns, date range, symbol coverage,
and (for events) the event_type breakdown. Finally it cross-checks that event
symbols exist in the price data and that their date ranges overlap.

RUN
    python inspect_event_data.py  "PATH\\to\\prices.csv"  "PATH\\to\\events.csv"
  or set PRICE_CSV / EVENT_CSV below and run:  python inspect_event_data.py
===============================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

# ── set paths here, or pass them on the command line ────────────────────────
PRICE_CSV = r""
EVENT_CSV = r""

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)

SEP = "=" * 82


def banner(t):
    print(f"\n{SEP}\n  {t}\n{SEP}")


def sec(t):
    print(f"\n{'-' * 70}\n  {t}\n{'-' * 70}")


def find_col(df, candidates):
    """Find the first matching column (case-insensitive)."""
    lower = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand in df.columns:
            return cand
        if cand.lower() in lower:
            return lower[cand.lower()]
    return None


def profile(path, kind):
    p = Path(path)
    if not p.exists():
        sys.exit(f"[ERROR] {kind} file not found: {path}")

    banner(f"{kind} FILE  —  {p.name}")
    print(f"  path : {p.resolve()}")
    print(f"  size : {p.stat().st_size / 1_048_576:.2f} MB")
    df = pd.read_csv(path, low_memory=False)
    print(f"  rows : {len(df):,}")
    print(f"  cols : {df.shape[1]}")

    sec("COLUMNS · DTYPES · NULLS")
    info = pd.DataFrame({
        "column": df.columns,
        "dtype": [str(t) for t in df.dtypes],
        "non_null": df.notna().sum().values,
        "null_%": (df.isna().mean() * 100).round(2).values,
        "sample_value": [df[c].dropna().iloc[0] if df[c].notna().any() else "" for c in df.columns],
    })
    print(info.to_string(index=False))

    sec("FIRST 5 ROWS")
    print(df.head().to_string(index=False))
    sec("LAST 5 ROWS")
    print(df.tail().to_string(index=False))
    if len(df) > 5:
        sec("5 RANDOM ROWS")
        print(df.sample(5, random_state=1).to_string(index=False))
    return df


def main():
    price_path = sys.argv[1] if len(sys.argv) > 1 else PRICE_CSV
    event_path = sys.argv[2] if len(sys.argv) > 2 else EVENT_CSV
    if not price_path or not event_path:
        sys.exit("Usage: python inspect_event_data.py <prices.csv> <events.csv>\n"
                 "   (or set PRICE_CSV / EVENT_CSV at the top of the file)")

    # ── PRICE ───────────────────────────────────────────────────────────────
    px = profile(price_path, "PRICE / OHLC")
    p_date = find_col(px, ["date", "Date", "DATE", "trade_date", "timestamp"])
    p_sym = find_col(px, ["symbol", "Symbol", "SYMBOL", "ticker", "scrip"])
    p_close = find_col(px, ["close", "Close", "CLOSE", "adj_close", "Adj Close"])
    p_open = find_col(px, ["open", "Open", "OPEN"])
    if p_date:
        px[p_date] = pd.to_datetime(px[p_date], errors="coerce", dayfirst=True)
        sec("PRICE · DATE RANGE & SYMBOLS")
        print(f"  date column   : '{p_date}'   range {px[p_date].min()} -> {px[p_date].max()}")
        print(f"  trading days  : {px[p_date].nunique():,}")
    if p_sym:
        vc = px[p_sym].value_counts()
        print(f"  symbol column : '{p_sym}'   unique symbols: {px[p_sym].nunique():,}")
        print(f"  rows per symbol (top 10):")
        print(vc.head(10).to_string())

    # ── EVENTS ──────────────────────────────────────────────────────────────
    ev = profile(event_path, "EVENT CALENDAR")
    e_date = find_col(ev, ["event_date", "date", "Date", "eventDate", "ex_date", "exDate"])
    e_sym = find_col(ev, ["symbol", "Symbol", "SYMBOL", "ticker", "scrip"])
    e_type = find_col(ev, ["event_type", "type", "eventType", "category", "event"])
    if e_date:
        ev[e_date] = pd.to_datetime(ev[e_date], errors="coerce", dayfirst=True)
        sec("EVENTS · DATE RANGE & TYPES")
        print(f"  event-date column : '{e_date}'   range {ev[e_date].min()} -> {ev[e_date].max()}")
    if e_type:
        print(f"  event-type column : '{e_type}'")
        print("  event_type breakdown:")
        print(ev[e_type].value_counts(dropna=False).to_string())
    if e_sym:
        print(f"  symbol column     : '{e_sym}'   unique symbols with events: {ev[e_sym].nunique():,}")

    # ── CROSS-CHECK ─────────────────────────────────────────────────────────
    banner("CROSS-CHECK  (do the two files line up?)")
    if p_sym and e_sym:
        ps, es = set(px[p_sym].unique()), set(ev[e_sym].unique())
        common = ps & es
        print(f"  symbols in prices        : {len(ps):,}")
        print(f"  symbols in events        : {len(es):,}")
        print(f"  symbols in BOTH          : {len(common):,}")
        missing = sorted(es - ps)[:10]
        if missing:
            print(f"  event symbols NOT in prices (first 10): {missing}")
    if p_date and e_date:
        overlap_lo = max(px[p_date].min(), ev[e_date].min())
        overlap_hi = min(px[p_date].max(), ev[e_date].max())
        print(f"  overlapping date window  : {overlap_lo.date()} -> {overlap_hi.date()}")

    # ── DETECTED MAPPING ────────────────────────────────────────────────────
    banner("DETECTED COLUMN MAPPING  (confirm these before I build the analysis)")
    mapping = {
        "price.date": p_date, "price.symbol": p_sym, "price.open": p_open,
        "price.close": p_close,
        "event.symbol": e_sym, "event.event_date": e_date, "event.event_type": e_type,
    }
    for k, v in mapping.items():
        flag = "OK   " if v else "MISS "
        print(f"  [{flag}] {k:20s} -> {v or '[NOT FOUND]'}")

    banner("NEXT")
    print("  Share this output with me. Once we confirm the columns, the date format,")
    print("  and how the two files join on symbol, I'll write the behaviour-analysis")
    print("  script (entry T-1..T-N  ->  exit T+1..T+M, per event_type + best entry).")


if __name__ == "__main__":
    main()
