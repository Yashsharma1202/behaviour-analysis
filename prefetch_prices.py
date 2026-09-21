"""
prefetch_prices.py
===============================================================================
Warm both price caches for the whole research universe before the heavy stages
need them.

Two caches, two purposes, both from Yahoo:

    processed/price_cache/<SYM>.csv       adjusted close only — what
                                          event_stats.build_panel() and the CAR
                                          computation in disclosure_timing.py
                                          read (via event_behaviour.fetch_prices_yahoo)

    processed/raw_price_cache/<SYM>.csv   OHLCV, dividend-UNadjusted — what
                                          technicals.py reads, and what the
                                          Stage-3 tiered cost model needs to
                                          bucket stocks by traded value

Doing this as its own step means the expensive stages fail fast on a missing
symbol instead of stalling 20 minutes into a GPU run. It is resumable: anything
already cached is skipped, so re-running after a rate-limit is free.

Runs against a different host from download_feeds.py (Yahoo, not NSE), so the
two can safely run at the same time.

    python prefetch_prices.py                # the universe from universe.json
    python prefetch_prices.py --only-adj     # just the CAR panel's needs
    python prefetch_prices.py RELIANCE TCS
===============================================================================
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("symbols", nargs="*")
    ap.add_argument("--only-adj", action="store_true",
                    help="adjusted close only (skip OHLCV)")
    ap.add_argument("--only-ohlcv", action="store_true")
    ap.add_argument("--all", action="store_true",
                    help="every symbol with feeds on disk, not just universe.json")
    args = ap.parse_args()

    if args.symbols:
        syms = [s.upper() for s in args.symbols]
    elif args.all:
        # Every symbol whose feeds are on disk — the CAR panel and the backtest
        # can only use stocks that have both filings and prices.
        syms = sorted(p.parent.name for p in ROOT.glob("*/announcements.csv"))
    else:
        import build_universe
        syms = build_universe.symbols()

    import event_behaviour as EB
    import results_dividend_behaviour as RDB

    adj_dir = ROOT / "processed" / "price_cache"
    raw_dir = ROOT / "processed" / "raw_price_cache"

    print("=" * 74)
    print(f"  PREFETCH PRICES — {len(syms)} symbols")
    print(f"  adjusted close : {len(list(adj_dir.glob('*.csv'))) if adj_dir.exists() else 0} cached")
    print(f"  OHLCV          : {len(list(raw_dir.glob('*.csv'))) if raw_dir.exists() else 0} cached")
    print("=" * 74)

    ok_a = ok_r = miss_a = miss_r = 0
    t0 = time.time()
    for i, s in enumerate(syms, 1):
        if not args.only_ohlcv:
            if (adj_dir / f"{s}.csv").exists():
                ok_a += 1
            else:
                try:
                    px = EB.fetch_prices_yahoo(s)
                    if px is not None and len(px):
                        ok_a += 1
                    else:
                        miss_a += 1
                except Exception:                 # noqa: BLE001
                    miss_a += 1
        if not args.only_adj:
            p = raw_dir / f"{s}.csv"
            if p.exists():
                ok_r += 1
            else:
                try:
                    df = RDB.fetch_raw_prices(s)
                    if df is not None and len(df):
                        ok_r += 1
                    else:
                        miss_r += 1
                except Exception:                 # noqa: BLE001
                    miss_r += 1
        if i % 25 == 0 or i == len(syms):
            print(f"  [{i:>4}/{len(syms)}]  adj ok {ok_a:>4} miss {miss_a:>3}  |  "
                  f"ohlcv ok {ok_r:>4} miss {miss_r:>3}  |  {time.time() - t0:>5.0f}s")

    print("=" * 74)
    print(f"  adjusted close : {ok_a} available, {miss_a} unavailable")
    print(f"  OHLCV          : {ok_r} available, {miss_r} unavailable")
    print(f"  {time.time() - t0:.0f}s total")
    if miss_a or miss_r:
        print("\n  Symbols Yahoo has no data for are simply dropped downstream —")
        print("  event_stats.build_panel and technicals both skip them.")


if __name__ == "__main__":
    main()
