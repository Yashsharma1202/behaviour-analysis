"""
results_tracker.py
===============================================================================
For every QUARTERLY-RESULT event, show three prices and the moves between them:

    * price 6 trading days BEFORE the result day   (configurable: BEFORE)
    * price ON the result day                       (first trading day on/after)
    * price 3 trading days AFTER the result day     (configurable: AFTER)

...plus the % move  before->result,  result->after,  and  before->after.

Uses the downloaded NSE result dates in each <SYMBOL>/financial_results.csv and
daily prices from Yahoo (adjusted close, cached in processed/price_cache/).

RUN
    python results_tracker.py                 # every stock with downloaded feeds
    python results_tracker.py RELIANCE        # one stock
    python results_tracker.py RELIANCE INFY   # several
    python results_tracker.py RELIANCE 10 5   # BEFORE=10, AFTER=5 (numbers last)

Output: printed table + processed/results_tracker.csv
===============================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

import event_behaviour as EB          # reuse feed loader + price fetch

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                     # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
PROC.mkdir(exist_ok=True)

BEFORE = 6        # trading days before the result day
AFTER = 3         # trading days after the result day


def track(symbol: str, px: pd.Series) -> list[dict]:
    """One stock: a row per result with the 3 prices and the moves between them."""
    events = EB.load_events_from_feeds([symbol])
    results = events[events["event_type"] == "RESULTS"]["event_date"].drop_duplicates()
    idx = px.index
    rows = []
    for d in sorted(results):
        pos = idx.searchsorted(pd.Timestamp(d))        # result day = first day on/after
        if pos - BEFORE < 0 or pos + AFTER >= len(idx) or pos >= len(idx):
            continue
        p_before = float(px.iloc[pos - BEFORE])
        p_result = float(px.iloc[pos])
        p_after = float(px.iloc[pos + AFTER])
        rows.append({
            "symbol": symbol,
            "result_day": idx[pos].date(),
            f"price_{BEFORE}d_before": round(p_before, 2),
            "price_on_result": round(p_result, 2),
            f"price_{AFTER}d_after": round(p_after, 2),
            "before_to_result_%": round((p_result / p_before - 1) * 100, 2),
            "result_to_after_%": round((p_after / p_result - 1) * 100, 2),
            "before_to_after_%": round((p_after / p_before - 1) * 100, 2),
            "outcome": "UP" if p_after > p_before else ("DOWN" if p_after < p_before else "FLAT"),
        })
    return rows


def main():
    # parse args: symbols (words) and optional BEFORE/AFTER (numbers)
    global BEFORE, AFTER
    args = sys.argv[1:]
    nums = [a for a in args if a.isdigit()]
    syms = [a.upper() for a in args if not a.isdigit()]
    if len(nums) >= 1:
        BEFORE = int(nums[0])
    if len(nums) >= 2:
        AFTER = int(nums[1])
    if not syms:
        syms = EB.discover_feed_symbols()
    if not syms:
        raise SystemExit("No downloaded event feeds found. Download some in the "
                         "stock browser first (python stock_server.py).")

    print(f"Tracking RESULTS: price {BEFORE}d before -> result day -> {AFTER}d after")
    print(f"Stocks: {', '.join(syms)}\n")

    all_rows = []
    for s in syms:
        px = EB.fetch_prices_yahoo(s)
        if px is None or len(px) < BEFORE + AFTER + 5:
            print(f"  {s}: no price data, skipped")
            continue
        rows = track(s, px)
        all_rows.extend(rows)
        print(f"  {s}: {len(rows)} results tracked")

    if not all_rows:
        raise SystemExit("No result events fell inside the available price history.")

    df = pd.DataFrame(all_rows).sort_values(["symbol", "result_day"])
    out = PROC / "results_tracker.csv"
    df.to_csv(out, index=False)

    # ── printed table ────────────────────────────────────────────────────────
    bcol, acol = f"price_{BEFORE}d_before", f"price_{AFTER}d_after"
    print("\n" + "=" * 92)
    print(f"  {'symbol':>10} {'result_day':>12} {f'{BEFORE}d before':>11} "
          f"{'on result':>10} {f'{AFTER}d after':>10} "
          f"{'bef→res':>8} {'res→aft':>8} {'bef→aft':>8}  out")
    print("=" * 92)
    for _, r in df.iterrows():
        print(f"  {r['symbol']:>10} {r['result_day']!s:>12} "
              f"{r[bcol]:>11,.2f} {r['price_on_result']:>10,.2f} {r[acol]:>10,.2f} "
              f"{r['before_to_result_%']:>+7.2f}% {r['result_to_after_%']:>+7.2f}% "
              f"{r['before_to_after_%']:>+7.2f}%  {r['outcome']}")
    print("=" * 92)

    # ── per-stock summary ─────────────────────────────────────────────────────
    print("\n  SUMMARY (across tracked results):")
    for s, g in df.groupby("symbol"):
        up = (g["outcome"] == "UP").mean() * 100
        print(f"   {s:>10}: n={len(g):>3}  avg before→after {g['before_to_after_%'].mean():+.2f}%  "
              f"avg result→after {g['result_to_after_%'].mean():+.2f}%  UP {up:.0f}% of the time")
    print(f"\n  Saved -> {out}")


if __name__ == "__main__":
    main()
