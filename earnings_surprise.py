"""
earnings_surprise.py   (v2 — powered by the Screener fundamentals)
===============================================================================
Does Reliance's post-earnings BEHAVIOUR depend on whether profit GREW or FELL
year-on-year? Now labelled from the clean quarterly fundamentals (fund_loader),
which give YoY profit growth for ~13 quarters -- far more than the old XBRL path.

PIPELINE
--------
  1. fund_loader -> quarterly Net Profit + YoY Profit Growth % (already computed)
  2. map each quarter-end to its results ANNOUNCEMENT date (from the event feed)
  3. label BEAT (YoY profit up) / MISS (YoY profit down)
  4. measure the 5-day market-adjusted return after each announcement
  5. compare BEAT vs MISS drift (the PEAD test)

    python earnings_surprise.py

Needs internet (Yahoo). Uses fund_loader + processed/unified_events.csv.
===============================================================================
"""

from __future__ import annotations

import json
import ssl
import sys
import time
import urllib.request
from pathlib import Path

import pandas as pd

import fund_loader

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                              # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
STOCK, MARKET = "RELIANCE.NS", "^NSEI"
WINDOW = 5

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = ssl.CERT_NONE


def fetch_adj(ticker: str) -> pd.Series:
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
           f"?period1=1483228800&period2={int(time.time())}&interval=1d")
    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=30, context=_SSL) as r:
        d = json.load(r)
    res = d["chart"]["result"][0]
    ind = res["indicators"]
    adj = ind.get("adjclose", [{}])[0].get("adjclose") or ind["quote"][0]["close"]
    return pd.Series(adj, index=pd.to_datetime(res["timestamp"], unit="s").normalize()).dropna().sort_index()


def announcement_pool() -> pd.Series:
    """All known results announcement / board-meeting dates, to snap quarters onto."""
    ev = pd.read_csv(PROC / "unified_events.csv", parse_dates=["date"])
    mask = (ev["source"] == "financial_result") | \
           ((ev["source"] == "board_meeting") & ev["event_type"].str.contains("RESULT", na=False))
    return ev[mask]["date"].dt.normalize().drop_duplicates().sort_values()


def build_events(symbol="RELIANCE") -> pd.DataFrame:
    """Quarter -> (announcement_date, net_profit, yoy_growth, label)."""
    data = fund_loader.load_stock(symbol)
    q = data.get("quarterly")
    if q is None:
        raise SystemExit("No quarterly fundamentals for " + symbol)
    yoy_col = next(c for c in q.columns if "YOY Profit Growth" in c)
    np_col = next(c for c in q.columns if c.strip() == "Net Profit")

    pool = announcement_pool()
    rows = []
    for qend, row in q.iterrows():
        g = row[yoy_col]
        if pd.isna(g):
            continue
        month_end = qend + pd.offsets.MonthEnd(0)
        lo, hi = month_end + pd.Timedelta(days=10), month_end + pd.Timedelta(days=80)
        cand = pool[(pool >= lo) & (pool <= hi)]
        if cand.empty:
            continue
        rows.append({"quarter": qend.strftime("%b %Y"),
                     "event_date": cand.iloc[0],
                     "net_profit": row[np_col],
                     "yoy_%": g,
                     "label": "BEAT" if g > 0 else "MISS"})
    return pd.DataFrame(rows)


def post_return(ar, date):
    pos = ar.index.searchsorted(pd.Timestamp(date).normalize())
    if pos <= 0 or pos + WINDOW >= len(ar):
        return float("nan")
    return ar.iloc[pos:pos + WINDOW].sum() * 100


def main():
    events = build_events("RELIANCE")
    if events.empty:
        raise SystemExit("Could not map any quarters to announcement dates.")

    print("Downloading prices ...")
    stock, market = fetch_adj(STOCK), fetch_adj(MARKET)
    ar = (stock.pct_change() - market.pct_change()).dropna()

    events["post5_ar_%"] = [post_return(ar, d) for d in events["event_date"]]
    events = events.dropna(subset=["post5_ar_%"])
    events.to_csv(PROC / "earnings_surprise.csv", index=False)

    print(f"\nEARNINGS labelled by YoY net-profit growth  ({len(events)} quarters)")
    print("-" * 64)
    print(f"  {'quarter':>9} {'results':>12} {'profit':>8} {'YoY':>7} {'label':>5} {'5d after':>9}")
    print("-" * 64)
    for _, r in events.sort_values("event_date").iterrows():
        print(f"  {r['quarter']:>9} {r['event_date'].date()!s:>12} "
              f"{r['net_profit']:>8,.0f} {r['yoy_%']:>+6.0f}% {r['label']:>5} "
              f"{r['post5_ar_%']:>+8.2f}%")
    print("-" * 64)

    beat = events[events["label"] == "BEAT"]["post5_ar_%"]
    miss = events[events["label"] == "MISS"]["post5_ar_%"]
    print("\n  5-DAY POST-RESULTS ABNORMAL RETURN:")
    print(f"    BEAT quarters : {beat.mean():+.2f}%   (n={len(beat)}, "
          f"{(beat > 0).mean()*100:.0f}% positive)")
    print(f"    MISS quarters : {miss.mean():+.2f}%   (n={len(miss)}, "
          f"{(miss > 0).mean()*100:.0f}% positive)")
    print(f"    spread        : {beat.mean() - miss.mean():+.2f}%  (BEAT - MISS)")
    print("\n  A positive spread = the market keeps drifting in the profit-surprise")
    print(f"  direction (PEAD). Sample now {len(events)} quarters (was 5) — still modest,")
    print("  but the fundamentals unlock makes bigger, multi-stock tests possible.")
    print(f"\n  Saved -> {PROC}\\earnings_surprise.csv")


if __name__ == "__main__":
    main()
