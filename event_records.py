"""
event_records.py
===============================================================================
Builds a detailed HISTORY LOG of every Reliance corporate event: WHEN it
happened, WHAT it was, HOW the stock behaved around it (5 days before / after),
and the ex-date PRICE ADJUSTMENT.

ADJUSTMENT LOGIC (important)
---------------------------
  * Dividend : Yahoo does NOT dividend-adjust its 'close', so the small ex-date
               drop is real -> we show the ACTUAL raw price change.
  * Bonus/Split : Yahoo already adjusts 'close' for these, so no drop appears in
               the data. We show the THEORETICAL adjustment from the ratio
               (a 1:1 bonus mechanically halves the price -> -50%).
  * Rights/Demerger/Other : show the actual as-traded change.

Returns before/after use ADJUSTED close (clean total return).

Output: processed/event_records.csv (newest first) -> read by the GUI 'Records' tab.

    python event_records.py

Needs internet (Yahoo). Uses processed/corporate_actions_clean.csv +
unified_events.csv + earnings_surprise.csv.
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

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                              # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
TICKER = "RELIANCE.NS"
W = 5

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = ssl.CERT_NONE


def fetch() -> pd.DataFrame:
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{TICKER}"
           f"?period1=1167609600&period2={int(time.time())}&interval=1d")
    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=30, context=_SSL) as r:
        d = json.load(r)
    res = d["chart"]["result"][0]
    q = res["indicators"]["quote"][0]
    adj = res["indicators"]["adjclose"][0]["adjclose"]
    return pd.DataFrame({"raw": q["close"], "adj": adj},
                        index=pd.to_datetime(res["timestamp"], unit="s").normalize()).dropna().sort_index()


def around(px, date):
    """(before_5d %, after_5d %, ex-day raw change %) around a date, or NaNs."""
    idx = px.index
    pos = idx.searchsorted(pd.Timestamp(date).normalize())
    if pos <= W or pos + W >= len(idx):
        return float("nan"), float("nan"), float("nan")
    a = px["adj"].values
    r = px["raw"].values
    before = (a[pos] / a[pos - W] - 1) * 100
    after = (a[pos + W] / a[pos] - 1) * 100
    exday = (r[pos] / r[pos - 1] - 1) * 100
    return before, after, exday


def bonus_theoretical(ratio: str) -> float:
    """Ratio 'a:b' (a new shares per b held) -> theoretical price adjustment %."""
    try:
        a, b = (float(x) for x in ratio.split(":"))
        return (b / (a + b) - 1) * 100          # 1:1 -> -50%
    except Exception:                            # noqa: BLE001
        return float("nan")


def main():
    px = fetch()
    records = []

    # --- corporate actions -------------------------------------------------
    ca = pd.read_csv(PROC / "corporate_actions_clean.csv", parse_dates=["exDate"])
    for _, r in ca.dropna(subset=["exDate"]).iterrows():
        before, after, exday_raw = around(px, r["exDate"])
        et = r["event_type"]
        if et == "DIVIDEND":
            adj = exday_raw                      # actual small drop
            detail = f"Rs {r['amount']:.2f}/share" if pd.notna(r["amount"]) else ""
        elif et in ("BONUS", "SPLIT"):
            adj = bonus_theoretical(str(r["ratio"]))     # theoretical
            detail = f"ratio {r['ratio']}"
        else:
            adj = exday_raw
            detail = str(r.get("ratio") or r.get("amount") or "")
        records.append({
            "date": r["exDate"], "category": et,
            "event": str(r["subject"]).strip()[:48],
            "detail": detail, "before_%": before, "adj_%": adj, "after_%": after,
        })

    # --- earnings (attach YoY label where we have it) ---------------------
    sur = PROC / "earnings_surprise.csv"
    yoy = {}
    if sur.exists():
        s = pd.read_csv(sur, parse_dates=["event_date"])
        yoy = {pd.Timestamp(d).normalize(): (lab, g)
               for d, lab, g in zip(s["event_date"], s["label"], s["yoy_%"])}
    ev = pd.read_csv(PROC / "unified_events.csv", parse_dates=["date"])
    edates = (ev[ev["event_type"].str.startswith("RESULT_", na=False)]["date"]
              .dt.normalize().drop_duplicates())
    for d in edates:
        before, after, _ = around(px, d)
        if pd.isna(before):
            continue
        lab, g = yoy.get(d, (None, None))
        detail = f"{lab} ({g:+.0f}% YoY)" if lab else ""
        records.append({
            "date": d, "category": "EARNINGS", "event": "Quarterly results",
            "detail": detail, "before_%": before, "adj_%": float("nan"),
            "after_%": after,
        })

    df = pd.DataFrame(records).dropna(subset=["before_%"])
    df = df.sort_values("date", ascending=False).reset_index(drop=True)
    df.to_csv(PROC / "event_records.csv", index=False)

    print(f"Built {len(df)} event records "
          f"({df['date'].min().date()} -> {df['date'].max().date()})\n")
    print(f"  {'date':>11} {'category':>9} {'before':>7} {'adj':>7} {'after':>7}  event")
    for _, r in df.head(18).iterrows():
        adj = f"{r['adj_%']:+.1f}%" if pd.notna(r["adj_%"]) else "   -"
        print(f"  {r['date'].date()!s:>11} {r['category']:>9} {r['before_%']:>+6.1f}% "
              f"{adj:>7} {r['after_%']:>+6.1f}%  {r['event'][:34]}")
    print(f"\nSaved -> {PROC}\\event_records.csv")


if __name__ == "__main__":
    main()
