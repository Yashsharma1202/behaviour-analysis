"""
dividend_recovery.py
===============================================================================
DIVIDEND RECOVERY ANALYSIS  +  raw before/after behaviour.

THE QUESTION (your ask)
-----------------------
When a stock goes EX-DIVIDEND it gets "discounted": on the ex-date it opens
lower, roughly by the dividend amount, because the cash is leaving the company.

    Day before ex (cum-dividend):  close = 2900   <- you still own the dividend
    Ex-date (ex-dividend):         open  = 2894   <- ~Rs 5.5 dividend knocked off

This script measures, for EVERY Reliance dividend:
    1. HOW MUCH the stock dropped on the ex-date (the "discount").
    2. HOW MANY trading days it took to RECOVER back to the pre-dividend price.
    3. The raw price behaviour in the days BEFORE and AFTER the event.

WHY RAW (unadjusted) PRICES HERE?
---------------------------------
event_study.py used ADJUSTED prices, which delete the ex-date drop (they pretend
the dividend was reinvested). To SEE and MEASURE the drop we must use the RAW,
as-traded price. So this is the deliberate opposite choice from the event study.

    python dividend_recovery.py

Needs internet (Yahoo Finance). Uses processed/corporate_actions_clean.csv.
===============================================================================
"""

from __future__ import annotations

import json
import ssl
import time
import urllib.request
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "processed"
CA = OUT_DIR / "corporate_actions_clean.csv"

TICKER = "RELIANCE.NS"
MAX_RECOVERY_DAYS = 120        # give up looking for recovery after this many days
BEFORE, AFTER = 5, 5           # window for the raw before/after return picture

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = ssl.CERT_NONE


# ---------------------------------------------------------------------------
def fetch_ohlc(ticker: str) -> pd.DataFrame:
    """Download RAW (unadjusted) daily open/close. Index = date, cols open/close."""
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
           f"?period1=1167609600&period2={int(time.time())}&interval=1d")
    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=30, context=_SSL) as r:
        data = json.load(r)
    res = data["chart"]["result"][0]
    q = res["indicators"]["quote"][0]
    df = pd.DataFrame(
        {"open": q["open"], "close": q["close"]},
        index=pd.to_datetime(res["timestamp"], unit="s").normalize(),
    ).dropna().sort_index()
    return df


# ---------------------------------------------------------------------------
def analyse_dividend(px: pd.DataFrame, ex_date: pd.Timestamp, div: float) -> dict | None:
    """Measure the ex-date discount and recovery time for one dividend.

    Input:  price frame, the ex-date, the dividend per share (Rs).
    Output: dict of measurements, or None if the date is outside our price range.
    """
    dates = px.index
    # first trading day on/after the stated ex-date = the day it trades ex-dividend
    pos = dates.searchsorted(pd.Timestamp(ex_date).normalize())
    if pos <= 0 or pos >= len(dates) - 1:
        return None

    p_before = px["close"].iloc[pos - 1]        # cum-dividend close (still owns div)
    ex_open = px["open"].iloc[pos]              # first ex-dividend price
    ex_close = px["close"].iloc[pos]

    # the ex-date gap-down at the open, and where it closed the day
    gap_pct = (ex_open / p_before - 1) * 100
    exday_pct = (ex_close / p_before - 1) * 100
    div_yield = div / p_before * 100 if div else float("nan")

    # RECOVERY: first trading day AFTER the ex-open whose close regains p_before
    recovery_days = None
    horizon = min(pos + MAX_RECOVERY_DAYS, len(dates))
    for k in range(pos, horizon):
        if px["close"].iloc[k] >= p_before:
            recovery_days = k - pos             # 0 = recovered on the ex-date itself
            break

    return {
        "ex_date": dates[pos].date(),
        "dividend": div,
        "p_before": round(p_before, 1),
        "div_yield_%": round(div_yield, 2),
        "ex_gap_%": round(gap_pct, 2),
        "ex_close_%": round(exday_pct, 2),
        "recovery_days": recovery_days,             # None => not recovered in horizon
    }


def raw_window(px: pd.DataFrame, dates: pd.Series, label: str) -> None:
    """Average RAW cumulative return path around a set of event dates (behaviour)."""
    close = px["close"]
    idx = px.index
    paths = []
    for ev in dates:
        pos = idx.searchsorted(pd.Timestamp(ev).normalize())
        if pos - BEFORE < 0 or pos + AFTER >= len(idx):
            continue
        window = close.iloc[pos - BEFORE: pos + AFTER + 1]
        rebased = (window.values / window.values[BEFORE] - 1) * 100  # 0% at event day
        paths.append(rebased)
    if not paths:
        return
    mean = pd.DataFrame(paths).mean()
    print(f"\n  Raw average price path around {label} ({len(paths)} events, 0% = event day):")
    for i, val in enumerate(mean):
        rel = i - BEFORE
        bar = "#" * int(abs(val) * 6)
        side = ">" if val >= 0 else "<"
        tag = "  <-- event" if rel == 0 else ""
        print(f"    t{rel:+d} {side}{bar:<20} {val:+5.2f}%{tag}")


# ---------------------------------------------------------------------------
def main() -> None:
    if not CA.exists():
        raise SystemExit("Run clean_events.py first (need corporate_actions_clean.csv)")

    ca = pd.read_csv(CA)
    divs = ca[(ca["event_type"] == "DIVIDEND") & ca["amount"].notna()].copy()
    divs["exDate"] = pd.to_datetime(divs["exDate"])
    divs = divs.dropna(subset=["exDate"]).sort_values("exDate")

    print("Downloading RAW prices from Yahoo Finance ...")
    px = fetch_ohlc(TICKER)
    print(f"  {TICKER}: {len(px)} days ({px.index.min().date()} -> {px.index.max().date()})\n")

    rows = []
    for _, d in divs.iterrows():
        r = analyse_dividend(px, d["exDate"], float(d["amount"]))
        if r:
            rows.append(r)
    res = pd.DataFrame(rows)
    res.to_csv(OUT_DIR / "dividend_recovery.csv", index=False)

    # ---- per-dividend table ----------------------------------------------
    print("DIVIDEND RECOVERY  (how much the stock dropped & how long to climb back)")
    print("-" * 78)
    print(f"  {'ex-date':>11} {'div':>5} {'pre-px':>8} {'yield%':>7} "
          f"{'gap%':>6} {'exclose%':>9} {'recovery':>10}")
    print("-" * 78)
    for _, r in res.iterrows():
        rec = ("same day" if r["recovery_days"] == 0
               else f"{int(r['recovery_days'])}d" if pd.notna(r["recovery_days"])
               else f">{MAX_RECOVERY_DAYS}d")
        print(f"  {r['ex_date']!s:>11} {r['dividend']:>5.2f} {r['p_before']:>8.0f} "
              f"{r['div_yield_%']:>6.2f}% {r['ex_gap_%']:>+6.2f} {r['ex_close_%']:>+8.2f}% "
              f"{rec:>10}")
    print("-" * 78)

    # ---- summary ----------------------------------------------------------
    recovered = res[res["recovery_days"].notna()]
    print("\nSUMMARY")
    print(f"  dividends analysed        : {len(res)}")
    print(f"  avg dividend yield        : {res['div_yield_%'].mean():.2f}%  "
          f"(theoretical ex-date drop)")
    print(f"  avg actual ex-date gap    : {res['ex_gap_%'].mean():+.2f}%  (at the open)")
    print(f"  avg ex-date close move    : {res['ex_close_%'].mean():+.2f}%")
    if not recovered.empty:
        med = recovered["recovery_days"].median()
        print(f"  recovered within {MAX_RECOVERY_DAYS}d      : "
              f"{len(recovered)}/{len(res)}")
        print(f"  median recovery time      : {med:.0f} trading days")
        print(f"  fastest / slowest         : {recovered['recovery_days'].min():.0f}d / "
              f"{recovered['recovery_days'].max():.0f}d")
    print("\n  INTERPRETATION: if the ex-gap is SMALLER than the dividend yield, the")
    print("  market was already bidding the stock up (dividend capture). A short median")
    print("  recovery time means the 'discount' is a temporary dip, not real value lost.")

    # ---- raw behaviour before/after dividends AND results -----------------
    raw_window(px, divs["exDate"], "DIVIDENDS (ex-date)")
    uni = OUT_DIR / "unified_events.csv"
    if uni.exists():
        ev = pd.read_csv(uni, parse_dates=["date"])
        earnings = ev[ev["event_type"].str.startswith("RESULT_", na=False)]["date"]
        raw_window(px, earnings, "EARNINGS (results day)")

    print(f"\nSaved -> {OUT_DIR}\\dividend_recovery.csv")


if __name__ == "__main__":
    main()
