"""
playbook.py
===============================================================================
"HOW DO I PLAY THIS STOCK AROUND ITS EVENTS?"

For each event type (EARNINGS and DIVIDEND ex-date) this scans every combination
of:
        ENTER  = how many trading days BEFORE the event you buy
        HOLD   = how many trading days you keep it

...and, across ALL past events, reports:
        WIN%   = fraction of times that trade made money   (your "chance")
        AVG%   = average profit per trade
        n      = how many historical events it was tested on

Then it picks the SWEET SPOT (highest win-rate play with a positive average) so
you get a concrete rule like:
        "Buy 5 days before earnings, hold 4 days -> won 62% of the time, +1.2% avg."

HONEST HEALTH WARNING
---------------------
This is measured on Reliance's OWN history (in-sample) and on a limited number
of events (72 earnings, 16 dividends). A high win-rate here is a HINT, not a
guarantee -- past behaviour need not repeat, and 16 dividends is a very small
sample. Trade real money only after out-of-sample + multi-stock confirmation.

    python playbook.py

Needs internet (Yahoo). Uses processed/unified_events.csv + corporate_actions_clean.csv.
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
COST = 0.001                                   # 0.10% round-trip, for the net avg

ENTER_RANGE = range(1, 9)                      # buy 1..8 trading days before event
HOLD_RANGE = range(1, 11)                      # hold 1..10 trading days

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = ssl.CERT_NONE


def fetch_adj(ticker: str) -> pd.Series:
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
           f"?period1=1167609600&period2={int(time.time())}&interval=1d")
    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=30, context=_SSL) as r:
        d = json.load(r)
    res = d["chart"]["result"][0]
    ind = res["indicators"]
    adj = ind.get("adjclose", [{}])[0].get("adjclose") or ind["quote"][0]["close"]
    return pd.Series(adj, index=pd.to_datetime(res["timestamp"], unit="s").normalize()).dropna().sort_index()


def scan(px: pd.Series, event_dates) -> dict:
    """For every (enter-days-before, hold-days) combo compute win-rate & avg return.

    Returns nested dict: results[enter][hold] = {win, avg, n}.
    """
    idx = px.index
    positions = [idx.searchsorted(pd.Timestamp(d).normalize()) for d in event_dates]
    out = {}
    for enter in ENTER_RANGE:
        out[enter] = {}
        for hold in HOLD_RANGE:
            rets = []
            for pos in positions:
                i_in = pos - enter
                i_out = i_in + hold
                if i_in < 0 or i_out >= len(idx):
                    continue
                rets.append(px.iloc[i_out] / px.iloc[i_in] - 1)
            if rets:
                s = pd.Series(rets)
                out[enter][hold] = {
                    "win": (s > 0).mean() * 100,
                    "avg": s.mean() * 100,
                    "net": (s.mean() - COST) * 100,
                    "n": len(s),
                }
    return out


def print_grid(results: dict, field: str, title: str, fmt: str) -> None:
    """Print a table: rows = enter-days-before, cols = hold-days."""
    print(f"\n  {title}")
    header = "  enter\\hold │" + "".join(f"{h:>6}" for h in HOLD_RANGE)
    print(header)
    print("  " + "-" * (len(header) - 2))
    for enter in ENTER_RANGE:
        cells = ""
        for hold in HOLD_RANGE:
            c = results[enter].get(hold)
            cells += f"{c[field]:>6.0f}" if c else "     -"
        print(f"  {enter:>3} d      │{cells}")


def best_plays(results: dict, min_frac=0.85, total=1) -> list:
    """Rank combos by win-rate (requiring a positive net average and enough events)."""
    rows = []
    for enter, holds in results.items():
        for hold, c in holds.items():
            if c["n"] >= min_frac * total and c["net"] > 0:
                rows.append((enter, hold, c["win"], c["avg"], c["net"], c["n"]))
    rows.sort(key=lambda r: (r[2], r[4]), reverse=True)     # win-rate, then net avg
    return rows


def report(px, name, dates):
    dates = list(dates)
    print("\n" + "=" * 70)
    print(f"  {name.upper()}  –  {len(dates)} historical events")
    print("=" * 70)
    res = scan(px, dates)
    total = len(dates)
    print_grid(res, "win", "CHANCE OF MAKING MONEY  (win %, rows=days before, cols=hold days)", "{:.0f}")
    print_grid(res, "avg", "AVERAGE RETURN  (%, gross)", "{:+.1f}")

    top = best_plays(res, total=total)
    print("\n  >> BEST PLAYS (ranked by historical win-rate, net of 0.1% cost):")
    if not top:
        print("     none with a positive edge.")
        return None
    print(f"     {'enter':>7} {'hold':>5} {'win%':>6} {'avg%':>7} {'net%':>7} {'n':>4}")
    for enter, hold, win, avg, net, n in top[:5]:
        print(f"     {enter:>4}d before {hold:>3}d  {win:>5.0f}% {avg:>+6.2f}% "
              f"{net:>+6.2f}% {n:>4}")
    e, h, win, avg, net, n = top[0]
    print(f"\n  RULE OF THUMB: buy ~{e} trading day(s) before {name.lower()}, hold ~{h} day(s)")
    print(f"  -> historically profitable {win:.0f}% of the time ({int(win/100*n)}/{n}), "
          f"avg {net:+.2f}% net.")
    return {"event": name, "enter": e, "hold": h, "win": win, "net": net, "n": n}


def main():
    uni = PROC / "unified_events.csv"
    ca = PROC / "corporate_actions_clean.csv"
    if not uni.exists():
        raise SystemExit("Run clean_events.py first.")

    ev = pd.read_csv(uni, parse_dates=["date"])
    earn = (ev[ev["event_type"].str.startswith("RESULT_", na=False)]["date"]
            .dt.normalize().drop_duplicates())
    div = []
    if ca.exists():
        c = pd.read_csv(ca, parse_dates=["exDate"])
        div = c[(c["event_type"] == "DIVIDEND")]["exDate"].dropna()

    print("Downloading prices ...")
    px = fetch_adj(TICKER)
    print(f"  {TICKER}: {len(px)} days ({px.index.min().date()} -> {px.index.max().date()})")

    plays = []
    r1 = report(px, "Earnings", earn)
    r2 = report(px, "Dividend", div)
    plays = [p for p in (r1, r2) if p]

    if plays:
        pd.DataFrame(plays).to_csv(PROC / "playbook.csv", index=False)
        print("\n" + "=" * 70)
        print("  YOUR PLAYBOOK (one line each):")
        for p in plays:
            print(f"   * {p['event']:9}: buy {p['enter']}d before, hold {p['hold']}d "
                  f"-> {p['win']:.0f}% win, {p['net']:+.2f}% avg")
        print("=" * 70)
    print("\n  !! In-sample, single stock, small samples (esp. dividends). A HINT to")
    print("     test out-of-sample, NOT a guaranteed money machine.")
    print(f"\n  Saved -> {PROC}\\playbook.csv")


if __name__ == "__main__":
    main()
