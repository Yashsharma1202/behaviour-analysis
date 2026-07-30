"""
download_feeds.py
===============================================================================
Bulk-download NSE event feeds (announcements / board meetings / corporate
actions / financial results) for a set of stocks, so the behaviour analysis can
run on them.

    python download_feeds.py nifty50          # the 50 Nifty constituents (live from NSE)
    python download_feeds.py RELIANCE TCS INFY # specific symbols
    python download_feeds.py                   # defaults to nifty50

It is POLITE (waits between stocks) and RESUMABLE (skips stocks already done),
so if NSE rate-limits you, just run it again and it continues where it left off.
Feeds are saved into a <SYMBOL>/ folder each.
===============================================================================
"""

from __future__ import annotations

import json
import sys
import time
import urllib.request
from pathlib import Path

import stock_server as S          # reuse the NSE fetcher + session handling

# Fallback Nifty-50 list (used only if the live fetch from NSE fails).
# Current Nifty 50 constituents, synced 2026-07-23 from the official NSE export
# (new data/08_Company_Directory.xlsx). Exactly 50 symbols.
# Ex-members removed in this sync (still on disk under their <SYMBOL>/ folders,
# just no longer part of the index): BPCL, BRITANNIA, HEROMOTOCO, INDUSINDBK,
# LTIM, TATAMOTORS. Added this sync: ETERNAL (ex-Zomato), INDIGO, MAXHEALTH,
# TMPV (Tata Motors Passenger Vehicles).
NIFTY50_FALLBACK = [
    "ADANIENT", "ADANIPORTS", "APOLLOHOSP", "ASIANPAINT", "AXISBANK",
    "BAJAJ-AUTO", "BAJAJFINSV", "BAJFINANCE", "BEL", "BHARTIARTL",
    "CIPLA", "COALINDIA", "DRREDDY", "EICHERMOT", "ETERNAL",
    "GRASIM", "HCLTECH", "HDFCBANK", "HDFCLIFE", "HINDALCO",
    "HINDUNILVR", "ICICIBANK", "INDIGO", "INFY", "ITC", "JIOFIN",
    "JSWSTEEL", "KOTAKBANK", "LT", "M&M", "MARUTI", "MAXHEALTH", "NESTLEIND",
    "NTPC", "ONGC", "POWERGRID", "RELIANCE", "SBILIFE", "SBIN", "SHRIRAMFIN",
    "SUNPHARMA", "TATACONSUM", "TATASTEEL", "TCS", "TECHM",
    "TITAN", "TMPV", "TRENT", "ULTRACEMCO", "WIPRO",
]

WAIT_BETWEEN = 2.0        # seconds between stocks (be polite to NSE)


def get_nifty50() -> list[str]:
    """Live Nifty-50 constituents from NSE, with a built-in fallback."""
    try:
        op = S._nse_prime("NIFTY")
        url = "https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%2050"
        raw = op.open(urllib.request.Request(url, headers=S._NSE_HEADERS),
                      timeout=25).read().decode("utf-8", "replace").strip()
        data = json.loads(raw)
        syms = sorted({r["symbol"] for r in data["data"]
                       if r.get("symbol") and r["symbol"] != "NIFTY 50"})
        if len(syms) >= 40:
            print(f"Nifty 50 fetched live from NSE: {len(syms)} stocks")
            return syms
    except Exception as e:                       # noqa: BLE001
        print(f"Live Nifty-50 fetch failed ({e}); using built-in list.")
    return NIFTY50_FALLBACK


def already_done(sym: str) -> bool:
    """Treat a stock as done if its financial_results feed is already saved."""
    return (S.ROOT / sym / "financial_results.csv").exists()


def main():
    args = sys.argv[1:]
    force = "--force" in args
    args = [a for a in args if a != "--force"]
    if not args or args[0].lower() in ("nifty50", "nifty"):
        symbols = get_nifty50()
    else:
        symbols = [a.upper() for a in args]

    print(f"\nDownloading NSE event feeds for {len(symbols)} stocks "
          f"(pausing {WAIT_BETWEEN}s between each)\n" + "-" * 66)

    done = skipped = failed = 0
    for i, s in enumerate(symbols, 1):
        if already_done(s) and not force:
            print(f"  [{i:>2}/{len(symbols)}] {s:<12} already downloaded — skip")
            skipped += 1
            continue
        try:
            res = S.fetch_nse_feeds(s)
            c = res["counts"]
            ok = "ok " if res["any"] else "EMPTY"
            print(f"  [{i:>2}/{len(symbols)}] {s:<12} {ok}  "
                  f"ann={c.get('announcements', 0)} board={c.get('board_meetings', 0)} "
                  f"corp={c.get('corporate_actions', 0)} results={c.get('financial_results', 0)}")
            done += 1 if res["any"] else 0
            failed += 0 if res["any"] else 1
        except Exception as e:                   # noqa: BLE001
            print(f"  [{i:>2}/{len(symbols)}] {s:<12} FAILED: {type(e).__name__}: {e}")
            failed += 1
        time.sleep(WAIT_BETWEEN)

    print("-" * 66)
    print(f"Done. downloaded {done}, skipped {skipped}, failed/empty {failed}.")
    if failed:
        print("Some stocks failed (usually NSE rate-limiting). Just run the command "
              "again — it will skip the ones already done and retry the rest.")
    print("\nNext:")
    print("  python results_tracker.py          # 6-days-before / result-day / 3-days-after")
    print("  python event_behaviour.py          # full best buy-before/sell-after analysis")
    print("  python stock_server.py 8080        # or browse the Behaviour tab in the dashboard")


if __name__ == "__main__":
    main()
