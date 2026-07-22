"""
event_behaviour.py
===============================================================================
HOW DOES A STOCK MOVE AROUND ITS CORPORATE EVENTS — AND HOW MANY DAYS BEFORE
THE EVENT SHOULD YOU BUY?
-------------------------------------------------------------------------------
For every historical event (financial results / board meeting / corporate action
/ announcement) we test:

    ENTRY  : buy at CLOSE, T-1 .. T-N trading days BEFORE the event   (N configurable)
    EXIT   : sell at CLOSE, T+1 .. T+M trading days AFTER  the event   (M configurable)

so the profit reflects how the stock reacts AFTER the news, not just up to the
event day. For each (days_before, days_after) pair we record return% and return₹,
then aggregate across ALL events per event_type to find the single BEST
"buy N days before / sell P days after" rule.

DATA (uses what's already in the project)
    events : the downloaded NSE feeds in each <SYMBOL>/ folder
             (financial_results / board_meetings / corporate_actions / announcements)
    prices : daily ADJUSTED close from Yahoo Finance (cached in processed/price_cache/)

    ...or bring your own CSVs:
        python event_behaviour.py --prices prices.csv --events events.csv
        prices.csv : date,symbol,open,high,low,close,volume
        events.csv : symbol,event_date,event_type

OUTPUT (in processed/)
    event_behaviour_detail.csv   one row per (event, days_before, days_after)
    event_behaviour_summary.csv  the aggregated grid + the best rule per event_type
    + a printed plain-English summary.

RUN
    python event_behaviour.py                 # all stocks with downloaded feeds
    python event_behaviour.py RELIANCE INFY   # specific symbols
===============================================================================
"""

from __future__ import annotations

import argparse
import json
import ssl
import sys
import time
import urllib.request
from urllib.parse import quote
from pathlib import Path

import pandas as pd

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                              # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
PRICE_CACHE = PROC / "price_cache"
PROC.mkdir(exist_ok=True)
PRICE_CACHE.mkdir(exist_ok=True)

# ── configurable knobs ──────────────────────────────────────────────────────
N_BEFORE = 8          # test entry T-1 .. T-N trading days before the event
N_AFTER = 8           # test exit  T+1 .. T+M trading days after the event
COST = 0.0            # round-trip cost fraction (0 = gross, e.g. 0.001 = 0.1%)
MIN_EVENTS = 10       # ignore aggregate cells with fewer than this many events

# event feeds -> (event_type label, date column, optional signal filter)
FEEDS = {
    "financial_results": ("RESULTS", "broadCastDate", None),
    "board_meetings":    ("BOARD_MEETING", "bm_date", None),
    "corporate_actions": ("CORPORATE_ACTION", "exDate", None),
    "announcements":     ("ANNOUNCEMENT", "sort_date", "desc"),
}
# only these announcement categories are treated as "major" signal
SIGNAL_DESC = {
    "Press Release", "Media Release", "Acquisition", "Financial Result Updates",
    "Investor Presentation", "Credit Rating",
    "Analysts/Institutional Investor Meet/Con. Call Updates",
    "Allotment of Securities", "Disclosure under SEBI Takeover Regulations",
    "Outcome of Board Meeting", "Dividend", "Bonus issue",
}

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = ssl.CERT_NONE


# ═══════════════════════════════════════════════════════════════════════════
# DATA SOURCES
# ═══════════════════════════════════════════════════════════════════════════
def discover_feed_symbols() -> list[str]:
    """Symbols whose <sym>/ folder holds at least one event feed CSV."""
    out = []
    for p in ROOT.iterdir():
        if p.is_dir() and any((p / f"{k}.csv").exists() for k in FEEDS):
            out.append(p.name)
    return sorted(out)


def load_events_from_feeds(symbols: list[str]) -> pd.DataFrame:
    """Build a tidy [symbol, event_date, event_type] table from the NSE feeds."""
    rows = []
    for sym in symbols:
        for feed, (label, datecol, sigcol) in FEEDS.items():
            path = ROOT / sym / f"{feed}.csv"
            if not path.exists():
                continue
            df = pd.read_csv(path, dtype=str).fillna("")
            if datecol not in df.columns:
                continue
            if sigcol and sigcol in df.columns:          # keep only 'major' announcements
                df = df[df[sigcol].isin(SIGNAL_DESC)]
            dates = pd.to_datetime(df[datecol], errors="coerce", dayfirst=True)
            for d in dates.dropna():
                rows.append({"symbol": sym, "event_date": d.normalize(),
                             "event_type": label})
    if not rows:
        return pd.DataFrame(columns=["symbol", "event_date", "event_type"])
    ev = pd.DataFrame(rows).drop_duplicates()
    return ev.sort_values(["symbol", "event_date"]).reset_index(drop=True)


def load_events_from_csv(path: str) -> pd.DataFrame:
    df = pd.read_csv(path)
    cmap = {c.lower(): c for c in df.columns}
    sym = cmap.get("symbol"); dt = cmap.get("event_date") or cmap.get("date")
    et = cmap.get("event_type") or cmap.get("type")
    ev = pd.DataFrame({
        "symbol": df[sym].astype(str),
        "event_date": pd.to_datetime(df[dt], errors="coerce", dayfirst=True).dt.normalize(),
        "event_type": df[et].astype(str).str.upper(),
    }).dropna(subset=["event_date"])
    return ev.sort_values(["symbol", "event_date"]).reset_index(drop=True)


def fetch_prices_yahoo(symbol: str) -> pd.Series | None:
    """Daily ADJUSTED close from Yahoo for <symbol>.NS, cached to disk."""
    cache = PRICE_CACHE / f"{symbol}.csv"
    if cache.exists():
        s = pd.read_csv(cache, parse_dates=["date"]).set_index("date")["adj"]
        return s.sort_index()
    # "M&M" must be percent-encoded, or the bare "&" ends the URL path and the
    # request silently becomes a lookup for "M".
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{quote(symbol, safe='')}.NS"
           f"?period1=1167609600&period2={int(time.time())}&interval=1d")
    try:
        req = urllib.request.Request(url, headers=_HEADERS)
        with urllib.request.urlopen(req, timeout=30, context=_SSL) as r:
            d = json.load(r)
        res = d["chart"]["result"][0]
        ind = res["indicators"]
        adj = ind.get("adjclose", [{}])[0].get("adjclose") or ind["quote"][0]["close"]
        s = pd.Series(adj, index=pd.to_datetime(res["timestamp"], unit="s").normalize())
        s = s.dropna().sort_index()
        s.rename("adj").rename_axis("date").to_frame().to_csv(cache)
        time.sleep(0.3)
        return s
    except Exception as e:                       # noqa: BLE001
        print(f"    ! price fetch failed for {symbol}: {type(e).__name__}: {e}")
        return None


def load_prices_from_csv(path: str) -> dict[str, pd.Series]:
    df = pd.read_csv(path)
    cmap = {c.lower(): c for c in df.columns}
    sym, dt = cmap.get("symbol"), (cmap.get("date") or cmap.get("trade_date"))
    close = cmap.get("close") or cmap.get("adj_close")
    df[dt] = pd.to_datetime(df[dt], errors="coerce", dayfirst=True).dt.normalize()
    out = {}
    for s, g in df.dropna(subset=[dt]).groupby(sym):
        out[str(s)] = (g.set_index(dt)[close].sort_index()
                        .astype(float).rename("adj"))
    return out


# ═══════════════════════════════════════════════════════════════════════════
# THE ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════
def analyse_event(px: pd.Series, event_date, symbol, event_type) -> list[dict]:
    """One event -> rows for every (days_before entry, days_after exit) pair."""
    idx = px.index
    pos = idx.searchsorted(pd.Timestamp(event_date))     # first trading day on/after
    if pos >= len(idx):
        return []
    rows = []
    for db in range(1, N_BEFORE + 1):
        i_in = pos - db
        if i_in < 0:
            continue
        entry = float(px.iloc[i_in])
        if entry <= 0:
            continue
        for da in range(1, N_AFTER + 1):
            i_out = pos + da
            if i_out >= len(idx):
                continue
            exit_px = float(px.iloc[i_out])
            ret_pct = (exit_px / entry - 1 - COST) * 100
            rows.append({
                "symbol": symbol,
                "event_date": pd.Timestamp(event_date).date(),
                "event_type": event_type,
                "days_before_entry": db,
                "entry_price": round(entry, 2),
                "post_event_day": da,
                "post_event_price": round(exit_px, 2),
                "return_pct": round(ret_pct, 3),
                "return_rs": round(exit_px - entry, 2),
                "win_loss": "WIN" if ret_pct > 0 else ("LOSS" if ret_pct < 0 else "FLAT"),
            })
    return rows


def aggregate(detail: pd.DataFrame) -> pd.DataFrame:
    """Grid of avg return% + win-rate per (event_type, days_before, days_after)."""
    g = detail.groupby(["event_type", "days_before_entry", "post_event_day"])
    agg = g.agg(
        n_events=("return_pct", "size"),
        avg_return_pct=("return_pct", "mean"),
        median_return_pct=("return_pct", "median"),
        win_rate_pct=("win_loss", lambda s: (s == "WIN").mean() * 100),
        avg_return_rs=("return_rs", "mean"),
    ).reset_index()
    for c in ("avg_return_pct", "median_return_pct", "win_rate_pct", "avg_return_rs"):
        agg[c] = agg[c].round(3)
    return agg


def best_rules(agg: pd.DataFrame) -> pd.DataFrame:
    """Pick the single best (days_before, days_after) per event_type and overall."""
    rows = []
    scope = agg[agg["n_events"] >= MIN_EVENTS]
    for et, sub in list(scope.groupby("event_type")) + [("__ALL__", scope)]:
        if sub.empty:
            continue
        # rank by average return, then win rate
        best = sub.sort_values(["avg_return_pct", "win_rate_pct"],
                               ascending=False).iloc[0]
        rows.append({
            "event_type": et,
            "best_days_before": int(best["days_before_entry"]),
            "best_days_after": int(best["post_event_day"]),
            "avg_return_pct": round(best["avg_return_pct"], 3),
            "win_rate_pct": round(best["win_rate_pct"], 1),
            "avg_return_rs": round(best["avg_return_rs"], 2),
            "n_events": int(best["n_events"]),
        })
    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("symbols", nargs="*", help="symbols to run (default: all with feeds)")
    ap.add_argument("--prices", help="custom price CSV (date,symbol,...,close)")
    ap.add_argument("--events", help="custom event CSV (symbol,event_date,event_type)")
    args = ap.parse_args()

    # ── events ──────────────────────────────────────────────────────────────
    if args.events:
        events = load_events_from_csv(args.events)
        print(f"Events from {args.events}: {len(events):,}")
    else:
        symbols = args.symbols or discover_feed_symbols()
        if not symbols:
            raise SystemExit("No downloaded event feeds found. Download some in the "
                             "stock browser first, or pass --events a CSV.")
        events = load_events_from_feeds(symbols)
        print(f"Events from NSE feeds ({len(symbols)} stocks): {len(events):,}")
    print(events["event_type"].value_counts().to_string())

    syms = sorted(events["symbol"].unique())

    # ── prices ──────────────────────────────────────────────────────────────
    if args.prices:
        price_map = load_prices_from_csv(args.prices)
        print(f"\nPrices from {args.prices}: {len(price_map)} symbols")
    else:
        print(f"\nFetching daily prices for {len(syms)} symbols (Yahoo, cached)...")
        price_map = {}
        for i, s in enumerate(syms, 1):
            px = fetch_prices_yahoo(s)
            if px is not None and len(px) > N_BEFORE + N_AFTER + 5:
                price_map[s] = px
            print(f"  [{i}/{len(syms)}] {s}: "
                  f"{'ok ' + str(len(px)) + ' days' if px is not None else 'no data'}")

    # ── analyse every event ─────────────────────────────────────────────────
    detail_rows = []
    used = 0
    for _, e in events.iterrows():
        px = price_map.get(e["symbol"])
        if px is None:
            continue
        r = analyse_event(px, e["event_date"], e["symbol"], e["event_type"])
        if r:
            detail_rows.extend(r)
            used += 1
    if not detail_rows:
        raise SystemExit("No analysable events (price history didn't cover them).")

    detail = pd.DataFrame(detail_rows)
    detail.to_csv(PROC / "event_behaviour_detail.csv", index=False)

    agg = aggregate(detail)
    best = best_rules(agg)
    # tag the best cell inside the grid too
    agg["is_best"] = False
    for _, b in best[best["event_type"] != "__ALL__"].iterrows():
        m = ((agg["event_type"] == b["event_type"]) &
             (agg["days_before_entry"] == b["best_days_before"]) &
             (agg["post_event_day"] == b["best_days_after"]))
        agg.loc[m, "is_best"] = True
    agg.to_csv(PROC / "event_behaviour_summary.csv", index=False)

    # ── printed summary ─────────────────────────────────────────────────────
    print(f"\nAnalysed {used:,} events across {len(price_map)} stocks "
          f"-> {len(detail):,} entry/exit combinations.")
    print("\n" + "=" * 78)
    print("  BEST 'BUY BEFORE / SELL AFTER' RULE  (ranked by avg return, then win rate)")
    print("=" * 78)
    for _, b in best.iterrows():
        name = "ALL EVENTS" if b["event_type"] == "__ALL__" else b["event_type"]
        print(f"\n  {name}:")
        print(f"    Buy {b['best_days_before']} trading day(s) BEFORE the event, "
              f"sell {b['best_days_after']} day(s) AFTER")
        print(f"    -> avg profit {b['avg_return_pct']:+.2f}%  "
              f"(₹{b['avg_return_rs']:+.2f}/share),  win rate {b['win_rate_pct']:.0f}%,  "
              f"n={b['n_events']} events")
    print("\n" + "=" * 78)
    print("  Plain English:")
    for _, b in best.iterrows():
        if b["event_type"] == "__ALL__":
            continue
        print(f"   For {b['event_type']}, buying {b['best_days_before']} day(s) before the "
              f"event gives the best average profit of {b['avg_return_pct']:+.2f}% "
              f"with a win rate of {b['win_rate_pct']:.0f}% (selling "
              f"{b['best_days_after']} day(s) after).")
    print("=" * 78)
    print(f"\nSaved:\n  {PROC}\\event_behaviour_detail.csv   (per-event rows)")
    print(f"  {PROC}\\event_behaviour_summary.csv  (aggregated grid + best cells)")


if __name__ == "__main__":
    main()
