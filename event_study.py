"""
event_study.py
===============================================================================
THE DESTINATION of the "behaviour analysis" project.

WHAT THIS ANSWERS
-----------------
"How does RELIANCE's stock actually BEHAVE in the days around a corporate event
 (a dividend ex-date, or an earnings release)?"

The event dates come from processed/unified_events.csv (built by clean_events.py).
The prices are downloaded live from Yahoo Finance. We then run a classic
EVENT STUDY: line every event up at day 0, measure the stock's MARKET-ADJUSTED
abnormal return on each of the surrounding days (-5 .. +5), and average across
all events of the same type to see the typical "behaviour signature".

WHY MARKET-ADJUSTED?
--------------------
Every day the whole market (Nifty 50) moves. If RELIANCE rose 2% on earnings day
but Nifty also rose 2%, the stock did nothing special. So:

        Abnormal Return (AR) = R_reliance  -  R_nifty50

This strips out market-wide moves and isolates the event's own effect.
(A more advanced version fits a beta:  AR = R_stock - (alpha + beta*R_mkt).
 We use the simpler market-adjusted model here — robust and easy to read.)

NOTE ON PRICES: we use ADJUSTED close, which already folds in dividends & the
2024 1:1 bonus. So the AR you see is the market's genuine RE-RATING of the stock
on news, not the mechanical ex-date price drop.

HOW TO RUN
----------
    python event_study.py

Needs internet (Yahoo Finance). Outputs go to ./processed/.
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
UNIFIED = OUT_DIR / "unified_events.csv"

# Yahoo tickers:  .NS = NSE listing ;  ^NSEI = Nifty 50 index (our market benchmark)
STOCK_TICKER = "RELIANCE.NS"
MARKET_TICKER = "^NSEI"

WINDOW = 5              # trading days on each side of the event (t-5 .. t+5)
EST_START = "2007-01-01"   # ignore events before Yahoo has reliable Nifty data

# Yahoo blocks requests without a browser-like User-Agent.
_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
# Some corporate networks have broken SSL chains; don't let that stop the demo.
_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode = ssl.CERT_NONE


# ---------------------------------------------------------------------------
# 1. PRICE DOWNLOAD  (Yahoo chart API -> pandas Series of adjusted closes)
# ---------------------------------------------------------------------------
def fetch_prices(ticker: str) -> pd.Series:
    """Download the full daily adjusted-close history for one ticker.

    Input:  a Yahoo ticker string, e.g. 'RELIANCE.NS'.
    Output: a pandas Series indexed by date (adjusted close), sorted ascending.
    """
    period1 = 1167609600                    # 2007-01-01 in unix seconds
    period2 = int(time.time())              # now
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
           f"?period1={period1}&period2={period2}&interval=1d")

    req = urllib.request.Request(url, headers=_HEADERS)
    with urllib.request.urlopen(req, timeout=30, context=_SSL) as resp:
        data = json.load(resp)

    result = data["chart"]["result"][0]
    timestamps = result["timestamp"]
    # adjusted close survives splits/bonuses/dividends; fall back to raw close
    adj = result.get("indicators", {}).get("adjclose", [{}])[0].get("adjclose")
    if adj is None:
        adj = result["indicators"]["quote"][0]["close"]

    s = pd.Series(adj, index=pd.to_datetime(timestamps, unit="s").normalize())
    return s.dropna().sort_index()


# ---------------------------------------------------------------------------
# 2. THE EVENT STUDY ITSELF
# ---------------------------------------------------------------------------
def run_event_study(event_dates: pd.Series, ar: pd.Series, label: str) -> pd.DataFrame:
    """Line every event up at day 0 and collect the abnormal return per relative day.

    Input:
        event_dates : Series of event Timestamps.
        ar          : daily abnormal-return Series indexed by trading date.
        label       : name of the event type (for messages).
    Output:
        DataFrame  rows = events, columns = relative days -5..+5 (abnormal returns).
    """
    trading_days = ar.index
    rows = []
    for ev in event_dates:
        ev = pd.Timestamp(ev).normalize()
        # find the event's trading day: the first trading day on/after the event date
        pos = trading_days.searchsorted(ev)
        if pos == 0 or pos >= len(trading_days):
            continue                      # event outside our price range
        if pos - WINDOW < 0 or pos + WINDOW >= len(trading_days):
            continue                      # not enough surrounding days
        window = ar.iloc[pos - WINDOW: pos + WINDOW + 1]
        rows.append(pd.Series(window.values,
                              index=range(-WINDOW, WINDOW + 1),
                              name=ev.date()))
    if not rows:
        print(f"  ! no usable events for {label}")
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    print(f"  {label:14s}: {len(df):3d} events analysed")
    return df


def significance(study: pd.DataFrame) -> dict:
    """Is the average abnormal move real, or just noise?

    We compute, PER EVENT, the Cumulative Abnormal Return (CAR) over the whole
    window and over the post-event days (0..+5). Then a one-sample t-test asks:
    "is the average CAR across all events reliably different from zero?"

        t = mean(CAR) / ( std(CAR) / sqrt(n) )

    |t| > ~2.0  ->  significant at ~5% (unlikely to be chance).
    No scipy needed — the t-statistic is a one-line formula.
    """
    n = len(study)
    car_full = study.sum(axis=1)                       # CAR over -5..+5 per event
    car_post = study.loc[:, 0:].sum(axis=1)            # CAR over 0..+5 per event

    def t_stat(x: pd.Series) -> float:
        sd = x.std(ddof=1)
        return float("nan") if sd == 0 else x.mean() / (sd / (n ** 0.5))

    return {
        "n": n,
        "car_full_mean": car_full.mean() * 100,
        "car_full_t": t_stat(car_full),
        "car_post_mean": car_post.mean() * 100,
        "car_post_t": t_stat(car_post),
    }


def ascii_car_chart(mean_ar: pd.Series, title: str) -> None:
    """Draw the average CUMULATIVE abnormal return as a horizontal ASCII bar chart.

    CAR = running sum of the average abnormal return -> the typical price path.
    """
    car = mean_ar.cumsum() * 100          # to percent
    lo, hi = min(car.min(), 0), max(car.max(), 0)
    span = (hi - lo) or 1.0
    width = 50
    zero_col = int(round((0 - lo) / span * width))

    print(f"\n  {title}")
    print(f"  (cumulative abnormal return %, day 0 = event)\n")
    for day, val in car.items():
        col = int(round((val - lo) / span * width))
        if col >= zero_col:
            bar = " " * zero_col + "#" * (col - zero_col)
        else:
            bar = " " * col + "#" * (zero_col - col)
        marker = "  <-- EVENT DAY" if day == 0 else ""
        print(f"   t{day:+d} |{bar:<{width}}| {val:+6.2f}%{marker}")
    print(f"        {' ' * zero_col}^ zero line")


# ---------------------------------------------------------------------------
# 3. MAIN
# ---------------------------------------------------------------------------
def main() -> None:
    if not UNIFIED.exists():
        raise SystemExit("Run clean_events.py first to build processed/unified_events.csv")

    events = pd.read_csv(UNIFIED, parse_dates=["date"])
    events = events[events["date"] >= EST_START]

    print("Downloading prices from Yahoo Finance ...")
    stock = fetch_prices(STOCK_TICKER)
    market = fetch_prices(MARKET_TICKER)
    print(f"  {STOCK_TICKER}: {len(stock)} days ({stock.index.min().date()} -> {stock.index.max().date()})")
    print(f"  {MARKET_TICKER}: {len(market)} days\n")

    # daily simple returns, aligned on common trading days
    r_stock = stock.pct_change()
    r_market = market.pct_change()
    ar = (r_stock - r_market).dropna()            # market-adjusted abnormal return

    # --- define the event groups we care about -----------------------------
    groups = {
        "Dividends": events[events["event_type"] == "DIVIDEND"]["date"],
        "Earnings":  events[events["event_type"].str.startswith("RESULT_", na=False)]["date"],
        "Acquisitions": events[events["event_type"] == "ACQUISITION"]["date"],
    }

    print("Running event studies (window = +/-%d trading days):" % WINDOW)
    summary = {}
    for label, dates in groups.items():
        study = run_event_study(dates.dropna().drop_duplicates(), ar, label)
        if study.empty:
            continue
        study.to_csv(OUT_DIR / f"event_study_{label.lower()}.csv")
        summary[label] = (study.mean(), significance(study))   # (avg AR path, stats)

    # --- report ------------------------------------------------------------
    for label, (mean_ar, stats) in summary.items():
        ascii_car_chart(mean_ar, f"{label}  (RELIANCE vs Nifty 50)")

    # --- significance table ------------------------------------------------
    print("\n  STATISTICAL SIGNIFICANCE (is the pattern real, or noise?)")
    print("  " + "-" * 68)
    print(f"  {'Event':14s} {'n':>4s} {'CAR[-5,+5]':>11s} {'t':>7s}  "
          f"{'CAR[0,+5]':>10s} {'t':>7s}  verdict")
    print("  " + "-" * 68)
    for label, (_, s) in summary.items():
        sig = "SIGNIFICANT" if abs(s["car_full_t"]) > 2 else "not sig. (noise)"
        print(f"  {label:14s} {s['n']:4d} {s['car_full_mean']:+10.2f}% "
              f"{s['car_full_t']:+7.2f}  {s['car_post_mean']:+9.2f}% "
              f"{s['car_post_t']:+7.2f}  {sig}")
    print("  " + "-" * 68)
    print("  Rule of thumb: |t| > 2.0 ~ 95% confidence the average move isn't zero.")

    print("\n  READING THE CHART:")
    print("   * Bars right of the zero line = stock beat the market around the event.")
    print("   * A run-up BEFORE t0 hints at anticipation / information leakage.")
    print("   * Drift AFTER t0 hints the market kept re-pricing on the news.")
    print(f"\n  Per-event detail saved to: {OUT_DIR}\\event_study_*.csv")


if __name__ == "__main__":
    main()
