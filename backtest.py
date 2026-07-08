"""
backtest.py
===============================================================================
Turns the event BEHAVIOUR we measured into actual TRADING STRATEGIES and asks
the only question that matters: would this have made money?

We test three earnings-timing strategies on real Reliance prices, each trade
priced NET OF COSTS, and report proper performance stats + a buy-&-hold baseline.

THE STRATEGIES
--------------
  A  Pre-earnings run-up (LONG)   buy 5 trading days before results, sell the day
                                  before. Bets on the "+0.6% drift into results".
  B  Post-earnings fade (SHORT)   short from results day, cover 5 days later.
                                  Bets on the "-1% sell-the-news" drift.
  C  PEAD conditional             on results day go LONG if profit beat YoY, SHORT
                                  if it missed; hold 5 days. Needs fundamentals
                                  (only 5 events -> illustrative, not conclusive).

THE METRICS (defined in plain English)
--------------------------------------
  Total return   compounding every trade's P&L (start with Rs 1, what do you end with).
  Win rate       fraction of trades that made money.
  Avg / trade    mean profit per trade (%).
  t-stat         mean / (std/sqrt(n)) -> is the average trade reliably > 0? |t|>2 ~ 95%.
  Sharpe         risk-adjusted return, annualised: (mean/std) x sqrt(trades per year).
                 Higher = more return per unit of wobble. >1 is good, >2 is excellent.
  Max drawdown   worst peak-to-trough drop of the equity curve (the pain you'd endure).

FAIR-COMPARISON NOTE
--------------------
These strategies are only IN the market a few days per quarter (mostly in cash),
so their TOTAL return can't be compared head-to-head with buy-&-hold (always
invested). Compare on Sharpe (risk-adjusted) and return-per-day-exposed instead.

    python backtest.py

Needs internet (Yahoo). Uses processed/unified_events.csv + earnings_surprise.csv.
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

# Windows consoles default to cp1252 and choke on box-drawing chars -> force UTF-8.
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                              # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"

TICKER = "RELIANCE.NS"
COST = 0.001          # 0.10% round-trip transaction cost per trade (brokerage+slippage)

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
    s = pd.Series(adj, index=pd.to_datetime(res["timestamp"], unit="s").normalize())
    return s.dropna().sort_index()


# ---------------------------------------------------------------------------
def run_strategy(px: pd.Series, events, entry_off: int, exit_off: int,
                 direction) -> pd.DataFrame:
    """Build the trade list for one strategy.

    px         : adjusted-close price series (index = trading dates).
    events     : list of (event_date, dir_hint) ; dir_hint used when direction='signal'.
    entry_off  : trading-day offset from the event to ENTER (e.g. -5).
    exit_off   : trading-day offset from the event to EXIT  (e.g. -1).
    direction  : +1 (always long), -1 (always short), or 'signal' (per-event).
    Returns a DataFrame of trades with entry/exit dates and net return.
    """
    idx = px.index
    trades = []
    for ev, hint in events:
        pos = idx.searchsorted(pd.Timestamp(ev).normalize())
        i_in, i_out = pos + entry_off, pos + exit_off
        if i_in < 0 or i_out >= len(idx) or i_in >= i_out:
            continue
        p_in, p_out = px.iloc[i_in], px.iloc[i_out]
        d = hint if direction == "signal" else direction
        gross = (p_out / p_in - 1) * d          # short flips the sign
        net = gross - COST                      # subtract round-trip cost
        trades.append({"event": pd.Timestamp(ev).date(),
                       "entry": idx[i_in].date(), "exit": idx[i_out].date(),
                       "dir": "LONG" if d > 0 else "SHORT",
                       "ret": net})
    return pd.DataFrame(trades)


def metrics(trades: pd.DataFrame, years: float) -> dict:
    """Compute performance stats from a trade list."""
    if trades.empty:
        return {}
    r = trades["ret"]
    n = len(r)
    equity = (1 + r).cumprod()                  # Rs 1 growing trade by trade
    peak = equity.cummax()
    max_dd = (equity / peak - 1).min()
    sd = r.std(ddof=1)
    t = r.mean() / (sd / n ** 0.5) if sd else float("nan")
    trades_per_year = n / years if years else n
    sharpe = (r.mean() / sd * (trades_per_year ** 0.5)) if sd else float("nan")
    wins = r[r > 0]
    losses = r[r <= 0]
    return {
        "n": n,
        "total_%": (equity.iloc[-1] - 1) * 100,
        "avg_%": r.mean() * 100,
        "win_rate": len(wins) / n * 100,
        "avg_win_%": wins.mean() * 100 if len(wins) else 0.0,
        "avg_loss_%": losses.mean() * 100 if len(losses) else 0.0,
        "t": t,
        "sharpe": sharpe,
        "max_dd_%": max_dd * 100,
    }


def spark(equity) -> str:
    """One-line ASCII equity curve (cp1252-safe ramp)."""
    ramp = ".:-=+*#@"
    lo, hi = equity.min(), equity.max()
    rng = (hi - lo) or 1
    return "".join(ramp[min(7, int((v - lo) / rng * 7))] for v in equity)


# ---------------------------------------------------------------------------
def main() -> None:
    uni = PROC / "unified_events.csv"
    if not uni.exists():
        raise SystemExit("Run clean_events.py first.")
    ev = pd.read_csv(uni, parse_dates=["date"])
    earn_dates = (ev[ev["event_type"].str.startswith("RESULT_", na=False)]["date"]
                  .dt.normalize().drop_duplicates().sort_values())

    print("Downloading prices ...")
    px = fetch_adj(TICKER)
    span_years = (px.index.max() - px.index.min()).days / 365.25
    print(f"  {TICKER}: {len(px)} days, {span_years:.1f} yrs\n")

    # event lists: (date, direction_hint).  hint only used by strategy C.
    all_earn = [(d, +1) for d in earn_dates]
    sur = PROC / "earnings_surprise.csv"
    pead = []
    if sur.exists():
        s = pd.read_csv(sur, parse_dates=["event_date"])
        pead = [(r["event_date"], +1 if r["label"] == "BEAT" else -1)
                for _, r in s.iterrows()]

    strategies = [
        ("A pre-earnings run-up (long)",  all_earn, -5, -1, +1),
        ("B post-earnings fade (short)",  all_earn,  0, +5, -1),
        ("C PEAD conditional (long/short)", pead,    0, +5, "signal"),
    ]

    print(f"BACKTEST  (cost {COST*100:.2f}%/trade)  –  {len(earn_dates)} earnings dates")
    print("=" * 92)
    print(f"  {'strategy':32} {'n':>3} {'total':>8} {'avg/tr':>7} {'win%':>6} "
          f"{'t':>6} {'Sharpe':>7} {'maxDD':>7}  equity")
    print("-" * 92)
    saved = {}
    for name, events, ein, eout, d in strategies:
        tr = run_strategy(px, events, ein, eout, d)
        if tr.empty:
            print(f"  {name:32} (no trades)")
            continue
        m = metrics(tr, span_years)
        eq = (1 + tr["ret"]).cumprod()
        print(f"  {name:32} {m['n']:>3} {m['total_%']:>+7.1f}% {m['avg_%']:>+6.2f}% "
              f"{m['win_rate']:>5.0f}% {m['t']:>+6.2f} {m['sharpe']:>7.2f} "
              f"{m['max_dd_%']:>6.1f}%  {spark(eq)}")
        tr.to_csv(PROC / f"backtest_{name[0].lower()}.csv", index=False)
        saved[name] = {**m, "strategy": name}

    # buy & hold baseline over the same span
    bh_ret = px.iloc[-1] / px.iloc[0] - 1
    daily = px.pct_change().dropna()
    bh_sharpe = daily.mean() / daily.std() * (252 ** 0.5)

    # persist a summary the GUI can read
    summary = pd.DataFrame(saved.values())
    summary = pd.concat([summary, pd.DataFrame([{
        "strategy": "buy & hold", "n": 0, "total_%": bh_ret * 100,
        "sharpe": bh_sharpe, "avg_%": 0, "win_rate": 0, "t": 0, "max_dd_%": 0,
    }])], ignore_index=True)
    summary.to_csv(PROC / "backtest_summary.csv", index=False)
    print("-" * 92)
    print(f"  {'buy & hold (always invested)':32} {'':>3} {bh_ret*100:>+7.1f}% "
          f"{'':>7} {'':>6} {'':>6} {bh_sharpe:>7.2f}")
    print("=" * 92)

    print("\nHOW TO READ IT:")
    print("  * total% compounds every trade (strategies sit in cash between events, so")
    print("    a smaller total than buy&hold is expected – compare Sharpe & t instead).")
    print("  * |t| > 2 and Sharpe > 1 would suggest a real, tradeable edge.")
    print("  * Strategy C has only ~5 trades -> treat as illustration, not evidence.")
    print(f"\nSaved trade logs -> {PROC}\\backtest_*.csv")


if __name__ == "__main__":
    main()
