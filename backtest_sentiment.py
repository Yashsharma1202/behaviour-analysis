"""
backtest_sentiment.py
===============================================================================
TURN THE TAGS INTO CALLS AND COUNT THE MONEY.

eval_sentiment_car.py showed the tag separates outcomes. That is not the same
as "it pays". This asks the only question that matters: if you had actually
taken every call, what would you have made, after costs?

THE CALL — AND WHY IT IS INVERTED
---------------------------------
The measured relation in this corpus is a REVERSAL: filings that read bullish
are followed by weaker abnormal returns (-1.14% over 20 days, p=0.0007). So the
strategy FADES the filing tone:

    tag Negative  ->  LONG   (bad news is over-sold, it reverts up)
    tag Positive  ->  SHORT  (good news is already priced, it drifts down)

Both directions are reported so you can see the raw fact rather than take the
inversion on trust.

RAILS — the same ones the rest of this repo uses
------------------------------------------------
  * Return is MARKET-ADJUSTED (market model removed). For a long-short book
    that is the honest measure; it is also ~the only one, since event_stats.py
    showed ~90% of a raw move is market drift.
  * Entry is the first tradeable CLOSE strictly after the filing instant. An
    after-hours disclosure is never credited with a move nobody could act on.
  * COSTS of 0.20% round trip (0.10% a side) are deducted from every call —
    brokerage, STT, impact. The rest of the repo uses 0.10%; this is stricter
    because a filing-driven trade is not a patient one.
  * IN-SAMPLE (<=2023) and OUT-OF-SAMPLE (2024+) are reported separately. The
    OOS block is the only one that means anything — everything in this project
    was built looking at the earlier data.

WHAT THIS BACKTEST DOES NOT MODEL
---------------------------------
  * Overlapping positions. Calls cluster in earnings season, so the monthly
    equity curve below is an equal-weight cohort approximation, not a
    fully-invested book. Effective independent bets are FEWER than the call
    count, so the Sharpe here is optimistic.
  * Shorting. Retail cannot easily short Indian single stocks beyond intraday
    or the F&O list. The LONG-ONLY block is the practically investable one.
  * Liquidity, impact beyond the flat cost, borrow cost, and the fact that
    every stock here is a current Nifty-50 member (survivorship).

    python backtest_sentiment.py
    python backtest_sentiment.py --horizon 5 --cost 0.30
===============================================================================
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import score_feeds
import sentiment_finbert as SF

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
CAR_FILE = ROOT / "processed" / "disclosure_timing_announcements.csv"
OOS_FROM = 2024


# --------------------------------------------------------------------------- #
def build_calls(horizon: int) -> pd.DataFrame:
    """Every filing the router tagged directionally, with its outcome."""
    car = pd.read_csv(CAR_FILE, usecols=["symbol", "ts", f"car_{horizon}"])
    car["ts"] = pd.to_datetime(car["ts"], errors="coerce")
    car = car.dropna(subset=["ts", f"car_{horizon}"])

    rows = []
    for sym, sub in car.groupby("symbol"):
        path = ROOT / sym / "announcements.csv"
        if not path.exists():
            continue
        try:
            a = pd.read_csv(path, dtype=str).fillna("")
        except Exception:                         # noqa: BLE001
            continue
        for c in ("desc", "attchmntText", "sort_date"):
            if c not in a.columns:
                a[c] = ""
        a["_ts"] = pd.to_datetime(a["sort_date"], errors="coerce", format="ISO8601")
        a = a.dropna(subset=["_ts"]).drop_duplicates(subset=["_ts"])
        cache = score_feeds.load_cache(sym)

        m = sub.merge(a[["_ts", "desc", "attchmntText"]], left_on="ts", right_on="_ts")
        for _i, r in m.iterrows():
            hit = cache.get(score_feeds.text_key(r["desc"], r["attchmntText"]))
            if not hit or hit["label"] == "Neutral":
                continue                          # no call, no trade
            rows.append({"symbol": sym, "ts": r["ts"],
                         "tag": hit["label"], "engine": hit.get("engine", "?"),
                         "car": float(r[f"car_{horizon}"]),
                         "category": str(r["desc"])[:40]})
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["year"] = df["ts"].dt.year
    df["month"] = df["ts"].dt.to_period("M")
    return df.sort_values("ts").reset_index(drop=True)


def apply_strategy(df: pd.DataFrame, cost: float, fade: bool = True) -> pd.DataFrame:
    """side = +1 long, -1 short. Return is market-adjusted, net of cost."""
    d = df.copy()
    if fade:                                      # fade the tone (the finding)
        d["side"] = np.where(d["tag"] == "Negative", 1, -1)
    else:                                         # take the tone at face value
        d["side"] = np.where(d["tag"] == "Positive", 1, -1)
    d["gross"] = d["side"] * d["car"]
    d["net"] = d["gross"] - cost
    d["win"] = d["net"] > 0
    return d


def stats(d: pd.DataFrame) -> dict:
    if d.empty:
        return {}
    n = len(d)
    net = d["net"].to_numpy()
    # t-stat clustered by month: filings bunch in earnings season, so treating
    # every call as an independent draw overstates significance badly.
    by_m = d.groupby("month")["net"].mean()
    t = (float(by_m.mean()) / (by_m.std(ddof=1) / np.sqrt(len(by_m)))
         if len(by_m) > 2 and by_m.std(ddof=1) > 0 else float("nan"))
    return {
        "n": n, "win": float(d["win"].mean() * 100),
        "mean": float(net.mean()), "median": float(np.median(net)),
        "sd": float(net.std(ddof=1)) if n > 1 else float("nan"),
        "best": float(net.max()), "worst": float(net.min()),
        "total": float(net.sum()), "months": int(len(by_m)),
        "t_month": t,
    }


def line(label: str, s: dict):
    if not s:
        print(f"  {label:<26}   no calls")
        return
    print(f"  {label:<26}{s['n']:>8,}{s['win']:>8.1f}%{s['mean']:>+9.3f}%"
          f"{s['median']:>+9.3f}%{s['sd']:>8.2f}{s['t_month']:>+8.2f}"
          f"{s['worst']:>+9.1f}%{s['best']:>+9.1f}%")


HDR = (f"  {'segment':<26}{'calls':>8}{'win%':>8}{'mean':>9}{'median':>9}"
       f"{'sd':>8}{'t(mth)':>8}{'worst':>9}{'best':>9}")


def equity(d: pd.DataFrame) -> pd.DataFrame:
    """Monthly equal-weight cohort return -> a curve you can look at.

    Every call opened in a month is one equal slice of that month's book. This
    is an APPROXIMATION: real positions overlap month boundaries and the book
    is not always fully invested.
    """
    m = d.groupby("month")["net"].agg(["mean", "count"]).rename(
        columns={"mean": "ret", "count": "n"})
    m["equity"] = (1 + m["ret"] / 100).cumprod()
    m["peak"] = m["equity"].cummax()
    m["dd"] = (m["equity"] / m["peak"] - 1) * 100
    return m


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--horizon", type=int, default=20, choices=[1, 5, 10, 20])
    ap.add_argument("--cost", type=float, default=0.20,
                    help="round-trip cost in %% deducted from every call")
    args = ap.parse_args()

    raw = build_calls(args.horizon)
    if raw.empty:
        raise SystemExit("no calls — run score_feeds.py and disclosure_timing.py first")

    print("=" * 104)
    print(f"  BACKTEST — fade the filing tone, hold {args.horizon} trading days, "
          f"{args.cost:.2f}% round-trip cost")
    print(f"  market-adjusted returns · entry = first tradeable close after the "
          f"filing · {raw['ts'].min().date()} to {raw['ts'].max().date()}")
    print("=" * 104)

    d = apply_strategy(raw, args.cost, fade=True)
    naive = apply_strategy(raw, args.cost, fade=False)

    print("\n  DIRECTION CHECK — is fading actually the right way round?")
    print(HDR); print("  " + "-" * 100)
    line("take tone at face value", stats(naive))
    line("FADE the tone", stats(d))

    print("\n  IN-SAMPLE vs OUT-OF-SAMPLE  (everything was built on <=2023)")
    print(HDR); print("  " + "-" * 100)
    line(f"in-sample  <={OOS_FROM - 1}", stats(d[d.year < OOS_FROM]))
    line(f"OUT-OF-SAMPLE {OOS_FROM}+", stats(d[d.year >= OOS_FROM]))

    print("\n  LONG-ONLY  (shorting Indian single stocks is not realistic for most)")
    print(HDR); print("  " + "-" * 100)
    lo = d[d["side"] == 1]
    line("long only (Negative tags)", stats(lo))
    line(f"  of which OOS {OOS_FROM}+", stats(lo[lo.year >= OOS_FROM]))

    print("\n  WHICH RULE PAYS")
    print(HDR); print("  " + "-" * 100)
    for eng in ("numeric", "finbert", "override"):
        line(eng, stats(d[d["engine"] == eng]))

    print("\n  BY YEAR")
    print(f"  {'year':<8}{'calls':>8}{'win%':>8}{'mean':>9}{'total':>10}")
    print("  " + "-" * 43)
    for y, g in d.groupby("year"):
        s = stats(g)
        print(f"  {y:<8}{s['n']:>8,}{s['win']:>8.1f}%{s['mean']:>+9.3f}%"
              f"{s['total']:>+9.1f}%")

    eq = equity(d)
    tot = (eq["equity"].iloc[-1] - 1) * 100
    yrs = (d["ts"].max() - d["ts"].min()).days / 365.25
    cagr = ((eq["equity"].iloc[-1]) ** (1 / yrs) - 1) * 100 if yrs > 0 else float("nan")
    sharpe = (eq["ret"].mean() / eq["ret"].std(ddof=1) * np.sqrt(12)
              if eq["ret"].std(ddof=1) > 0 else float("nan"))
    print(f"\n  EQUAL-WEIGHT MONTHLY COHORT CURVE  ({len(eq)} months)")
    print(f"    total return      {tot:+.1f}%")
    print(f"    CAGR              {cagr:+.2f}%   over {yrs:.1f} years")
    print(f"    monthly Sharpe    {sharpe:.2f}   (annualised, no risk-free)")
    print(f"    worst month       {eq['ret'].min():+.2f}%")
    print(f"    max drawdown      {eq['dd'].min():.1f}%")
    print(f"    positive months   {(eq['ret'] > 0).mean() * 100:.0f}%")

    print("\n  READ THIS BEFORE BELIEVING THE SHARPE")
    print("    * calls cluster in earnings season, so independent bets are far fewer")
    print("      than the call count and the Sharpe above is optimistic")
    print("    * the OOS block is the only one that was not looked at while building")
    print("    * shorting single stocks in India is impractical for most — judge the")
    print("      long-only block, not the long-short headline")
    print("    * every symbol is a CURRENT Nifty-50 member: survivorship bias")


if __name__ == "__main__":
    main()
