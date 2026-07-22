"""
earnings_surprise.py
===============================================================================
POST-EARNINGS DRIFT — do quarterly BEATS drift up and MISSES drift down?

This is the one thing annual data could not do: measure the earnings SURPRISE and
test whether the market keeps drifting in the surprise's direction after a result
(the "PEAD" anomaly). Uses the quarterly/ folder.

HARD LIMITS — read them, they bound every conclusion below:
  * quarterly data is only Mar-2023 → Mar-2026 (13 quarters), ONE bull-market
    regime. ~9-13 result events per stock, ~700 pooled.
  * there are NO analyst estimates, so "surprise" cannot mean "vs consensus".
    We use the estimate-free textbook proxy: YoY growth and its ACCELERATION vs
    the company's own trailing trend (a seasonal random walk).
  * 3 years / one regime -> NO honest train/test is possible. This is
    DESCRIPTIVE. It says what happened 2023-2026, not what will happen next.

Honesty rails kept from the rest of the project:
  * entry at the CLOSE on the result day (T+0, the next close-able bar after the
    after-hours broadcast) -> every trade is real.
  * returns MARKET-ADJUSTED (NIFTY removed) and net of 0.10% cost.
  * t is MONTH-CLUSTERED (pooled stocks share market days).
  * a long-beat / short-miss spread is shown: that is market-neutral, so it
    cannot be explained by "the market went up in 2023-2026".

    python earnings_surprise.py

Output: processed/earnings_surprise.csv + printed report
===============================================================================
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import event_behaviour as EB
import nifty_backtest as NBT

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                  # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
COST = 0.001
MAXAFT = 40                    # look this many trading days past the result
_MON = {"Mar": 3, "Jun": 6, "Sep": 9, "Dec": 12}


def _num(x):
    if not isinstance(x, str):
        return np.nan
    x = x.strip().replace(",", "").replace("%", "")
    try:
        return float(x)
    except ValueError:
        return np.nan


def quarterly(sym: str) -> pd.DataFrame | None:
    """-> tidy per-quarter table with the metrics we bucket on."""
    f = ROOT / "quarterly" / f"{sym}.csv"
    if not f.exists():
        return None
    df = pd.read_csv(f, dtype=str).fillna("")
    df["m"] = df["Metric"].str.replace("└─", "", regex=False).str.strip()
    qcols = [c for c in df.columns if re.match(r"(Mar|Jun|Sep|Dec)\s*\d{4}", c)]
    if len(qcols) < 6:
        return None

    def series(name):
        r = df[df["m"] == name]
        return [_num(v) for v in r.iloc[0][qcols]] if len(r) else [np.nan] * len(qcols)

    def qend(c):
        mo, yr = c.split()[0], int(re.search(r"\d{4}", c).group())
        return pd.Timestamp(yr, _MON[mo], 1) + pd.offsets.MonthEnd(0)

    out = pd.DataFrame({
        "qend": [qend(c) for c in qcols],
        "sales": series("Sales"),
        "npat": series("Net Profit"),
        "eps": series("EPS in Rs"),
        "yoy_sales": series("YOY Sales Growth %"),
        "yoy_profit": series("YOY Profit Growth %"),
        "opm": series("OPM %"),
    }).sort_values("qend").reset_index(drop=True)
    # SURPRISE proxy = acceleration: this quarter's YoY profit growth minus the
    # trailing 4-quarter average. Beating your own trend is the estimate-free
    # stand-in for beating expectations.
    out["accel"] = out["yoy_profit"] - out["yoy_profit"].shift(1).rolling(4).mean()
    return out


def result_dates(sym: str) -> dict:
    """quarter-end -> the date that quarter's result was broadcast (first close-able)."""
    f = ROOT / sym / "financial_results.csv"
    if not f.exists():
        return {}
    fr = pd.read_csv(f, dtype=str).fillna("")
    if not {"period", "toDate", "broadCastDate"} <= set(fr.columns):
        return {}
    fr = fr[fr["period"] == "Quarterly"]
    td = pd.to_datetime(fr["toDate"], errors="coerce", dayfirst=True)
    bd = pd.to_datetime(fr["broadCastDate"], errors="coerce", format="mixed", dayfirst=True)
    out = {}
    for t, b in zip(td, bd):
        if pd.isna(t) or pd.isna(b):
            continue
        qe = (t + pd.offsets.MonthEnd(0)).normalize()
        out[qe] = min(out.get(qe, b), b)          # earliest broadcast for that quarter
    return out


def main():
    syms = [s for s in EB.discover_feed_symbols()
            if (ROOT / "quarterly" / f"{s}.csv").exists()]
    mkt_ix = NBT.fetch_index()

    print("POST-EARNINGS DRIFT  —  quarterly beats vs misses, pooled across stocks")
    print("quarterly data Mar-2023 → Mar-2026 · ONE bull regime · DESCRIPTIVE ONLY")
    print(f"entry at the close on result day · {COST*100:.2f}% cost · market-adjusted\n")

    rows = []
    for sym in syms:
        q = quarterly(sym)
        rd = result_dates(sym)
        px = EB.fetch_prices_yahoo(sym)
        if q is None or not rd or px is None:
            continue
        mk = mkt_ix.reindex(px.index).ffill().bfill()
        p, m, I = px.values.astype(float), mk.values.astype(float), px.index
        for _, r in q.iterrows():
            bc = rd.get(r["qend"])
            if bc is None:
                continue
            # broadcasts are after-hours -> first trade is the NEXT close
            pos = I.searchsorted(bc.normalize() + pd.Timedelta(days=1))
            if pos <= 0 or pos + MAXAFT >= len(I):
                continue
            base = p[pos]
            path = {f"d{k}": (p[pos + k] / base - 1 - (m[pos + k] / m[pos] - 1)) * 100
                    for k in range(0, MAXAFT + 1)}
            rows.append({"symbol": sym, "qend": r["qend"].date(), "buy_date": I[pos].date(),
                         "yoy_profit": r["yoy_profit"], "yoy_sales": r["yoy_sales"],
                         "accel": r["accel"], **path})
    if not rows:
        raise SystemExit("no matched quarters")
    df = pd.DataFrame(rows)
    df["mo"] = pd.to_datetime(df["buy_date"]).dt.to_period("M").astype(str)
    df.to_csv(PROC / "earnings_surprise.csv", index=False)
    print(f"   matched {len(df)} quarterly results to prices "
          f"({df['buy_date'].min()} → {df['buy_date'].max()})\n")

    def tmonth(sub, col):
        s = sub.dropna(subset=[col])
        if len(s) < 5:
            return float("nan")
        mm = s.groupby("mo")[col].mean()
        if len(mm) < 3 or mm.std(ddof=1) == 0:
            return float("nan")
        return mm.mean() / (mm.std(ddof=1) / len(mm) ** 0.5)

    for dim, lbl in [("accel", "SURPRISE  (YoY profit growth vs the company's own trailing trend)"),
                     ("yoy_profit", "RAW  YoY Profit Growth %"),
                     ("yoy_sales", "RAW  YoY Sales Growth %")]:
        d = df.dropna(subset=[dim]).copy()
        d["b"] = d.groupby("mo")[dim].transform(
            lambda s: pd.qcut(s.rank(method="first"), 3, labels=["MISS", "MID", "BEAT"])
            if s.notna().sum() >= 6 else np.nan)
        print("=" * 100)
        print(f"  BUCKET BY {lbl}")
        print("=" * 100)
        print(f"  {'bucket':<7} {'n':>4}  " + "".join(f"{'T+'+str(k):>7}" for k in (1, 3, 5, 10, 20, 30, 40)))
        for b in ("BEAT", "MID", "MISS"):
            s = d[d["b"] == b]
            if s.empty:
                continue
            cells = "".join(f"{s['d'+str(k)].mean():>+7.2f}" for k in (1, 3, 5, 10, 20, 30, 40))
            print(f"  {b:<7} {len(s):>4}  {cells}")
        be, mi = d[d["b"] == "BEAT"], d[d["b"] == "MISS"]
        print(f"\n  LONG-BEAT / SHORT-MISS spread (market-neutral):")
        print(f"    {'':13}" + "".join(f"{'T+'+str(k):>9}" for k in (5, 10, 20, 30, 40)))
        srow, trow = "    spread   ", "    t(month) "
        for k in (5, 10, 20, 30, 40):
            col = f"d{k}"
            sp = be[col].mean() - mi[col].mean()
            tmp = pd.concat([be.assign(_x=be[col]), mi.assign(_x=-mi[col])])
            tv = tmonth(tmp, "_x")
            srow += f"{sp:>+8.2f}%"
            trow += f"{tv:>+9.2f}"
        print(srow); print(trow)
        print()

    print("=" * 100)
    print("  READ")
    print("   BEAT/MID/MISS = top / middle / bottom third of the surprise, ranked")
    print("   against the OTHER stocks reporting the same month. Numbers are the")
    print("   average market-adjusted % move from the result-day close to T+k.")
    print("   The LONG-BEAT/SHORT-MISS spread is market-neutral: a positive, growing")
    print("   spread with t(month) > 2 is genuine post-earnings drift. If the spread")
    print("   is flat or noisy, beats and misses behaved the same — no signal.")
    print("   CAVEAT: 3 years, one bull regime, no out-of-sample. Descriptive only.")
    print(f"\n  Saved -> {PROC}\\earnings_surprise.csv")


if __name__ == "__main__":
    main()
