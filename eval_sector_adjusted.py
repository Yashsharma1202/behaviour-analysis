"""
eval_sector_adjusted.py
===============================================================================
HOW MUCH OF THE -1.14% CAR SPREAD WAS ACTUALLY SECTOR DRIFT?

eval_sentiment_car.py reported that filings reading bullish were followed by
-1.14% abnormal return over 20 days (p=0.0007), robust to sub-period and to
dropping any single symbol. Stage 3 then found that every sector-neutral
configuration of the backtest was NEGATIVE while every market-adjusted one was
positive — which points at the market model.

disclosure_timing.py removes MARKET beta (alpha + beta * NIFTY). It does not
remove INDUSTRY exposure. So a filing published by a company whose whole sector
was about to mean-revert gets credited with an "abnormal" return that was
really a sector return. Leave-one-symbol-out cannot catch that: dropping one
stock leaves the sector effect intact.

This measures it directly. Same calls, same CAR, one difference:

    raw               spread of car
    sector-adjusted   spread of (car - mean car of the same industry x month)

If the spread survives demeaning, the signal is about the filing. If it
collapses, the signal was about the sector and the headline finding needs
retracting.

    python eval_sector_adjusted.py
    python eval_sector_adjusted.py --perms 5000
===============================================================================
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
CALLS = ROOT / "processed" / "grid_calls.parquet"


def spread(d: pd.DataFrame, col: str) -> tuple[float, int, int]:
    p = d.loc[d["tag"] == "Positive", col]
    n = d.loc[d["tag"] == "Negative", col]
    if len(p) < 20 or len(n) < 20:
        return float("nan"), len(p), len(n)
    return float(p.mean() - n.mean()), len(p), len(n)


def perm_p(d: pd.DataFrame, col: str, obs: float, perms: int, rng) -> float:
    """Shuffle tags within symbol; two-sided."""
    if not np.isfinite(obs):
        return float("nan")
    groups = [(g[col].to_numpy(), g["tag"].to_numpy())
              for _s, g in d.groupby("symbol", sort=False)]
    hits = 0
    for _ in range(perms):
        pv, nv = [], []
        for v, t in groups:
            sh = rng.permutation(t)
            pv.append(v[sh == "Positive"]); nv.append(v[sh == "Negative"])
        pv, nv = np.concatenate(pv), np.concatenate(nv)
        if len(pv) < 20 or len(nv) < 20:
            continue
        if abs(pv.mean() - nv.mean()) >= abs(obs):
            hits += 1
    return (hits + 1) / (perms + 1)


def row(label: str, d: pd.DataFrame, col: str, perms: int, rng):
    s, npos, nneg = spread(d, col)
    if not np.isfinite(s):
        print(f"  {label:<26}  too few ({npos}/{nneg})")
        return None
    p = perm_p(d, col, s, perms, rng)
    star = "***" if p < 0.01 else ("**" if p < 0.05 else ("*" if p < 0.10 else ""))
    print(f"  {label:<26}{s:>+10.3f}%{p:>10.4f} {star:<4}"
          f"{d.loc[d.tag == 'Positive', col].mean():>+10.3f}"
          f"{d.loc[d.tag == 'Negative', col].mean():>+10.3f}{npos:>9,}{nneg:>9,}")
    return s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--perms", type=int, default=2000)
    args = ap.parse_args()
    rng = np.random.default_rng(42)

    if not CALLS.exists():
        raise SystemExit(f"missing {CALLS} — run backtest_grid.py first")
    d = pd.read_parquet(CALLS)
    d = d[d["industry"].notna() & (d["industry"] != "")]

    # Industry x calendar-month mean, removed. Same construction Stage 3 used.
    d["car_sect"] = d["car"] - d.groupby(["industry", "month"])["car"].transform("mean")

    print("=" * 100)
    print("  IS THE SENTIMENT-CAR SPREAD A FILING EFFECT OR A SECTOR EFFECT?")
    print(f"  {len(d):,} directional calls · {d['symbol'].nunique():,} symbols · "
          f"{d['industry'].nunique()} industries · CAR +20d")
    print("=" * 100)
    print(f"  {'':<26}{'spread':>10}{'perm p':>10} {'':<4}"
          f"{'mean +':>10}{'mean -':>10}{'n pos':>9}{'n neg':>9}")
    print("  " + "-" * 96)
    raw = row("raw (market-adjusted)", d, "car", args.perms, rng)
    adj = row("sector-adjusted", d, "car_sect", args.perms, rng)

    if raw is not None and adj is not None and raw != 0:
        kept = adj / raw * 100
        print(f"\n  {kept:.0f}% of the raw spread survives industry-month demeaning.")
        if abs(adj) < abs(raw) * 0.4:
            print("  The signal was largely SECTOR DRIFT, not filing information.")
            print("  The market model removes market beta but not industry exposure,")
            print("  and that gap is where most of the headline result was living.")
        elif abs(adj) < abs(raw) * 0.75:
            print("  A substantial part was sector drift; some filing-specific")
            print("  signal remains, but the headline figure overstated it.")
        else:
            print("  The spread is mostly filing-specific — sector drift does not")
            print("  explain it, and the Stage-3 sector result needs another cause.")

    print("\n  by era (sector-adjusted)")
    print(f"  {'era':<26}{'spread':>10}{'n':>10}")
    print("  " + "-" * 46)
    for lo, hi, name in ((2008, 2015, "2008-2015"), (2016, 2021, "2016-2021"),
                         (2022, 2026, "2022-2026")):
        sub = d[(d.year >= lo) & (d.year <= hi)]
        s, npos, nneg = spread(sub, "car_sect")
        print(f"  {name:<26}{s:>+10.3f}%{npos + nneg:>10,}")

    print("\n  Note: industry comes from smIndustry in each announcements.csv, and")
    print("  is current classification applied historically — a company that")
    print("  changed sector is mislabelled in its early years.")


if __name__ == "__main__":
    main()
