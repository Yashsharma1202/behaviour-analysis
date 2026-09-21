"""
eval_sentiment_car.py
===============================================================================
DOES THE SENTIMENT TAG PREDICT ANYTHING?

test_sentiment_finbert.py scores the engines against 200 rows I hand-labelled.
That measures agreement with my opinion about tone. It does not measure whether
a tag is worth showing to anyone.

This measures the thing that matters: after a filing tagged Positive, does the
stock actually beat the market? Ground truth is the CUMULATIVE ABNORMAL RETURN
already computed by disclosure_timing.py — 88,548 filings across 50 symbols,
market model removed, no labelling opinion involved.

WHAT IS REPORTED
    spread   mean CAR after Positive filings minus mean CAR after Negative ones.
             This is the number that decides whether the tag is worth its pixels.
    perm p   labels shuffled WITHIN each symbol, spread recomputed. Answers
             "how often would a tag this useless look this good by luck?"
             Shuffling within symbol keeps each stock's own return distribution
             intact, so a stock that simply drifted up cannot manufacture a
             spread.
    by stage rule / numeric / finbert / override measured separately — which
             stage of the router is actually earning its place.

A spread that is large but insignificant is noise. A spread that is significant
but tiny is real and useless. Both have to clear.

    python eval_sentiment_car.py
    python eval_sentiment_car.py --horizon 1 --perms 5000
===============================================================================
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import score_feeds
from sentiment_analyzer import analyze_sentiment

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
CAR_FILE = ROOT / "processed" / "disclosure_timing_announcements.csv"


def build_panel(horizon: int) -> pd.DataFrame:
    """Join every filing's CAR to the tag each engine gave it."""
    car = pd.read_csv(CAR_FILE, usecols=["symbol", "ts", "desc", f"car_{horizon}"])
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
        m = sub.merge(a[["_ts", "desc", "attchmntText"]], left_on="ts",
                      right_on="_ts", suffixes=("", "_a"))
        if m.empty:
            continue
        for _i, r in m.iterrows():
            d, t = r["desc_a"], r["attchmntText"]
            hit = cache.get(score_feeds.text_key(d, t))
            if not hit:
                continue
            rows.append({
                "symbol": sym, "ts": r["ts"], "car": float(r[f"car_{horizon}"]),
                "router": hit["label"], "engine": hit.get("engine", "?"),
                "lexicon": analyze_sentiment(f"{d} {t}")[0],
            })
    return pd.DataFrame(rows)


def spread(df: pd.DataFrame, col: str) -> tuple[float, int, int]:
    p = df.loc[df[col] == "Positive", "car"]
    n = df.loc[df[col] == "Negative", "car"]
    if len(p) < 20 or len(n) < 20:
        return float("nan"), len(p), len(n)
    return float(p.mean() - n.mean()), len(p), len(n)


def perm_p(df: pd.DataFrame, col: str, obs: float, perms: int, rng) -> float:
    """Shuffle labels within each symbol; how often does luck beat the observed
    spread? Two-sided, so a confidently WRONG tag is not rewarded."""
    if not np.isfinite(obs):
        return float("nan")
    groups = [g for _s, g in df.groupby("symbol", sort=False)]
    cars = [g["car"].to_numpy() for g in groups]
    labs = [g[col].to_numpy() for g in groups]
    hits = 0
    for _ in range(perms):
        pv, nv = [], []
        for c, l in zip(cars, labs):
            sh = rng.permutation(l)
            pv.append(c[sh == "Positive"])
            nv.append(c[sh == "Negative"])
        pv, nv = np.concatenate(pv), np.concatenate(nv)
        if len(pv) < 20 or len(nv) < 20:
            continue
        if abs(pv.mean() - nv.mean()) >= abs(obs):
            hits += 1
    return (hits + 1) / (perms + 1)


def report(df: pd.DataFrame, col: str, label: str, perms: int, rng):
    s, np_, nn = spread(df, col)
    if not np.isfinite(s):
        print(f"  {label:<22}   too few tagged rows (pos={np_}, neg={nn})")
        return
    p = perm_p(df, col, s, perms, rng)
    star = "***" if p < 0.01 else ("**" if p < 0.05 else ("*" if p < 0.10 else ""))
    pos = df.loc[df[col] == "Positive", "car"].mean()
    neg = df.loc[df[col] == "Negative", "car"].mean()
    print(f"  {label:<22}{s:>+9.3f}%{p:>9.4f} {star:<4}"
          f"{pos:>+9.3f}{neg:>+9.3f}{np_:>9,}{nn:>9,}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--horizon", type=int, default=5, choices=[1, 5, 10, 20])
    ap.add_argument("--perms", type=int, default=2000)
    args = ap.parse_args()
    rng = np.random.default_rng(42)

    print("=" * 96)
    print(f"  DOES THE TAG PREDICT THE MOVE?  ground truth = CAR over {args.horizon} "
          f"trading day(s), market model removed")
    print("=" * 96)

    df = build_panel(args.horizon)
    if df.empty:
        raise SystemExit("no rows joined — run score_feeds.py and disclosure_timing.py first")
    print(f"  {len(df):,} filings joined across {df['symbol'].nunique()} symbols "
          f"({df['ts'].min().date()} to {df['ts'].max().date()})\n")

    hdr = (f"  {'engine':<22}{'spread':>10}{'perm p':>9} {'':<4}"
           f"{'mean +':>9}{'mean -':>9}{'n pos':>9}{'n neg':>9}")
    print(hdr)
    print("  " + "-" * 92)
    report(df, "lexicon", "lexicon", args.perms, rng)
    report(df, "router", "router (all stages)", args.perms, rng)

    print("\n  which STAGE of the router earns its place")
    print("  " + "-" * 92)
    for eng, name in (("numeric", "numeric parser"), ("finbert", "FinBERT"),
                      ("override", "rule overrides")):
        sub = df[df["engine"] == eng]
        if len(sub) < 100:
            print(f"  {name:<22}   only {len(sub):,} rows — not measurable")
            continue
        report(sub, "router", name, args.perms, rng)

    print("\n  * p<0.10   ** p<0.05   *** p<0.01   (two-sided, labels shuffled "
          "within symbol)")
    print("  Spread is in percentage points of abnormal return. A tag is worth "
          "showing only if")
    print("  the spread is both statistically real AND large enough to matter "
          "after costs (~0.10%).")


if __name__ == "__main__":
    main()
