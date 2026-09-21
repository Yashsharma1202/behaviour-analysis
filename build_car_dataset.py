"""
build_car_dataset.py
===============================================================================
Build the training set for finetune_sentiment.py: every NSE filing paired with
what the stock actually did afterwards.

THE LABEL
---------
Not my opinion about tone — the STANDARDISED abnormal return:

    sar = car_h / (sigma * sqrt(h))

car_h and sigma both come from disclosure_timing.py: market model removed,
entry at the first tradeable close strictly after the filing instant, sigma the
stock's own estimation-period residual sd. Dividing by sigma*sqrt(h) makes a
2% move in a sleepy utility and a 2% move in a small-cap comparable, which
matters when one model has to read both.

Classes are TERCILES of sar, computed on the TRAINING ERA ONLY and then applied
to later eras. Computing them over the whole sample would leak the future
distribution into the split — a subtle leak that would flatter the test score.

WHAT IS EXCLUDED
    * boilerplate filings (sentiment_finbert.is_boilerplate) — procedural rows
      carry no information and are 17% of the corpus. Training on them teaches
      the model that most text means nothing.
    * rows with no matching announcement text
    * rows where car/sigma is missing or sigma is ~0

    python build_car_dataset.py                  # horizon 20, default
    python build_car_dataset.py --horizon 5
Output: processed/car_dataset.parquet
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
OUT = ROOT / "processed" / "car_dataset.parquet"


def build(horizon: int) -> pd.DataFrame:
    car = pd.read_csv(CAR_FILE,
                      usecols=["symbol", "ts", f"car_{horizon}", "sigma"])
    car["ts"] = pd.to_datetime(car["ts"], errors="coerce")
    car = car.dropna(subset=["ts", f"car_{horizon}", "sigma"])
    car = car[car["sigma"] > 1e-6]

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

        m = sub.merge(a[["_ts", "desc", "attchmntText"]],
                      left_on="ts", right_on="_ts")
        if m.empty:
            continue
        for _i, r in m.iterrows():
            d, t = str(r["desc"]), str(r["attchmntText"])
            if SF.is_boilerplate(d, t):
                continue
            txt = SF.strip_preamble(f"{d}. {t}".strip())
            if len(txt) < 20:
                continue
            car_v = float(r[f"car_{horizon}"])
            sig = float(r["sigma"])
            rows.append({
                "symbol": sym, "ts": r["ts"], "text": txt, "category": d,
                "car": car_v, "sigma": sig,
                "sar": car_v / 100.0 / (sig * np.sqrt(horizon)),
            })
    df = pd.DataFrame(rows).sort_values("ts").reset_index(drop=True)
    # A handful of filings sit on top of enormous moves that have nothing to do
    # with the filing (index rebalances, circuit breakers). Winsorise so the
    # loss is not dominated by a dozen rows.
    lo, hi = df["sar"].quantile([0.005, 0.995])
    df["sar"] = df["sar"].clip(lo, hi)
    return df


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--horizon", type=int, default=20, choices=[1, 5, 10, 20])
    args = ap.parse_args()

    print("=" * 78)
    print(f"  CAR DATASET  —  horizon +{args.horizon}d, standardised by sigma*sqrt(h)")
    print("=" * 78)
    df = build(args.horizon)
    df["horizon"] = args.horizon

    OUT.parent.mkdir(parents=True, exist_ok=True)
    df.to_parquet(OUT, index=False)

    print(f"  {len(df):,} filings with text + outcome, "
          f"{df['symbol'].nunique()} symbols")
    print(f"  {df['ts'].min().date()} .. {df['ts'].max().date()}")
    print(f"  sar: mean {df['sar'].mean():+.3f}  sd {df['sar'].std():.3f}  "
          f"median {df['sar'].median():+.3f}")
    print("\n  filings per year")
    yr = df.groupby(df["ts"].dt.year).size()
    for y, n in yr.items():
        print(f"    {y}  {n:>6,}  {'#' * int(n / max(yr.max(), 1) * 46)}")
    print(f"\n  -> {OUT}")


if __name__ == "__main__":
    main()
