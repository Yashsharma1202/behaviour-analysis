"""
make_bench_sample.py
===============================================================================
Draw a reproducible, category-stratified sample of REAL announcement rows for
hand-labelling, and report the row-level gate statistics.

The benchmark in test_sentiment_finbert.py is the ship gate for Phase 1: FinBERT
only replaces the lexicon if it beats it on rows that actually came off NSE.
This script produces the rows to label; it never assigns labels itself.

    python make_bench_sample.py            # writes processed/_bench_sample.csv
    python make_bench_sample.py --n 200
===============================================================================
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

import sentiment_finbert as SF

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "processed" / "_bench_sample.csv"
SEED = 42


def load_all() -> pd.DataFrame:
    frames = []
    for p in sorted(ROOT.glob("*/announcements.csv")):
        try:
            df = pd.read_csv(p, dtype=str).fillna("")
        except Exception:                         # noqa: BLE001
            continue
        if df.empty or "desc" not in df.columns:
            continue
        if "attchmntText" not in df.columns:
            df["attchmntText"] = ""
        df["symbol"] = p.parent.name
        frames.append(df[["symbol", "desc", "attchmntText"]])
    if not frames:
        raise SystemExit("no announcement feeds found")
    a = pd.concat(frames, ignore_index=True)
    a["desc"] = a["desc"].str.strip()
    a["attchmntText"] = a["attchmntText"].str.strip()
    return a


def gate_stats(a: pd.DataFrame) -> dict:
    """Row-level gate rate — the number that describes what a USER sees in the
    Announcements tab, as opposed to the unique-text rate score_feeds reports."""
    gated = [SF.is_boilerplate(c, t) for c, t in zip(a["desc"], a["attchmntText"])]
    n = len(a)
    return {"rows": n, "gated": int(sum(gated)),
            "pct": round(sum(gated) / n * 100, 1) if n else 0.0}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=200, help="rows to sample")
    ap.add_argument("--per-cat", type=int, default=6)
    args = ap.parse_args()

    a = load_all()
    st = gate_stats(a)
    print(f"corpus : {st['rows']:,} announcement rows")
    print(f"gate   : {st['gated']:,} rows ({st['pct']}%) are procedural boilerplate")

    a["text"] = (a["desc"] + " | " + a["attchmntText"]).str.slice(0, 400)
    a = a.drop_duplicates(subset=["text"])
    print(f"unique : {len(a):,} distinct texts")

    # Stratify by category so the benchmark is not 40% "Loss of Share
    # Certificates". groupby.apply is avoided deliberately — pandas 3.0 removed
    # include_groups and the loop is clearer besides.
    vc = a["desc"].value_counts()
    cats = [c for c in vc.index if vc[c] >= 8]
    picks = [g.sample(min(args.per_cat, len(g)), random_state=SEED)
             for c in cats for g in [a[a["desc"] == c]]]
    samp = pd.concat(picks, ignore_index=True)
    samp = samp.sample(min(args.n, len(samp)), random_state=SEED).reset_index(drop=True)

    samp["label"] = ""                            # <- to be filled in by hand
    samp[["symbol", "desc", "attchmntText", "label"]].to_csv(OUT, index=False)
    print(f"\nsampled {len(samp)} rows across {samp['desc'].nunique()} categories")
    print(f"-> {OUT}   (fill the `label` column with Positive/Negative/Neutral)")


if __name__ == "__main__":
    main()
