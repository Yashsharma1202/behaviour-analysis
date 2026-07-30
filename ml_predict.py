"""
ml_predict.py
===============================================================================
POOLED, REGULARISED ML for event-trade outcomes.
-------------------------------------------------------------------------------
The earlier ml_model.py trained PER STOCK (~48 rows) and failed. This version
fixes the real problem: it POOLS every scheduled event across all 50 stocks into
one dataset (thousands of rows), engineers as-of-entry-day features (no
look-ahead), and trains REGULARISED models with a strict TIME-BASED split, so the
score is honest out-of-time — not curve-fitted.

TASK
    For each scheduled event (Results / Board meeting / Corporate action), define
    a standard trade — buy N days before, sell M days after — and predict whether
    it finishes positive (WIN=1 / LOSS=0).

MODELS (both regularised)
    • Logistic Regression  — L2 penalty (C small = strong regularisation), scaled
    • HistGradientBoosting — L2 leaf reg + shallow trees + early stopping

VALIDATION
    Time split: train on the older events, test on the most recent — never random.
    Baseline to beat: "always predict WIN" (= the WIN rate in the test period).

    python ml_predict.py
===============================================================================
"""
from __future__ import annotations

import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")

import event_behaviour as EB
from download_feeds import NIFTY50_FALLBACK

from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import HistGradientBoostingClassifier
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import accuracy_score, roc_auc_score

# ── trade definition & knobs ────────────────────────────────────────────────
N_BEFORE = 5           # buy 5 trading days before the event
M_AFTER = 3            # sell 3 trading days after
EVENTS = ["RESULTS", "BOARD_MEETING", "CORPORATE_ACTION"]   # scheduled only
TEST_FROM = "2021-07-01"          # events on/after this date = out-of-time TEST set
MIN_HIST = 252                    # need 1y of prior prices for the 52w feature


def nifty() -> pd.Series | None:
    p = EB.PRICE_CACHE / "_NSEI.csv"
    if p.exists():
        return pd.read_csv(p, parse_dates=["date"]).set_index("date")["adj"].sort_index()
    return EB.fetch_prices_yahoo("^NSEI")


def features_for(px: pd.Series, nif: pd.Series, pos: int) -> dict | None:
    """As-of-entry-day features. Entry = close at pos-N_BEFORE; NOTHING after it."""
    i = pos - N_BEFORE
    if i - MIN_HIST < 0 or pos + M_AFTER >= len(px):
        return None
    w = px.iloc[:i + 1]                      # strictly up to & including entry day
    entry = float(w.iloc[-1])
    r = w.pct_change().dropna()
    mom20 = entry / float(w.iloc[-21]) - 1
    mom60 = entry / float(w.iloc[-61]) - 1
    vol20 = float(r.iloc[-20:].std())
    hi52 = float(w.iloc[-252:].max())
    ma50 = float(w.iloc[-50:].mean())
    # market relative strength on the entry date
    edate = w.index[-1]
    nw = nif.loc[:edate]
    nmom20 = (float(nw.iloc[-1]) / float(nw.iloc[-21]) - 1) if len(nw) > 21 else 0.0
    return {
        "mom20": mom20, "mom60": mom60, "vol20": vol20,
        "dist_52w_high": entry / hi52 - 1,
        "above_ma50": 1.0 if entry > ma50 else 0.0,
        "rel_strength": mom20 - nmom20,
        "month": edate.month,
    }


def build_dataset() -> pd.DataFrame:
    nif = nifty()
    if nif is None:
        raise SystemExit("no NIFTY price series")
    rows = []
    syms = sorted(set(NIFTY50_FALLBACK))
    for k, sym in enumerate(syms, 1):
        px = EB.fetch_prices_yahoo(sym)
        if px is None or len(px) < MIN_HIST + 30:
            continue
        ev = EB.load_events_from_feeds([sym])
        ev = ev[ev["event_type"].isin(EVENTS)]
        idx = px.index
        for _, e in ev.iterrows():
            pos = idx.searchsorted(pd.Timestamp(e["event_date"]))
            if pos <= 0 or pos >= len(idx):
                continue
            f = features_for(px, nif, pos)
            if f is None:
                continue
            entry = float(px.iloc[pos - N_BEFORE]); exit_ = float(px.iloc[pos + M_AFTER])
            f.update({"symbol": sym, "date": idx[pos],
                      "etype": e["event_type"],
                      "win": int((exit_ / entry - 1) > 0)})
            rows.append(f)
        if k % 10 == 0:
            print(f"  built {k}/{len(syms)} stocks…")
    df = pd.DataFrame(rows).dropna()
    # one-hot the event type
    df = pd.concat([df, pd.get_dummies(df["etype"], prefix="ev")], axis=1)
    return df


def main() -> None:
    print(f"Building pooled dataset — trade = buy {N_BEFORE}d before / sell {M_AFTER}d after, "
          f"events={EVENTS}\n")
    df = build_dataset()
    feat = ["mom20", "mom60", "vol20", "dist_52w_high", "above_ma50", "rel_strength", "month"] \
        + [c for c in df.columns if c.startswith("ev_")]

    tr = df[df["date"] < TEST_FROM]
    te = df[df["date"] >= TEST_FROM]
    print(f"\nDataset: {len(df):,} events  ({len(tr):,} train < {TEST_FROM}, {len(te):,} test ≥)")
    print(f"Overall WIN rate: {df['win'].mean()*100:.1f}%")

    Xtr, ytr = tr[feat].astype(float), tr["win"]
    Xte, yte = te[feat].astype(float), te["win"]

    base = yte.mean()                     # "always predict WIN"
    base_acc = max(base, 1 - base)
    print("\n" + "=" * 60)
    print(f"  BASELINE  (always predict majority class): {base_acc*100:.1f}%")
    print("=" * 60)

    # ── Regularised Logistic Regression (L2) ───────────────────────────────
    sc = StandardScaler().fit(Xtr)
    lr = LogisticRegression(C=0.1, penalty="l2", class_weight="balanced",
                            max_iter=2000).fit(sc.transform(Xtr), ytr)
    pl = lr.predict(sc.transform(Xte)); pp = lr.predict_proba(sc.transform(Xte))[:, 1]
    print(f"\n  Logistic Regression (L2, C=0.1) : acc {accuracy_score(yte,pl)*100:.1f}%  "
          f"AUC {roc_auc_score(yte,pp):.3f}")

    # ── Regularised Gradient Boosting ──────────────────────────────────────
    gb = HistGradientBoostingClassifier(
        max_depth=3, learning_rate=0.05, l2_regularization=1.0,
        max_iter=400, early_stopping=True, validation_fraction=0.15,
        class_weight="balanced", random_state=0).fit(Xtr, ytr)
    pg = gb.predict(Xte); pgp = gb.predict_proba(Xte)[:, 1]
    print(f"  HistGradientBoosting (regularised): acc {accuracy_score(yte,pg)*100:.1f}%  "
          f"AUC {roc_auc_score(yte,pgp):.3f}")

    # ── verdict ────────────────────────────────────────────────────────────
    best = max(accuracy_score(yte, pl), accuracy_score(yte, pg))
    print("\n" + "=" * 60)
    if best > base_acc + 0.01:
        print(f"  ML BEATS the baseline by {(best-base_acc)*100:.1f} pts — a real (if modest) edge.")
    else:
        print(f"  ML does NOT beat the baseline ({best*100:.1f}% vs {base_acc*100:.1f}%).")
        print("  Honest read: on these features, event outcomes are ~coin-flip beyond the")
        print("  base rate. More/better features (fundamentals, regime) are the next lever.")
    print("=" * 60)

    # feature signal (logistic coefficients, standardised)
    coef = pd.Series(lr.coef_[0], index=feat).sort_values(key=abs, ascending=False)
    print("\nStrongest signals (standardised logistic coefficients):")
    for k, v in coef.head(8).items():
        print(f"   {k:<16} {v:+.3f}  ({'higher → more likely WIN' if v>0 else 'higher → more likely LOSS'})")


if __name__ == "__main__":
    main()
