"""
test_sentiment_finbert.py
===============================================================================
THE PHASE-1 SHIP GATE.

FinBERT only earns its place if it beats the lexicon that is already wired into
the dashboard, on rows that actually came off NSE. This scores four engines on
the same 200 hand-labelled filings:

    always-Neutral   the majority-class stub — the floor any engine must clear
    lexicon          sentiment_analyzer.analyze_sentiment (what ships today)
    finbert-raw      the bare model, no gate, no overrides, no preamble strip
    router           sentiment_finbert.score_many — the thing being proposed

`finbert-raw` is in here on purpose: without it a good router score cannot be
attributed. If raw FinBERT already beats the lexicon, the gate and overrides are
decoration; if it does not and the router does, the routing is what earned it.

METRIC: MACRO-F1, not accuracy. The corpus is 80% Neutral, so accuracy flatters
anything that says Neutral a lot — the stub scores 80% while being useless.

    python test_sentiment_finbert.py
    python test_sentiment_finbert.py --show-errors     # every row the router got wrong

Exit code 0 if the router beats the lexicon on macro-F1, 1 otherwise.
===============================================================================
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

import sentiment_finbert as SF
from sentiment_analyzer import analyze_sentiment

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
LABELS = ROOT / "processed" / "sentiment_bench_labels.csv"
CLASSES = ("Positive", "Neutral", "Negative")


# --------------------------------------------------------------------------- #
# Metrics — written out rather than pulled from sklearn so the definition of
# every number printed below is visible in this file.
# --------------------------------------------------------------------------- #
def prf(gold: list[str], pred: list[str], cls: str) -> tuple[float, float, float, int]:
    tp = sum(g == cls and p == cls for g, p in zip(gold, pred))
    fp = sum(g != cls and p == cls for g, p in zip(gold, pred))
    fn = sum(g == cls and p != cls for g, p in zip(gold, pred))
    prec = tp / (tp + fp) if tp + fp else 0.0
    rec = tp / (tp + fn) if tp + fn else 0.0
    f1 = 2 * prec * rec / (prec + rec) if prec + rec else 0.0
    return prec, rec, f1, tp + fn


def evaluate(gold: list[str], pred: list[str]) -> dict:
    per = {c: prf(gold, pred, c) for c in CLASSES}
    return {
        "acc": sum(g == p for g, p in zip(gold, pred)) / len(gold),
        "macro_f1": sum(per[c][2] for c in CLASSES) / len(CLASSES),
        "per_class": per,
    }


def confusion(gold: list[str], pred: list[str]) -> pd.DataFrame:
    m = pd.DataFrame(0, index=list(CLASSES), columns=list(CLASSES), dtype=int)
    for g, p in zip(gold, pred):
        if g in m.index and p in m.columns:
            m.loc[g, p] += 1
    return m


# --------------------------------------------------------------------------- #
# The engines under test
# --------------------------------------------------------------------------- #
def predict_stub(texts, cats) -> list[str]:
    return ["Neutral"] * len(texts)


def predict_lexicon(texts, cats) -> list[str]:
    return [analyze_sentiment(f"{c} {t}")[0] for c, t in zip(cats, texts)]


def predict_finbert_raw(texts, cats) -> list[str]:
    """The bare model: no boilerplate gate, no preamble strip, no overrides."""
    if not SF._load_model():
        return ["Neutral"] * len(texts)
    return [lab for _sc, _conf, lab in
            SF._finbert_batch([f"{c}. {t}" for c, t in zip(cats, texts)])]


def predict_router(texts, cats) -> list[str]:
    return [r.label for r in SF.score_many(texts, cats)]


ENGINES = [
    ("always-Neutral", predict_stub, "majority-class floor"),
    ("lexicon", predict_lexicon, "what ships today"),
    ("finbert-raw", predict_finbert_raw, "bare model, no routing"),
    ("router", predict_router, "gate + FinBERT + overrides"),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--show-errors", action="store_true")
    args = ap.parse_args()

    if not LABELS.exists():
        raise SystemExit(f"missing {LABELS}\n"
                         f"run: python make_bench_sample.py && python make_bench_labels.py")
    df = pd.read_csv(LABELS, dtype=str).fillna("")
    gold = df["label"].tolist()
    cats = df["desc"].tolist()
    texts = df["attchmntText"].tolist()

    info = SF.model_info()
    print("=" * 78)
    print("  PHASE-1 SHIP GATE — sentiment engine benchmark")
    print(f"  {len(df)} hand-labelled NSE announcement rows")
    print(f"  model: {info['model']}  device: {info.get('device') or 'n/a'}"
          f"{'' if info['available'] else '  (UNAVAILABLE)'}")
    dist = pd.Series(gold).value_counts()
    print("  gold: " + "  ".join(f"{c}={int(dist.get(c, 0))}" for c in CLASSES))
    print("=" * 78)

    results = {}
    for name, fn, note in ENGINES:
        pred = fn(texts, cats)
        results[name] = (evaluate(gold, pred), pred, note)

    print(f"\n  {'engine':<16}{'macro-F1':>10}{'accuracy':>10}   note")
    print("  " + "-" * 74)
    for name, _fn, _note in ENGINES:
        r, _pred, note = results[name]
        print(f"  {name:<16}{r['macro_f1']:>10.3f}{r['acc']:>10.1%}   {note}")

    print("\n  per-class F1")
    print(f"  {'engine':<16}" + "".join(f"{c:>12}" for c in CLASSES))
    print("  " + "-" * 74)
    for name, _fn, _note in ENGINES:
        r = results[name][0]
        print(f"  {name:<16}" + "".join(f"{r['per_class'][c][2]:>12.3f}" for c in CLASSES))

    print("\n  router confusion matrix (rows = gold, cols = predicted)")
    cm = confusion(gold, results["router"][1])
    print("    " + cm.to_string().replace("\n", "\n    "))

    lex = results["lexicon"][0]["macro_f1"]
    raw = results["finbert-raw"][0]["macro_f1"]
    rtr = results["router"][0]["macro_f1"]

    print("\n  attribution")
    print(f"    raw model vs lexicon : {raw - lex:+.3f} macro-F1")
    print(f"    routing adds         : {rtr - raw:+.3f} macro-F1")
    print(f"    router vs lexicon    : {rtr - lex:+.3f} macro-F1")

    if args.show_errors:
        print("\n  router errors")
        pred = results["router"][1]
        for i, (g, p) in enumerate(zip(gold, pred)):
            if g != p:
                t = f"{cats[i]} :: {texts[i]}".replace("\n", " ")[:110]
                print(f"    [{i:>3}] gold={g:<9} pred={p:<9} {t}")

    print("\n" + "=" * 78)
    if rtr > lex:
        print(f"  GATE PASSED — router {rtr:.3f} > lexicon {lex:.3f} macro-F1")
        print("=" * 78)
        return 0
    print(f"  GATE FAILED — router {rtr:.3f} <= lexicon {lex:.3f} macro-F1")
    print("  Per the plan: keep the lexicon and say so. Do not ship FinBERT.")
    print("=" * 78)
    return 1


if __name__ == "__main__":
    sys.exit(main())
