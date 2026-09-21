"""
make_bench_labels.py
===============================================================================
Attach hand-assigned labels to the stratified sample and write the frozen
benchmark file that test_sentiment_finbert.py scores against.

WHO LABELLED THESE
------------------
Claude (the assistant) labelled all 200 rows by reading the actual filing text,
not the user. They are a considered first pass, not gospel — see LIMITS below.
Re-label any row you disagree with: edit processed/sentiment_bench_labels.csv
directly, it is the file the test reads.

LABELLING CONVENTION
--------------------
  Positive  materially good for the shareholder — profit/sales grew against the
            stated prior period, order win, bonus, buyback, dividend
            recommendation, commercial production start, value-accretive deal
  Negative  materially bad — profit fell against the stated prior period, a
            loss, auditor/director resignation, strike, trading suspension, an
            exchange clarification demand, a restatement of results
  Neutral   procedural or directionless — AGM/EGM notices, book closure, RTA
            changes, ESOP allotments, routine appointments, statutory
            disclosures, results filings quoting NO prior-period comparison

For "Results Update" rows the label follows NET PROFIT against whatever prior
period the filing itself quotes. Where the filing quotes no comparison, the row
is Neutral — the filing genuinely carries no direction.

LIMITS — read before trusting any number that comes out of this
---------------------------------------------------------------
  * The Neutral class is ~80% of the sample. That is a true reflection of NSE
    filings, but it means ACCURACY IS A USELESS METRIC here: a stub that always
    answers "Neutral" scores ~80%. The test reports macro-F1 and prints the
    majority-class baseline for exactly this reason.
  * The Neutral/mild-Negative boundary is a judgement call (is a routine auditor
    rotation a red flag?). That ambiguity hits all three engines equally, so the
    COMPARISON stands even where an individual label is arguable.
  * 200 rows across 160 categories is thin per category. Treat the result as a
    go/no-go gate, not as a precise accuracy estimate.

    python make_bench_labels.py
===============================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
SAMPLE = ROOT / "processed" / "_bench_sample.csv"
OUT = ROOT / "processed" / "sentiment_bench_labels.csv"

# index -> label, for every row that is NOT Neutral. Everything unlisted is
# Neutral, which keeps this table readable and reviewable.
POSITIVE = {
    3,    # Sun Pharma FY09: PAT 1,265cr vs 1,014cr prior
    8,    # Infosys FY10: PAT 6,266cr vs 5,988cr
    13,   # Sun Pharma: board to consider dividend
    17,   # ITC FY13: PAT 7,418cr vs 6,162cr
    19,   # TCS Q2FY10: PAT 1,348cr vs 1,276cr
    25,   # ITC FY09 consol: PAT 3,325cr vs 3,158cr
    27,   # buyback public notice
    46,   # Infosys FY08: PAT 4,659cr vs 3,856cr
    47,   # Maruti: final dividend recommended
    55,   # Asian Paints: dividend recommended
    58,   # L&T: board to consider dividend declaration
    60,   # Coal India Q3FY13: PAT 4,395cr vs 3,078cr QoQ and 4,038cr YoY
    62,   # Bajaj Auto H1FY10: PAT 696cr vs 360cr
    66,   # UltraTech FY11: PAT 1,404cr vs 1,093cr
    80,   # Adani: incorporation of a JV
    83,   # L&T secures ultra-mega 6,400MW contract from Adani Power
    89,   # JSW Q3FY10: PAT 430cr vs a 188cr LOSS — turnaround
    90,   # Adani Ent H1FY09: PAT 207cr vs 116cr
    104,  # POWERGRID acquires Bhadla-III & Bikaner-III transmission SPVs
    116,  # M&M signs Power Delivery Agreement
    125,  # commencement of commercial production/operations
    126,  # RIL media release: "Robust Operating Performance; Net Profit Crosses..."
    139,  # Trent: bonus issue
    170,  # Axis Bank: stock split approved
}

NEGATIVE = {
    0,    # Cipla FY11: PAT 990cr vs 1,083cr — fell despite sales growth
    6,    # resignation of independent director
    26,   # Bharti Q4FY10: PAT 2,024cr vs 2,237cr
    30,   # UltraTech Q1FY14: PAT 673cr vs 778cr, sales also down
    35,   # suspension of trading in Shriram Transport NCDs
    38,   # Tata Tea Q3FY09: PAT 396cr vs 1,674cr — down 76%
    56,   # Dr Reddy's Q4FY09: PAT 156cr vs 162cr
    57,   # demise of an independent director
    64,   # HCL Tech Q2FY11: PAT 195cr vs 238cr
    65,   # exchange sought clarification on results under Reg 33
    73,   # Dr Reddy's FY09 consol: LOSS of 917cr vs 438cr profit
    76,   # exchange sought clarification re "Suspension of..."
    133,  # resignation of director
    138,  # resignation of Director/KMP/SMP
    149,  # SBI officers' federation serves two-day strike notice
    188,  # REVISED financial result — a restatement
}


def main():
    if not SAMPLE.exists():
        raise SystemExit(f"missing {SAMPLE} — run make_bench_sample.py first")
    df = pd.read_csv(SAMPLE, dtype=str).fillna("")

    overlap = POSITIVE & NEGATIVE
    if overlap:
        raise SystemExit(f"row(s) labelled both ways: {sorted(overlap)}")
    bad = [i for i in (POSITIVE | NEGATIVE) if i >= len(df)]
    if bad:
        raise SystemExit(f"label index out of range for a {len(df)}-row sample: {bad}")

    df["label"] = ["Positive" if i in POSITIVE else
                   "Negative" if i in NEGATIVE else "Neutral"
                   for i in range(len(df))]
    df[["symbol", "desc", "attchmntText", "label"]].to_csv(OUT, index=False)

    vc = df["label"].value_counts()
    n = len(df)
    print(f"wrote {OUT}  ({n} rows)")
    for lab in ("Positive", "Negative", "Neutral"):
        c = int(vc.get(lab, 0))
        print(f"  {lab:<9} {c:>4}  {c / n * 100:>5.1f}%")
    print(f"\n  majority-class baseline (always Neutral) = "
          f"{vc.max() / n * 100:.1f}% accuracy, macro-F1 {1 / 3 * (2 * (vc.max() / n) / (1 + vc.max() / n)):.3f}")
    print("  -> judge on macro-F1, not accuracy.")


if __name__ == "__main__":
    main()
