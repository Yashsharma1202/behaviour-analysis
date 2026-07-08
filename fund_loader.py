"""
fund_loader.py
===============================================================================
Parses the Screener.in-style fundamental CSVs (balance_sheet / cash_flow / pnl /
quarterly / ratios) into clean, tidy time-series tables -- for ANY of the ~3,400
stocks in those folders, not just Reliance.

THE RAW FORMAT
--------------
Each file is "wide": first column = Metric, remaining columns = periods.
Metric rows use tree characters to show hierarchy, e.g.
        Net Profit,23640,29861,...
          └─ Profit Growth %,8%,34%,...
Values may be plain numbers (374372) or percentages (-13.64%).

WHAT THIS DOES
--------------
  parse_screener(path) -> tidy DataFrame  (index = period Timestamp,
                                           columns = clean metric names, numeric)
  load_stock(symbol)   -> dict of the 5 statements for one stock

Run directly to dump Reliance's tidy tables + a quick fundamentals snapshot:
        python fund_loader.py            # RELIANCE
        python fund_loader.py TCS        # any symbol
===============================================================================
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
PROC.mkdir(exist_ok=True)
STATEMENTS = ["pnl", "quarterly", "balance_sheet", "cash_flow", "ratios"]


def _clean_metric(name: str) -> str:
    """Strip the tree characters / indentation from a metric label."""
    name = name.replace("└─", " ").replace("├─", " ")
    name = re.sub(r"\s+", " ", name).strip()
    return name


def _to_number(v):
    """'374372' -> 374372.0 ; '-13.64%' -> -13.64 ; '' -> NaN."""
    if v is None:
        return float("nan")
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "nan", "None"):
        return float("nan")
    s = s.replace("%", "")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def parse_screener(path: Path) -> pd.DataFrame:
    """Read one Screener-style CSV -> tidy DataFrame (rows=period, cols=metric).

    Input : path to e.g. quarterly/RELIANCE.csv
    Output: DataFrame indexed by period (Timestamp), one column per metric.
    """
    raw = pd.read_csv(path, dtype=str)
    raw = raw.rename(columns={raw.columns[0]: "Metric"})
    raw["Metric"] = raw["Metric"].map(_clean_metric)
    raw = raw[raw["Metric"] != ""]
    raw = raw.drop_duplicates(subset="Metric", keep="first")   # first wins on dupes

    period_cols = [c for c in raw.columns if c != "Metric" and c != "Raw PDF"]
    tidy = raw.set_index("Metric")[period_cols].T          # periods become rows
    tidy = tidy.apply(lambda col: col.map(_to_number))     # numeric everywhere
    tidy.index = pd.to_datetime(tidy.index, format="%b %Y", errors="coerce")
    tidy = tidy[tidy.index.notna()].sort_index()
    tidy.index.name = "period"
    return tidy


def load_stock(symbol: str) -> dict[str, pd.DataFrame]:
    """Return {statement: tidy DataFrame} for one symbol (missing files skipped)."""
    out = {}
    for st in STATEMENTS:
        p = ROOT / st / f"{symbol}.csv"
        if p.exists():
            try:
                out[st] = parse_screener(p)
            except Exception as e:                          # noqa: BLE001
                print(f"  ! {st}/{symbol}.csv failed: {e}")
    return out


def main():
    symbol = sys.argv[1] if len(sys.argv) > 1 else "RELIANCE"
    data = load_stock(symbol)
    if not data:
        raise SystemExit(f"No fundamental files found for {symbol}")

    print(f"Loaded {symbol}: " + ", ".join(f"{k}({len(v)}p)" for k, v in data.items()))

    # save tidy tables
    for st, df in data.items():
        df.to_csv(PROC / f"{symbol}_{st}.csv")

    # --- quick snapshot from the annual P&L -------------------------------
    if "pnl" in data:
        p = data["pnl"]
        print(f"\n{symbol} — annual P&L snapshot (Rs crore)")
        print(f"  {'year':>8} {'Sales':>10} {'NetProfit':>10} {'OPM%':>6} {'EPS':>8}")
        for dt, row in p.tail(6).iterrows():
            print(f"  {dt.strftime('%Y'):>8} {row.get('Sales', float('nan')):>10,.0f} "
                  f"{row.get('Net Profit', float('nan')):>10,.0f} "
                  f"{row.get('OPM %', float('nan')):>5.0f}% "
                  f"{row.get('EPS in Rs', float('nan')):>8.2f}")

    # --- quarterly YoY profit growth (the earnings-surprise fuel) ----------
    if "quarterly" in data:
        q = data["quarterly"]
        col = next((c for c in q.columns if "YOY Profit Growth" in c), None)
        npc = next((c for c in q.columns if c.strip() == "Net Profit"), None)
        print(f"\n{symbol} — quarterly net profit & YoY growth")
        print(f"  {'quarter':>9} {'NetProfit':>10} {'YoY%':>8}")
        for dt, row in q.tail(8).iterrows():
            g = row.get(col, float("nan")) if col else float("nan")
            print(f"  {dt.strftime('%b %Y'):>9} {row.get(npc, float('nan')):>10,.0f} "
                  f"{g:>+7.0f}%")

    print(f"\nSaved tidy tables -> {PROC}\\{symbol}_*.csv")


if __name__ == "__main__":
    main()
