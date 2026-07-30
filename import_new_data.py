"""
import_new_data.py
===============================================================================
Refresh the per-stock event feeds from the official NSE export dropped in
`new data/` (the 32 *.xlsx disclosure files). Splits the four big sheets by
symbol and writes each stock's <SYMBOL>/<feed>.csv, for the current Nifty 50
only (read from 08_Company_Directory.xlsx).

SAFE / ADDITIVE:
  • Only writes a feed CSV for a symbol that actually HAS rows in the export.
    A symbol missing from a sheet (e.g. M&M, whose '&' the exporter dropped) is
    left exactly as it was on disk — never overwritten with blanks.
  • Creates folders for the four new members (ETERNAL, INDIGO, MAXHEALTH, TMPV).
  • Ex-members' folders are untouched; they simply leave the Nifty 50 list.

RUN
    python import_new_data.py            # write the CSVs
    python import_new_data.py --dry      # report only, write nothing
===============================================================================
"""
from __future__ import annotations

import sys
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "new data"
DRY = "--dry" in sys.argv

# feed file  ->  (source xlsx, symbol column in that sheet, output csv name)
FEEDS = [
    ("02_Announcements.xlsx",       "symbol",    "announcements.csv"),
    ("04_Board_Meetings (1).xlsx",  "symbol",    "board_meetings.csv"),
    ("05_Corporate_Actions.xlsx",   "symbol",    "corporate_actions.csv"),
    ("13_Financial_Results.xlsx",   "symbol",    "financial_results.csv"),
]


def nifty50() -> list[str]:
    df = pd.read_excel(SRC / "08_Company_Directory.xlsx", dtype=str)
    return sorted(df["symbol"].dropna().str.strip().unique())


def main() -> None:
    if not SRC.exists():
        raise SystemExit(f"Source folder not found: {SRC}")
    universe = set(nifty50())
    print(f"Nifty 50 from Company Directory: {len(universe)} symbols"
          + ("   [DRY RUN — nothing written]" if DRY else ""))

    grand = {}
    for xlsx, symcol, outname in FEEDS:
        path = SRC / xlsx
        if not path.exists():
            print(f"\n! missing source: {xlsx}")
            continue
        df = pd.read_excel(path, dtype=str).fillna("")
        if symcol not in df.columns:            # board meetings fallback
            symcol = "bm_symbol"
        df[symcol] = df[symcol].str.strip()
        wrote, skipped = [], []
        print(f"\n{outname}  <-  {xlsx}   ({len(df):,} rows)")
        for sym in sorted(universe):
            sub = df[df[symcol] == sym]
            if sub.empty:
                skipped.append(sym)
                continue
            folder = ROOT / sym
            if not DRY:
                folder.mkdir(exist_ok=True)
                sub.to_csv(folder / outname, index=False)
            wrote.append(sym)
            grand.setdefault(sym, set()).add(outname)
        print(f"   wrote {len(wrote)} stocks; "
              f"no rows for {len(skipped)}: {', '.join(skipped) if skipped else '—'}")

    # per-stock feed coverage summary for the four NEW members
    print("\nNew members — feeds now present:")
    for sym in ("ETERNAL", "INDIGO", "MAXHEALTH", "TMPV"):
        feeds = sorted(f.replace('.csv', '') for f in grand.get(sym, []))
        print(f"   {sym:<10} {', '.join(feeds) if feeds else '(none in export)'}")
    print("\nDone." if not DRY else "\nDry run complete.")


if __name__ == "__main__":
    main()
