"""
enrich_board_meetings.py
===============================================================================
Backfill the thin board-meetings feed (NSE only serves ~2 recent years) with the
historical "Outcome of Board Meeting" records that live in the announcements feed
(which goes back to ~2006). Real board-meeting events — the outcome-announcement
date is when the news actually hits — so this is a legitimate, ADDITIVE deepening
of the BOARD_MEETING sample.

For each stock:
  • pull announcement rows whose desc is a board-meeting category
  • map them into the board_meetings schema (bm_date = announcement date)
  • append to the existing board_meetings.csv, de-dupe by date (keep the richest
    row, so original structured intimations win over synthesised ones), sort newest-first

RUN
    python enrich_board_meetings.py --dry     # report only
    python enrich_board_meetings.py           # write
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
DRY = "--dry" in sys.argv

# announcement desc values that denote a board meeting
BM_DESC = {"Outcome of Board Meeting", "Board Meeting Intimation",
           "Board Meeting", "Board meeting postponed"}


def to_bmdate(series: pd.Series) -> pd.Series:
    """Parse an announcement date (ISO-ish) and render as dd-Mon-YYYY to match
    the existing bm_date format."""
    dt = pd.to_datetime(series, errors="coerce")          # ISO -> no dayfirst
    return dt.dt.strftime("%d-%b-%Y"), dt


def main() -> None:
    folders = sorted(p for p in ROOT.iterdir()
                     if p.is_dir() and (p / "announcements.csv").exists())
    print(f"Backfilling board meetings for {len(folders)} stocks"
          + ("   [DRY RUN]" if DRY else "") + "\n")
    tot_before = tot_after = 0
    changed = []
    for f in folders:
        sym = f.name
        ann = pd.read_csv(f / "announcements.csv", dtype=str).fillna("")
        m = ann[ann["desc"].str.strip().isin(BM_DESC)].copy()
        if m.empty:
            continue
        datecol = "an_dt" if "an_dt" in m.columns else "sort_date"
        bm_str, bm_dt = to_bmdate(m[datecol])
        add = pd.DataFrame({
            "bm_symbol": sym, "symbol": sym,
            "bm_date": bm_str,
            "bm_purpose": m["desc"].values,
            "bm_desc": m.get("attchmntText", m["desc"]).values,
            "sm_name": m.get("sm_name", ""), "sm_isin": m.get("sm_isin", ""),
            "company_name": m.get("company_name", ""),
            "bm_source": "announcements",
        })
        add = add[bm_dt.notna().values]                   # drop unparseable dates

        bmp = f / "board_meetings.csv"
        old = pd.read_csv(bmp, dtype=str).fillna("") if bmp.exists() else pd.DataFrame()
        before = pd.to_datetime(old.get("bm_date"), errors="coerce",
                                dayfirst=True).dropna().nunique() if len(old) else 0

        cols = list(dict.fromkeys(list(old.columns) + list(add.columns)))
        big = pd.concat([old.reindex(columns=cols), add.reindex(columns=cols)],
                        ignore_index=True).fillna("")
        big["_rich"] = (big != "").sum(axis=1)
        big["_s"] = pd.to_datetime(big["bm_date"], errors="coerce", dayfirst=True)
        big = (big.sort_values("_rich", ascending=False)
                  .drop_duplicates(subset="bm_date", keep="first")
                  .sort_values("_s", ascending=False, na_position="last")
                  .drop(columns=["_rich", "_s"]))
        after = pd.to_datetime(big["bm_date"], errors="coerce",
                               dayfirst=True).dropna().nunique()
        tot_before += before; tot_after += after
        if after > before:
            changed.append((sym, before, after))
            if not DRY:
                big.to_csv(bmp, index=False)

    changed.sort(key=lambda r: r[2] - r[1], reverse=True)
    print(f"{'stock':<12}{'before':>7}{'after':>7}{'+':>6}")
    for sym, b, a in changed[:15]:
        print(f"{sym:<12}{b:>7}{a:>7}{a-b:>+6}")
    if len(changed) > 15:
        print(f"   ... and {len(changed)-15} more")
    print(f"\nStocks enriched: {len(changed)} | board-meeting dates {tot_before} -> {tot_after} (+{tot_after-tot_before})")
    print("Dry run — nothing written." if DRY else "Written.")


if __name__ == "__main__":
    main()
