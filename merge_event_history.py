"""
merge_event_history.py
===============================================================================
Build the FULLEST possible event history per stock by MERGING two sources:

  • what's on disk now  (the fresh NSE export imported from `new data/`)
  • the original CSVs    (recovered from the last git commit, HEAD)

For every <SYMBOL>/<feed>.csv we take the UNION of rows, de-duplicate on a
natural key, sort newest-first and write it back. This recovers history that
either source was missing — most importantly the deep financial-results history
that the recent NSE export had truncated — so the behaviour win-rates are based
on the maximum number of past events (better accuracy).

Purely additive in spirit: it only ever ADDS rows back (union), never drops a
genuine event. The old versions remain in git history regardless.

RUN
    python merge_event_history.py --dry     # report row/date deltas, write nothing
    python merge_event_history.py           # write the merged CSVs
===============================================================================
"""
from __future__ import annotations

import io
import subprocess
import sys
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parent
DRY = "--dry" in sys.argv

# feed -> (date column for sorting, natural-key columns for de-dup)
FEEDS = {
    "announcements":     ("sort_date",     ["seq_id", "sort_date", "desc"]),
    # one board meeting per date — key on bm_date ONLY, else the same meeting
    # described slightly differently by two sources gets kept twice.
    "board_meetings":    ("bm_date",       ["bm_date"]),
    "corporate_actions": ("exDate",        ["exDate", "subject", "series", "recDate"]),
    "financial_results": ("broadCastDate", ["seqNumber", "broadCastDate", "period", "toDate"]),
}


def git_head(relpath: str) -> pd.DataFrame | None:
    r = subprocess.run(["git", "show", f"HEAD:{relpath}"], capture_output=True)
    if r.returncode != 0 or not r.stdout.strip():
        return None
    try:
        return pd.read_csv(io.BytesIO(r.stdout), dtype=str).fillna("")
    except Exception:                                # noqa: BLE001
        return None


def read_disk(path: Path) -> pd.DataFrame | None:
    if not path.exists():
        return None
    try:
        return pd.read_csv(path, dtype=str).fillna("")
    except Exception:                                # noqa: BLE001
        return None


def ndates(df: pd.DataFrame | None, col: str) -> int:
    if df is None or col not in df.columns:
        return 0
    return pd.to_datetime(df[col], errors="coerce", dayfirst=True).dropna().dt.normalize().nunique()


def merge_one(cur, old, datecol, keycols):
    frames = [f for f in (cur, old) if f is not None and len(f)]
    if not frames:
        return None
    cols = list(dict.fromkeys(c for f in frames for c in f.columns))   # union, ordered
    big = pd.concat([f.reindex(columns=cols) for f in frames], ignore_index=True).fillna("")
    keys = [k for k in keycols if k in big.columns] or cols
    big["_k"] = big[keys].astype(str).agg("".join, axis=1)
    big = big.drop_duplicates(subset="_k").drop(columns="_k")
    if datecol in big.columns:
        big["_s"] = pd.to_datetime(big[datecol], errors="coerce", dayfirst=True)
        big = big.sort_values("_s", ascending=False, na_position="last").drop(columns="_s")
    return big.reset_index(drop=True)


def main() -> None:
    folders = sorted(p.name for p in ROOT.iterdir()
                     if p.is_dir() and any((p / f"{k}.csv").exists() for k in FEEDS))
    print(f"Merging event history for {len(folders)} stock folders"
          + ("   [DRY RUN]" if DRY else "") + "\n")
    tot_added = 0
    header = f"{'stock':<12}{'feed':<18}{'disk':>6}{'git':>6}{'merged':>8}{'+dates':>8}"
    for sym in folders:
        printed = False
        for feed, (datecol, keycols) in FEEDS.items():
            rel = f"{sym}/{feed}.csv"
            path = ROOT / rel
            cur, old = read_disk(path), git_head(rel)
            if cur is None and old is None:
                continue
            merged = merge_one(cur, old, datecol, keycols)
            d_disk, d_git, d_merged = ndates(cur, datecol), ndates(old, datecol), ndates(merged, datecol)
            added = d_merged - d_disk
            tot_added += max(0, added)
            if added != 0:                            # only show feeds that changed
                if not printed:
                    print(header); printed = True
                print(f"{sym:<12}{feed:<18}{d_disk:>6}{d_git:>6}{d_merged:>8}{added:>+8}")
            if not DRY and merged is not None:
                merged.to_csv(path, index=False)
        if printed:
            print()
    print(f"{'Total unique event-dates recovered across all feeds:':<52}{tot_added:>+8}")
    print("\nDry run — nothing written." if DRY else "\nMerged CSVs written.")


if __name__ == "__main__":
    main()
