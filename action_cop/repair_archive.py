"""
repair_archive.py
===============================================================================
ONE-TIME repair of output_excels/ after the dedupe bug.

Until 2026-07-27, 35 of 39 save_to_excel() calls de-duplicated on ALL columns.
Feeds carrying a volatile field (difference / fileSize / sysTime) therefore
re-appended the same record on every run: 02_Announcements held 185,611 rows
for 78,476 real announcements — 58% phantom.

This script, per feed:
  * drops "No records found" placeholder rows
  * de-duplicates on the feed's natural key (DEDUPE_KEYS in scrap.py)
  * backfills a `fetched_at` stamp (file mtime) so pre-fix rows are dated
  * reports rows before -> after

SAFE BY DEFAULT — reports only. Nothing is written without --apply.

    python repair_archive.py            # report what would change
    python repair_archive.py --apply    # actually rewrite the files
===============================================================================
"""

import os
import sys
import shutil
from datetime import datetime

import pandas as pd

from scrap import OUTPUT_DIR, DEDUPE_KEYS, VOLATILE_COLS, _dedupe

APPLY = "--apply" in sys.argv


def natural_key(df, filename):
    key = [c for c in DEDUPE_KEYS.get(filename, []) if c in df.columns]
    if key:
        return key, "natural"
    key = [c for c in df.columns if c not in VOLATILE_COLS]
    return key, "all-cols-minus-volatile"


def main():
    if not os.path.isdir(OUTPUT_DIR):
        raise SystemExit(f"Not found: {OUTPUT_DIR}")

    files = sorted(f for f in os.listdir(OUTPUT_DIR) if f.endswith(".xlsx"))
    print("=" * 96)
    print(f"  ARCHIVE REPAIR — {OUTPUT_DIR}")
    print(f"  mode: {'APPLY (files will be rewritten)' if APPLY else 'REPORT ONLY (use --apply to write)'}")
    print("=" * 96)
    print(f"  {'feed':<46} {'before':>9} {'after':>9} {'dropped':>9}  key")
    print("-" * 96)

    tot_before = tot_after = 0
    changed = []

    for fn in files:
        path = os.path.join(OUTPUT_DIR, fn)
        try:
            df = pd.read_excel(path)
        except Exception as e:
            print(f"  {fn[:-5]:<46} {'ERR':>9}  {type(e).__name__}: {e}")
            continue

        before = len(df)

        # placeholder-only file: leave it alone
        if len(df.columns) == 1 and "Message" in df.columns:
            print(f"  {fn[:-5]:<46} {'-':>9} {'-':>9} {'-':>9}  (placeholder, skipped)")
            continue

        # strip any placeholder rows mixed into real data
        if "Message" in df.columns:
            df = df[~df["Message"].astype(str).str.contains("No records found", na=False)]
            if df["Message"].isna().all():
                df = df.drop(columns=["Message"])

        # backfill point-in-time stamp for rows collected before the fix
        backfilled = "fetched_at" not in df.columns
        if backfilled:
            mtime = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S")
            df["fetched_at"] = f"{mtime} (backfilled)"

        # Use the scraper's own dedupe so repair and steady-state agree exactly.
        key, kind = natural_key(df, fn)
        df = _dedupe(df, fn)

        after = len(df)
        dropped = before - after
        tot_before += before
        tot_after += after

        flag = " <<<" if dropped else ""
        print(f"  {fn[:-5]:<46} {before:>9,} {after:>9,} {dropped:>9,}  {kind}{flag}")

        if dropped or backfilled:
            changed.append((fn, path, df, before, after))

    print("-" * 96)
    print(f"  {'TOTAL':<46} {tot_before:>9,} {tot_after:>9,} {tot_before - tot_after:>9,}")
    print("=" * 96)

    if not changed:
        print("\n  Nothing to repair.")
        return

    if not APPLY:
        print(f"\n  {len(changed)} file(s) would be rewritten, "
              f"{tot_before - tot_after:,} phantom rows removed.")
        print("  Re-run with --apply to write.  (Back up output_excels/ first.)")
        return

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safety = os.path.join(os.path.dirname(OUTPUT_DIR), f"_prerepair_{stamp}")
    os.makedirs(safety, exist_ok=True)
    print(f"\n  Safety copy -> {safety}")

    for fn, path, df, before, after in changed:
        shutil.copy2(path, os.path.join(safety, fn))
        sheet = fn[3:-5][:31] or "Sheet1"
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as w:
                df.to_excel(w, sheet_name=sheet, index=False)
                ws = w.sheets[sheet]
                sample = df.head(200)
                for i, col in enumerate(df.columns, start=1):
                    # An all-null column gives NaN here, and NaN is truthy, so
                    # `int(x or 0)` raises. Guard explicitly.
                    m = sample[col].astype(str).str.len().max()
                    width = max(len(str(col)), int(m) if pd.notna(m) else 0)
                    ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = \
                        min(max(width + 3, 12), 60)
            print(f"    [WRITTEN] {fn}: {before:,} -> {after:,}")
        except PermissionError:
            print(f"    [LOCKED]  {fn}: open in Excel — skipped")
        except Exception as e:
            print(f"    [FAILED]  {fn}: {type(e).__name__}: {e}")

    print(f"\n  Done. {tot_before - tot_after:,} phantom rows removed.")


if __name__ == "__main__":
    main()
