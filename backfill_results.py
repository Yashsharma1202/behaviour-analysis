"""
backfill_results.py
===============================================================================
Build a Quarterly-Results feed for the handful of stocks that have no
financial_results.csv (NSE isn't reachable from here to download it), by deriving
the result dates from their own announcements feed — the "Financial Result
Updates" disclosures, which ARE the quarterly-result announcements.

ADDITIVE & SAFE: only writes financial_results.csv for a stock that doesn't
already have one; never touches a stock that has a real NSE results feed.

RUN
    python backfill_results.py --dry     # report only
    python backfill_results.py           # write
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

from download_feeds import NIFTY50_FALLBACK

# the announcement categories that ARE a quarterly-result disclosure
RESULT_DESC = {"Financial Result Updates", "Financial Results Updates"}


def main() -> None:
    universe = sorted(set(NIFTY50_FALLBACK))
    done = []
    for sym in universe:
        fr = ROOT / sym / "financial_results.csv"
        # skip stocks that already have a (non-empty) results feed
        if fr.exists():
            try:
                if len(pd.read_csv(fr, dtype=str)):
                    continue
            except Exception:                          # noqa: BLE001
                continue
        ann = ROOT / sym / "announcements.csv"
        if not ann.exists():
            continue
        a = pd.read_csv(ann, dtype=str).fillna("")
        m = a[a["desc"].str.strip().isin(RESULT_DESC)].copy()
        if m.empty:
            continue
        datecol = "an_dt" if "an_dt" in m.columns else "sort_date"
        dt = pd.to_datetime(m[datecol], errors="coerce", format="mixed")
        m = m[dt.notna().values]
        out = pd.DataFrame({
            "broadCastDate": dt.dropna().dt.strftime("%d-%b-%Y").values,
            "symbol": sym,
            "companyName": m.get("sm_name", sym).values,
            "isin": m.get("sm_isin", "").values,
            "consolidated": "",
            "period": "Quarterly",
            "resultDescription": "Financial Result Updates",
            "fr_source": "announcements",
        }).drop_duplicates(subset="broadCastDate")
        out["_s"] = pd.to_datetime(out["broadCastDate"], format="mixed", errors="coerce")
        out = out.sort_values("_s", ascending=False).drop(columns="_s")
        done.append((sym, len(out), out["broadCastDate"].iloc[-1], out["broadCastDate"].iloc[0]))
        if not DRY:
            out.to_csv(fr, index=False)

    print(f"{'stock':<12}{'results':>8}   span")
    for s, n, old, new in done:
        print(f"{s:<12}{n:>8}   {old}  ->  {new}")
    print(f"\n{len(done)} stocks backfilled." + ("  [DRY RUN — nothing written]" if DRY else "  Written."))


if __name__ == "__main__":
    main()
