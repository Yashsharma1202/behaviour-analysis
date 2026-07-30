"""
enrich_corporate_actions.py
===============================================================================
Deepen the corporate-actions feed. NSE's feed is capped at ~20 entries per stock,
so quarterly-dividend payers (e.g. TCS) get chopped to just the last few years,
which inflates their win rates. Yahoo Finance carries the COMPLETE dividend &
split ex-date history (back to ~2007), so we backfill from there.

For each stock:
  • fetch dividend + split ex-dates from Yahoo (events=div,split)
  • add any Yahoo ex-date that isn't already within a few days of an existing
    NSE ex-date (so we extend history without creating near-duplicates)
  • merge into <SYM>/corporate_actions.csv, sort newest-first

ADDITIVE: existing NSE rows are kept as-is; we only ADD older ex-dates.

RUN
    python enrich_corporate_actions.py --dry     # report only
    python enrich_corporate_actions.py           # write
===============================================================================
"""
from __future__ import annotations

import datetime as dt
import json
import ssl
import sys
import time
import urllib.request
import warnings
from pathlib import Path

import pandas as pd

warnings.filterwarnings("ignore")
sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).resolve().parent
DRY = "--dry" in sys.argv
TOL_DAYS = 3          # treat a Yahoo ex-date within this many days of an NSE one as the same event

from download_feeds import NIFTY50_FALLBACK

_CTX = ssl.create_default_context(); _CTX.check_hostname = False; _CTX.verify_mode = ssl.CERT_NONE
_HDR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def yahoo_actions(sym: str) -> list[tuple[dt.date, str]]:
    """(ex-date, subject) for every dividend & split from Yahoo, oldest first."""
    from urllib.parse import quote
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{quote(sym, safe='')}.NS"
           f"?period1=1167609600&period2={int(time.time())}&interval=1d&events=div,split")
    try:
        req = urllib.request.Request(url, headers=_HDR)
        d = json.load(urllib.request.urlopen(req, timeout=30, context=_CTX))
        ev = d["chart"]["result"][0].get("events", {})
    except Exception as e:                          # noqa: BLE001
        print(f"    ! {sym}: {type(e).__name__}")
        return []
    out = []
    for v in ev.get("dividends", {}).values():
        d0 = dt.datetime.fromtimestamp(v["date"], dt.timezone.utc).date()
        amt = v.get("amount")
        out.append((d0, f"Dividend - Rs {amt} Per Share" if amt else "Dividend"))
    for v in ev.get("splits", {}).values():
        d0 = dt.datetime.fromtimestamp(v["date"], dt.timezone.utc).date()
        out.append((d0, v.get("splitRatio", "Split")))
    return sorted(out)


def main() -> None:
    universe = sorted(set(NIFTY50_FALLBACK))
    print(f"Backfilling corporate actions from Yahoo for {len(universe)} stocks"
          + ("   [DRY RUN]" if DRY else "") + "\n")
    tot_before = tot_after = 0
    changed = []
    for sym in universe:
        p = ROOT / sym / "corporate_actions.csv"
        old = pd.read_csv(p, dtype=str).fillna("") if p.exists() else pd.DataFrame()
        old_dts = sorted(pd.to_datetime(old.get("exDate"), errors="coerce", format="mixed")
                         .dropna().dt.date.unique()) if len(old) else []
        before = len(old_dts)

        ya = yahoo_actions(sym)
        add = []
        kept = list(old_dts)
        for d0, subj in ya:
            if all(abs((d0 - e).days) > TOL_DAYS for e in kept):   # not already present
                ds = d0.strftime("%d-%b-%Y")
                # Yahoo gives only the ex-date; the record date is the same day
                # (T+1 settlement) or ~1 day later — use the ex-date so the
                # RECORD DATE column isn't blank in the dashboard.
                add.append({"exDate": ds, "recDate": ds, "subject": subj,
                            "symbol": sym, "comp": sym, "company_name": sym,
                            "ca_source": "yahoo"})
                kept.append(d0)
        after = len(kept)
        tot_before += before; tot_after += after
        if add:
            changed.append((sym, before, after))
            if not DRY:
                cols = list(dict.fromkeys(list(old.columns) + list(add[0].keys())))
                big = pd.concat([old.reindex(columns=cols),
                                 pd.DataFrame(add).reindex(columns=cols)],
                                ignore_index=True).fillna("")
                big["_s"] = pd.to_datetime(big["exDate"], errors="coerce", format="mixed")
                big = big.sort_values("_s", ascending=False, na_position="last").drop(columns="_s")
                big.to_csv(p, index=False)
        time.sleep(0.3)

    changed.sort(key=lambda r: r[2] - r[1], reverse=True)
    print(f"{'stock':<12}{'before':>7}{'after':>7}{'+':>6}")
    for s, b, a in changed[:18]:
        print(f"{s:<12}{b:>7}{a:>7}{a-b:>+6}")
    if len(changed) > 18:
        print(f"   ... and {len(changed)-18} more")
    print(f"\nStocks deepened: {len(changed)} | corporate-action dates {tot_before} -> {tot_after} (+{tot_after-tot_before})")
    print("Dry run — nothing written." if DRY else "Written.")


if __name__ == "__main__":
    main()
