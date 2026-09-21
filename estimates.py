"""
estimates.py
===============================================================================
Merge the providers in estimate_providers.py into one forward view per stock:
what next quarter probably looks like, and when it will actually be reported.

MERGE POLICY
------------
External consensus wins when it is both confident and fresh; otherwise the
company's own history does. The payload always names the winner, so a scraped
number is never displayed as if it were ours.

TWO MODES, AND THE DIFFERENCE MATTERS
-------------------------------------
    for_display   every provider, newest data, whatever is available
    for_model     own_history + calendar ONLY, evaluated as of a given date

Phase 5 must use for_model. Today's consensus applied to a 2023 event is
look-ahead bias, and no amount of purged cross-validation detects it — the
leak is in the feature, not the split.

    python estimates.py                  # Nifty 50 table
    python estimates.py RELIANCE TCS     # detail for a few
    python estimates.py --providers screener   # enable the scraper
===============================================================================
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import pandas as pd

import estimate_providers as P

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "processed" / "estimates"


def build(sym: str, for_model: bool = False,
          asof: pd.Timestamp | None = None, fetcher=None) -> dict:
    """The full forward view for one symbol."""
    ests: dict[str, dict] = {}

    for metric in ("Net Profit", "Sales"):
        e = P.own_history(sym, metric=metric, asof=asof)
        if e:
            ests[f"own_history:{metric}"] = e.to_dict()

    cal = P.calendar(sym, today=asof)
    if cal:
        ests["calendar"] = cal.to_dict()

    if not for_model:
        for fn, kw in ((P.web_search, {"fetcher": fetcher}), (P.screener_scrape, {})):
            try:
                e = fn(sym, **kw)
            except Exception:                     # noqa: BLE001
                e = None
            if e:
                ests[e.source] = e.to_dict()

    # ---- pick the headline PAT estimate ---------------------------------- #
    own = ests.get("own_history:Net Profit")
    best, why = own, "own_history (no external consensus available)"
    for k in ("screener_scrape", "web_search"):
        cand = ests.get(k)
        if cand and cand.get("value") is not None and cand["confidence"] >= 0.6:
            best, why = cand, f"{k} (confident and fresh; own_history kept as fallback)"
            break

    return {
        "symbol": sym,
        "for_model": for_model,
        "asof": str(asof.date()) if asof is not None else None,
        "estimates": ests,
        "headline": best,
        "headline_source_note": why,
        "next_result": (ests.get("calendar") or {}).get("detail"),
        "providers_used": sorted(ests),
        "model_safe_only": for_model,
    }


def growth_gap(sym: str, asof: pd.Timestamp | None = None) -> float | None:
    """`est_growth_gap` — the Phase-5 ML feature.

    How far the company's expected next-quarter growth sits from its own longer
    trend. Positive means the recent four quarters are accelerating against the
    company's history; negative means decelerating. Built ONLY from own_history,
    and only from quarters filed on or before `asof`.
    """
    e = P.own_history(sym, "Net Profit", asof=asof)
    if not e:
        return None
    import fund_loader
    try:
        q = fund_loader.load_stock(sym).get("quarterly")
    except Exception:                             # noqa: BLE001
        return None
    if q is None or "Net Profit" not in q.columns:
        return None
    s = pd.to_numeric(q["Net Profit"], errors="coerce").dropna()
    if asof is not None:
        s = s[s.index <= asof]
    yoy = (s / s.shift(4) - 1).dropna()
    if len(yoy) < 8:
        return None
    return round(float(e.detail["yoy_growth_median_4q"] / 100 - yoy.median()), 4)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("symbols", nargs="*")
    ap.add_argument("--providers", nargs="*", default=[],
                    help="extra providers to enable: screener")
    ap.add_argument("--save", action="store_true", help="write processed/estimates/<SYM>.json")
    args = ap.parse_args()

    if "screener" in args.providers:
        P.SCREENER_ENABLED = True
        print("  screener_scrape ENABLED — respect the site's terms of service\n")

    syms = [s.upper() for s in args.symbols]
    if not syms:
        from download_feeds import NIFTY50_FALLBACK
        syms = sorted(set(NIFTY50_FALLBACK))

    print("=" * 92)
    print("  FORWARD ESTIMATES — next-quarter Net Profit from the company's own history")
    print("=" * 92)
    print(f"  {'symbol':<13}{'quarter':>10}{'estimate':>13}{'low':>12}{'high':>12}"
          f"{'YoY%':>8}{'conf':>7}   next result")
    print("  " + "-" * 88)

    rows = []
    for sym in syms:
        d = build(sym)
        pat = d["estimates"].get("own_history:Net Profit")
        nr = d.get("next_result") or {}
        if not pat:
            print(f"  {sym:<13}{'—':>10}   (not enough quarterly history)")
            continue
        det = pat["detail"]
        nxt = (f"{nr.get('next_result_date')} ({nr.get('days_away')}d)"
               if nr.get("next_result_date") else "—")
        print(f"  {sym:<13}{det['for_quarter']:>10}{pat['value']:>13,.0f}"
              f"{det['low']:>12,.0f}{det['high']:>12,.0f}"
              f"{det['yoy_growth_median_4q']:>8.1f}{pat['confidence']:>7.2f}   {nxt}")
        rows.append(d)
        if args.save:
            OUT_DIR.mkdir(parents=True, exist_ok=True)
            (OUT_DIR / f"{sym}.json").write_text(json.dumps(d, indent=1),
                                                 encoding="utf-8")

    print(f"\n  {len(rows)}/{len(syms)} symbols estimated"
          + (f"  ->  {OUT_DIR}" if args.save else ""))
    print("  Figures are in the units of quarterly/<SYM>.csv (Rs crore) and are a")
    print("  seasonal-naive extrapolation, NOT analyst consensus. See estimates.py.")

    if len(syms) <= 3:
        for d in rows:
            print(f"\n  {d['symbol']}  headline: {d['headline_source_note']}")
            print(f"    providers: {', '.join(d['providers_used'])}")
            print(f"    est_growth_gap (ML feature): {growth_gap(d['symbol'])}")


if __name__ == "__main__":
    main()
