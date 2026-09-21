"""
build_universe.py
===============================================================================
Pick the research universe once, write it down, and have everything else read
it from there.

WHY A FILE INSTEAD OF A CONSTANT
--------------------------------
disclosure_timing.py, score_feeds.py, backtest_sentiment.py and the dashboard
each decided their own universe, mostly by intersecting with
download_feeds.NIFTY50_FALLBACK. Once the universe is 400+ names that has to
come from one place, or the CAR panel and the backtest quietly disagree about
which stocks exist.

    processed/universe.json   {generated, index, symbols: [{symbol, industry}]}

SECTOR COMES FROM HERE TOO
--------------------------
NSE's static index CSV carries an Industry column, so sector is known for a
stock BEFORE its feeds are downloaded. The alternative, `smIndustry` inside
each announcements.csv, only exists after the fact and is unavailable for any
stock that failed to download. Stage-3 sector-neutral construction needs the
former.

SELECTION
    Nifty 500 constituents, kept if quarterly/<SYM>.csv already exists (so the
    fundamentals and estimate machinery work). No size or liquidity filter —
    the whole point of expanding is to reach down the cap curve, where
    post-news overreaction is expected to be strongest. Filtering to the
    biggest names would rebuild the problem we are escaping.

    python build_universe.py                 # Nifty 500 -> universe.json
    python build_universe.py --index "NIFTY 200"
    python build_universe.py --require-feeds  # only stocks already downloaded
===============================================================================
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "processed" / "universe.json"


def load() -> dict:
    """The universe every other module should use. Falls back to the Nifty 50
    list if build_universe.py has never been run."""
    if OUT.exists():
        try:
            return json.loads(OUT.read_text(encoding="utf-8"))
        except Exception:                         # noqa: BLE001
            pass
    from download_feeds import NIFTY50_FALLBACK
    return {"index": "NIFTY 50 (fallback)", "generated": None,
            "symbols": [{"symbol": s, "industry": ""}
                        for s in sorted(set(NIFTY50_FALLBACK))]}


def symbols() -> list[str]:
    return [r["symbol"] for r in load()["symbols"]]


def industry_map() -> dict[str, str]:
    return {r["symbol"]: r.get("industry", "") for r in load()["symbols"]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--index", default="NIFTY 500")
    ap.add_argument("--require-feeds", action="store_true",
                    help="keep only stocks whose feeds are already downloaded")
    ap.add_argument("--min-expected", type=int, default=300)
    args = ap.parse_args()

    import download_feeds as D
    members = D.index_members(args.index, args.min_expected)
    if not members:
        raise SystemExit(f"could not fetch {args.index} constituents")

    kept, dropped = [], []
    for r in members:
        s = r["symbol"]
        if not (ROOT / "quarterly" / f"{s}.csv").exists():
            dropped.append((s, "no quarterly/ fundamentals"))
            continue
        if args.require_feeds and not (ROOT / s / "financial_results.csv").exists():
            dropped.append((s, "feeds not downloaded"))
            continue
        kept.append({"symbol": s, "industry": r["industry"],
                     "company": r["company"]})

    ready = sum(1 for r in kept
                if (ROOT / r["symbol"] / "financial_results.csv").exists())

    payload = {"index": args.index,
               "generated": __import__("datetime").date.today().isoformat(),
               "n": len(kept), "symbols": kept}
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(payload, indent=1), encoding="utf-8")

    print("=" * 72)
    print(f"  UNIVERSE — {args.index}")
    print("=" * 72)
    print(f"  constituents          {len(members):>5}")
    print(f"  kept                  {len(kept):>5}   (have quarterly/ fundamentals)")
    print(f"  dropped               {len(dropped):>5}")
    print(f"  feeds already on disk {ready:>5}")
    print(f"  still to download     {len(kept) - ready:>5}"
          f"   (~{(len(kept) - ready) * 5 / 60:.0f} min at 2s/stock + fetch)")

    import collections
    ind = collections.Counter(r["industry"] for r in kept)
    print(f"\n  {len(ind)} industries, largest:")
    for name, n in ind.most_common(8):
        print(f"    {n:>4}  {name}")
    print(f"\n  -> {OUT}")
    print(f"\n  next:  python download_feeds.py $(python -c "
          f"\"import build_universe as U;print(' '.join(U.symbols()))\")")


if __name__ == "__main__":
    main()
