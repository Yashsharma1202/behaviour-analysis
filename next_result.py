"""
next_result.py
===============================================================================
For each stock: the LAST (most recent) quarterly result with
    * price 6 trading days BEFORE the result day
    * price ON the result day
    * price 3 trading days AFTER the result day
...and a PREDICTION for the NEXT result, built from how this stock has behaved
around ALL its past results (average move + how often it went up), together with
the next upcoming result date read from the board-meeting feed.

    python next_result.py                 # all stocks with downloaded feeds
    python next_result.py RELIANCE INFY   # specific stocks
    python next_result.py 6 3             # BEFORE=6, AFTER=3 (numbers = window)

Output: printed report + processed/next_result.csv
===============================================================================
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

import event_behaviour as EB

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                     # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
PROC.mkdir(exist_ok=True)

BEFORE = 6
AFTER = 3


def next_result_date(sym: str, today: pd.Timestamp):
    """Earliest FUTURE board meeting whose purpose mentions results (the next result)."""
    path = ROOT / sym / "board_meetings.csv"
    if not path.exists():
        return None
    df = pd.read_csv(path, dtype=str).fillna("")
    if "bm_date" not in df.columns:
        return None
    df["_d"] = pd.to_datetime(df["bm_date"], errors="coerce", dayfirst=True, format="mixed")
    purpose = (df.get("bm_purpose", "") + " " + df.get("bm_desc", "")).str.lower()
    fut = df[(df["_d"] > today) & purpose.str.contains("result", na=False)]
    return fut["_d"].min() if not fut.empty else None


def analyse(sym: str, px: pd.Series, today: pd.Timestamp) -> dict | None:
    events = EB.load_events_from_feeds([sym])
    results = sorted(events[events["event_type"] == "RESULTS"]["event_date"].drop_duplicates())
    if not results:
        return None
    idx = px.index

    # ---- historical stats across ALL past results (the "prediction" basis) ----
    aft_moves, bef_moves, tot_moves = [], [], []
    last = None
    for d in results:
        pos = idx.searchsorted(pd.Timestamp(d))
        if pos - BEFORE < 0 or pos + AFTER >= len(idx):
            continue
        pb, pr, pa = float(px.iloc[pos - BEFORE]), float(px.iloc[pos]), float(px.iloc[pos + AFTER])
        bef_moves.append((pr / pb - 1) * 100)
        aft_moves.append((pa / pr - 1) * 100)
        tot_moves.append((pa / pb - 1) * 100)
        last = {"date": idx[pos].date(), "pb": pb, "pr": pr, "pa": pa,
                "bef": (pr / pb - 1) * 100, "aft": (pa / pr - 1) * 100,
                "tot": (pa / pb - 1) * 100}
    if last is None:
        return None

    n = len(tot_moves)
    hist = {
        "n": n,
        "avg_bef": sum(bef_moves) / n,
        "avg_aft": sum(aft_moves) / n,
        "avg_tot": sum(tot_moves) / n,
        "win_aft": sum(1 for x in aft_moves if x > 0) / n * 100,
        "win_tot": sum(1 for x in tot_moves if x > 0) / n * 100,
    }
    nxt = next_result_date(sym, today)
    last_close = float(px.iloc[-1])
    # rough expected price 3 days after the next result, from the historical avg
    pred_price = last_close * (1 + hist["avg_tot"] / 100)

    return {"symbol": sym, "last": last, "hist": hist,
            "next_date": nxt.date() if nxt is not None else None,
            "last_close": last_close, "pred_after_price": pred_price}


def main():
    global BEFORE, AFTER
    args = sys.argv[1:]
    nums = [a for a in args if a.isdigit()]
    syms = [a.upper() for a in args if not a.isdigit()]
    if len(nums) >= 1:
        BEFORE = int(nums[0])
    if len(nums) >= 2:
        AFTER = int(nums[1])
    if not syms:
        syms = EB.discover_feed_symbols()
    if not syms:
        raise SystemExit("No downloaded feeds. Run 'python download_feeds.py nifty50' first.")

    today = pd.Timestamp.today().normalize()
    print(f"LAST result (px {BEFORE}d before / on day / {AFTER}d after) + NEXT-result prediction")
    print(f"as of {today.date()}  ·  {len(syms)} stock(s)\n")

    out = []
    for s in syms:
        px = EB.fetch_prices_yahoo(s)
        if px is None or len(px) < BEFORE + AFTER + 5:
            continue
        r = analyse(s, px, today)
        if r:
            out.append(r)

    if not out:
        raise SystemExit("No result events inside price history.")

    # ---- table: last result snapshot ----
    print("=" * 118)
    print(f"  {'symbol':>11} {'last result':>12} {f'{BEFORE}d bef':>9} {'on day':>9} "
          f"{f'{AFTER}d aft':>9} {'bef→res':>8} {'res→aft':>8} | "
          f"{'hist res→aft':>12} {'win%':>5} {'n':>3}  next result")
    print("=" * 118)
    rows = []
    for r in sorted(out, key=lambda x: (x["next_date"] is None, x["next_date"] or today.date())):
        L, H = r["last"], r["hist"]
        nd = str(r["next_date"]) if r["next_date"] else "—"
        print(f"  {r['symbol']:>11} {L['date']!s:>12} {L['pb']:>9,.1f} {L['pr']:>9,.1f} "
              f"{L['pa']:>9,.1f} {L['bef']:>+7.2f}% {L['aft']:>+7.2f}% | "
              f"{H['avg_aft']:>+11.2f}% {H['win_aft']:>4.0f}% {H['n']:>3}  {nd}")
        rows.append({
            "symbol": r["symbol"], "last_result_day": L["date"],
            f"price_{BEFORE}d_before": round(L["pb"], 2),
            "price_on_result": round(L["pr"], 2),
            f"price_{AFTER}d_after": round(L["pa"], 2),
            "last_before_to_result_%": round(L["bef"], 2),
            "last_result_to_after_%": round(L["aft"], 2),
            "hist_avg_before_to_result_%": round(H["avg_bef"], 2),
            "hist_avg_result_to_after_%": round(H["avg_aft"], 2),
            "hist_avg_total_%": round(H["avg_tot"], 2),
            "hist_win_rate_after_%": round(H["win_aft"], 0),
            "n_results": H["n"],
            "next_result_date": r["next_date"],
            "last_close": round(r["last_close"], 2),
            "predicted_after_price": round(r["pred_after_price"], 2),
        })
    print("=" * 118)

    pd.DataFrame(rows).to_csv(PROC / "next_result.csv", index=False)

    # ---- predictions for stocks with an upcoming result ----
    upcoming = [r for r in out if r["next_date"]]
    if upcoming:
        print(f"\n  PREDICTION FOR THE NEXT RESULT  ({len(upcoming)} stocks have one scheduled):")
        print("  " + "-" * 100)
        for r in sorted(upcoming, key=lambda x: x["next_date"]):
            H = r["hist"]
            lean = "UP" if H["avg_aft"] > 0 else "DOWN"
            print(f"   {r['symbol']:<11} next result ~{r['next_date']}:  "
                  f"expect ~{H['avg_aft']:+.2f}% in the {AFTER} days after "
                  f"({H['win_aft']:.0f}% of past {H['n']} went up) → lean {lean}.  "
                  f"Historically best to buy ~{BEFORE}d before.")
    print(f"\n  Saved -> {PROC}\\next_result.csv")
    print("  Note: prices are Yahoo adjusted close; predictions are historical averages, "
          "NOT guarantees.")


if __name__ == "__main__":
    main()
