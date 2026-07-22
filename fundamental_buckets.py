"""
fundamental_buckets.py
===============================================================================
DOES THE EVENT EDGE LIVE IN A PARTICULAR KIND OF COMPANY?

The per-stock search (optimal_rules.py) failed because ~50 events cannot support
a 130-cell grid search. This fixes that: POOL all 54 stocks and split the events
by what the company looked like AT THE TIME of the event. Each bucket then holds
hundreds-to-thousands of events, so a t-stat finally means something.

BUCKETS (each event is ranked against its peers in the SAME calendar year, so
"high ROCE" means high relative to the market then — not high in absolute terms,
which drifts over a decade):

    GROWTH        Sales Growth %, Profit Growth %
    PROFITABILITY ROCE %, OPM %
    LEVERAGE      Borrowings / Reserves
    VALUATION     P/E  (price at the event ÷ EPS of the last published year)

LOOKAHEAD CONTROL — the whole point:
  * a fiscal year (Mar-YYYY) is only PUBLIC when its annual result is broadcast.
    We read that date from the company's own financial_results feed, so an event
    in Feb-2020 sees FY2019 numbers, never FY2020's.
  * entry is never allowed before the event itself was public (results: from the
    board-meeting intimation; announcements: only after the news is out).
  * the rule is fitted on the oldest 70% of events and reported on the unseen 30%.
  * a random-date PLACEBO runs alongside every bucket.
  * t is MONTH-CLUSTERED: pooled trades across 54 stocks share the same market
    days, so the naive t is badly inflated. t(month) is the one to read.

    python fundamental_buckets.py            # RESULTS + ANNOUNCEMENT
    python fundamental_buckets.py RESULTS    # one event type

Output: processed/fundamental_buckets.csv + a printed report
===============================================================================
"""
from __future__ import annotations

import random
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import event_behaviour as EB
import nifty_backtest as NBT
from optimal_rules import load_events

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                  # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"

COST = 0.001
SPLIT = 0.70
ENTRY = range(-8, 3)
EXIT = range(-1, 13)
MIN_TRAIN = 60          # a grid cell needs this many train events (pooled -> can be strict)
N_PLACEBO = 600
SEED = 7


def _num(x):
    if not isinstance(x, str):
        return np.nan
    x = x.strip().replace(",", "").replace("%", "")
    try:
        return float(x)
    except ValueError:
        return np.nan


def _fy(col: str) -> int | None:
    """'Mar 2016' -> 2016.  Some stocks have odd 15-month periods ('Mar 201615m')."""
    m = re.search(r"(\d{4})", col)
    return int(m.group(1)) if m else None


def load_fundamentals(sym: str) -> pd.DataFrame | None:
    """-> one row per fiscal year, with the date that year became PUBLIC."""
    want = {
        "pnl": ["Sales Growth %", "Profit Growth %", "OPM %", "EPS in Rs"],
        "ratios": ["ROCE %"],
        "balance_sheet": ["Borrowings", "Reserves"],
    }
    # the three folders don't always carry the same year columns for a stock, so
    # key every metric by FISCAL YEAR and align on that, never on column position.
    vals: dict[str, dict[int, float]] = {}
    for folder, metrics in want.items():
        f = ROOT / folder / f"{sym}.csv"
        if not f.exists():
            return None
        df = pd.read_csv(f, dtype=str).fillna("")
        df["m"] = df["Metric"].str.replace("└─", "", regex=False).str.strip()
        cols = [c for c in df.columns if c.startswith("Mar ")]
        if not cols:
            return None
        for met in metrics:
            row = df[df["m"] == met]
            vals[met] = ({_fy(c): _num(row.iloc[0][c]) for c in cols if _fy(c)}
                         if len(row) else {})
    years = sorted({y for d in vals.values() for y in d})
    if not years:
        return None
    fu = pd.DataFrame({met: [d.get(y, np.nan) for y in years]
                       for met, d in vals.items()})
    fu["fy"] = years

    # when did each fiscal year actually become public? read it from the company's
    # OWN annual-result broadcast date. fall back to 1 June (SEBI: audited annual
    # results within 60 days of the 31-Mar year end).
    pub = {}
    fpath = ROOT / sym / "financial_results.csv"
    if fpath.exists():
        fr = pd.read_csv(fpath, dtype=str).fillna("")
        if {"period", "toDate", "broadCastDate"} <= set(fr.columns):
            fr = fr[fr["period"] == "Annual"]
            td = pd.to_datetime(fr["toDate"], errors="coerce", dayfirst=True)
            bd = pd.to_datetime(fr["broadCastDate"], errors="coerce",
                                format="mixed", dayfirst=True)
            for t, b in zip(td, bd):
                if pd.isna(t) or pd.isna(b):
                    continue
                y = t.year if t.month >= 3 else t.year
                pub[y] = min(pub.get(y, b), b)
    fu["public_from"] = [pub.get(y, pd.Timestamp(y, 6, 1)) for y in fu["fy"]]
    fu["leverage"] = fu["Borrowings"] / fu["Reserves"].replace(0, np.nan)
    return fu


DIMS = {
    "GROWTH · Sales Growth %":   "Sales Growth %",
    "GROWTH · Profit Growth %":  "Profit Growth %",
    "PROFIT · ROCE %":           "ROCE %",
    "PROFIT · OPM %":            "OPM %",
    "LEVERAGE · Debt/Reserves":  "leverage",
    "VALUATION · P/E":           "pe",
}


def build_events(symbols, mkt_ix):
    """Every event, with the fundamentals that were PUBLIC on that date attached."""
    rows = []
    for i, sym in enumerate(symbols, 1):
        px = EB.fetch_prices_yahoo(sym)
        fu = load_fundamentals(sym)
        if px is None or fu is None or len(px) < 200:
            continue
        ev = load_events(sym)
        if ev.empty:
            continue
        I = px.index
        fu = fu.sort_values("public_from")
        for _, e in ev.iterrows():
            if e["type"] not in ("RESULTS", "ANNOUNCEMENT"):
                continue
            d = pd.Timestamp(e["d"])
            pos = I.searchsorted(d)
            if pos >= len(I) or pos < 10:
                continue
            # ← the ONLY fiscal year we are allowed to see: the last one published
            #    strictly before this event happened.
            vis = fu[fu["public_from"] < d]
            if vis.empty:
                continue
            f = vis.iloc[-1]
            eps = f["EPS in Rs"]
            price = float(px.iloc[pos])
            rows.append({
                "symbol": sym, "d": d, "type": e["type"], "known": pd.Timestamp(e["known"]),
                "pos": pos, "fy_used": int(f["fy"]),
                "Sales Growth %": f["Sales Growth %"], "Profit Growth %": f["Profit Growth %"],
                "ROCE %": f["ROCE %"], "OPM %": f["OPM %"], "leverage": f["leverage"],
                "pe": (price / eps) if (eps and eps == eps and eps > 0) else np.nan,
            })
        if i % 15 == 0:
            print(f"   loaded {i}/{len(symbols)} stocks…")
    return pd.DataFrame(rows)


def returns_grid(ev, prices, mkts):
    """dict[(entry,exit)] -> array of market-adjusted returns aligned with ev."""
    out = {}
    n = len(ev)
    sym = ev["symbol"].values
    pos = ev["pos"].values
    kpos = np.array([prices[s].index.searchsorted(k)
                     for s, k in zip(sym, ev["known"])])
    for ei in ENTRY:
        for xo in EXIT:
            if xo <= ei:
                continue
            v = np.full(n, np.nan)
            for j in range(n):
                p, m = prices[sym[j]].values, mkts[sym[j]]
                a, b = pos[j] + ei, pos[j] + xo
                if a < 0 or b >= len(p) or a < kpos[j]:      # ← knowability
                    continue
                g = p[b] / p[a] - 1 - COST
                v[j] = g - (m[b] / m[a] - 1)
            if np.isfinite(v).sum() >= 30:
                out[(ei, xo)] = v
    return out


def tm(v, months):
    """month-clustered t: average within each calendar month, then t-test."""
    ok = ~np.isnan(v)
    if ok.sum() < 5:
        return float("nan")
    s = pd.Series(v[ok]).groupby(pd.Series(months)[ok].values).mean()
    if len(s) < 3 or s.std(ddof=1) == 0:
        return float("nan")
    return s.mean() / (s.std(ddof=1) / len(s) ** 0.5)


def main():
    only = [a.upper() for a in sys.argv[1:]] or ["RESULTS", "ANNOUNCEMENT"]
    syms = [s for s in EB.discover_feed_symbols()
            if (ROOT / "pnl" / f"{s}.csv").exists()]
    mkt_ix = NBT.fetch_index()

    print(f"FUNDAMENTAL BUCKETS — {len(syms)} stocks with feeds AND fundamentals")
    print("events split by what the company looked like AT THE TIME (lagged to the")
    print(f"annual-result publication date) · {COST*100:.2f}% cost · market-adjusted\n")

    ev = build_events(syms, mkt_ix)
    if ev.empty:
        raise SystemExit("no events with fundamentals")

    prices, mkts = {}, {}
    for s in ev["symbol"].unique():
        p = EB.fetch_prices_yahoo(s)
        prices[s] = p
        mkts[s] = mkt_ix.reindex(p.index).ffill().bfill().values

    print(f"\n   {len(ev):,} events carry usable fundamentals  "
          f"({ev['d'].min().date()} → {ev['d'].max().date()})")

    out = []
    for et in only:
        sub = ev[ev["type"] == et].sort_values("d").reset_index(drop=True)
        if len(sub) < 200:
            print(f"\n   {et}: only {len(sub)} events — skipped")
            continue
        # rank each event against its PEERS IN THE SAME YEAR -> tercile
        sub["yr"] = sub["d"].dt.year
        sub["mo"] = sub["d"].dt.to_period("M").astype(str)
        grids = returns_grid(sub, prices, mkts)
        if not grids:
            continue
        cut = int(len(sub) * SPLIT)
        is_tr = np.arange(len(sub)) < cut

        print("\n" + "=" * 112)
        print(f"  {et}   ({len(sub):,} events, fitted on the oldest {SPLIT:.0%}, "
              f"reported on the unseen {1-SPLIT:.0%})")
        print("=" * 112)

        for label, col in DIMS.items():
            v = sub[col]
            if v.notna().sum() < 200:
                continue
            q = v.groupby(sub["yr"]).transform(
                lambda s: pd.qcut(s.rank(method="first"), 3,
                                  labels=["LOW", "MID", "HIGH"])
                if s.notna().sum() >= 6 else pd.Series(index=s.index, dtype=object))
            print(f"\n  {label}")
            print(f"    {'bucket':<7} {'rule':<17} {'TEST avg':>9} {'win':>5} "
                  f"{'t(mo)':>7} {'n':>6} │ {'random':>8} {'EDGE':>8}  verdict")
            print("    " + "-" * 98)
            for b in ("LOW", "MID", "HIGH"):
                mask = (q == b).values
                if mask.sum() < 100:
                    continue
                best, bs = None, -9e9
                for k, arr in grids.items():
                    tr = arr[mask & is_tr]
                    tr = tr[~np.isnan(tr)]
                    if len(tr) < MIN_TRAIN or tr.mean() <= 0:
                        continue
                    sc = tr.mean()
                    if sc > bs:
                        bs, best = sc, k
                if best is None:
                    print(f"    {b:<7} {'— no profitable rule on train —':<17}")
                    continue
                ei, xo = best
                arr = grids[best]
                te_m = mask & ~is_tr
                te = arr[te_m]
                mon = sub["mo"].values[te_m]
                te_ok = te[~np.isnan(te)]
                if len(te_ok) < 30:
                    continue
                # placebo: random dates, same hold, same stocks, same test window
                random.seed(SEED)
                pl = []
                pool = sub.loc[te_m, "symbol"].unique()
                lo, hi = sub.loc[te_m, "d"].min(), sub.loc[te_m, "d"].max()
                for _ in range(N_PLACEBO):
                    s = random.choice(pool)
                    p, m = prices[s].values, mkts[s]
                    idx = prices[s].index
                    a0, a1 = idx.searchsorted(lo), idx.searchsorted(hi)
                    if a1 - a0 < 30:
                        continue
                    a = random.randint(a0, a1)
                    b2 = a + (xo - ei)
                    if b2 >= len(p):
                        continue
                    pl.append((p[b2] / p[a] - 1 - COST) - (m[b2] / m[a] - 1))
                plm = float(np.mean(pl)) * 100 if pl else float("nan")
                avg = te_ok.mean() * 100
                t = tm(te, mon)
                edge = avg - plm
                good = edge > 0 and t == t and t > 2
                vd = "REAL?" if good else ("no edge" if edge <= 0 else "not significant")
                rule = f"buy T{ei:+d} sell T{xo:+d}"
                print(f"    {b:<7} {rule:<17} {avg:>+8.2f}% {(te_ok>0).mean()*100:>4.0f}% "
                      f"{t:>+7.2f} {len(te_ok):>6} │ {plm:>+7.2f}% {edge:>+7.2f}%  {vd}")
                out.append({"event_type": et, "dimension": label, "bucket": b,
                            "buy": ei, "sell": xo, "test_avg": avg,
                            "test_win": (te_ok > 0).mean() * 100, "t_month": t,
                            "n_test": len(te_ok), "placebo": plm, "edge": edge})

    if out:
        pd.DataFrame(out).to_csv(PROC / "fundamental_buckets.csv", index=False)
    print("\n" + "=" * 112)
    print("  READ: EDGE = out-of-sample return minus what random dates earned over")
    print("  the same holding period. t(month) absorbs the fact that 54 stocks share")
    print("  the same market days — the naive t would be 3-5x bigger and a lie.")
    print("  A bucket needs EDGE > 0 AND t(month) > 2 to be worth anything.")
    print(f"\n  Saved -> {PROC}\\fundamental_buckets.csv")


if __name__ == "__main__":
    main()
