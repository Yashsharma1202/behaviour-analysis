"""
optimal_rules.py
===============================================================================
PER-STOCK OPTIMAL "BUY x DAYS BEFORE / SELL y DAYS AFTER" RULE.

For every stock and every event type it searches the full entry/exit grid and
reports the best rule — twice, side by side:

  DESCRIPTIVE : entry allowed from T-8, even for events nobody could have known
                were coming. This is what the current dashboard does. It shows
                the true historical shape (including the pre-event run-up) but
                the announcement rules are NOT tradeable.

  TRADEABLE   : entry blocked until the event was actually PUBLIC.
                  results / board meetings -> from the intimation (~5 trading
                                              days ahead; measured from the feed)
                  dividends                -> declared ~3 weeks before ex-date
                  announcements            -> only AFTER the news is out
                Every rule here is one you could really have placed.

Honesty rails (all of them on):
  * announcement dates read from an_dt, NOT sort_date  (sort_date + dayfirst=True
    silently turns 2026-07-10 into 2026-10-07 — it corrupts ~40% of the dates)
  * 0.10% round-trip cost charged on every trade
  * every return is also MARKET-ADJUSTED (NIFTY move over the same days removed)
  * the rule is fitted on the OLDEST 70% of events and reported on the newest 30%
    that it has never seen
  * a PLACEBO of random dates with the same holding period is run alongside —
    a rule that cannot beat random dates is not a rule

    python optimal_rules.py                 # every stock with feeds
    python optimal_rules.py RELIANCE TCS     # specific stocks

Output: processed/optimal_rules.csv  +  a printed league table
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

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                   # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
PROC.mkdir(exist_ok=True)

COST = 0.001            # 0.10% round-trip
SPLIT = 0.70            # oldest 70% of events -> fit the rule; newest 30% -> test it
ENTRY = range(-8, 3)    # buy from 8 days before .. 2 days after the event
EXIT = range(-1, 13)    # sell up to 12 days after (must be after the entry)
MIN_EVENTS = 30         # don't even try a stock/event-type with fewer than this
MIN_TRAIN = 12          # a grid cell needs this many train events to be eligible
N_PLACEBO = 400
SEED = 7


# ── events, with the date each one became PUBLIC ────────────────────────────
def load_events(sym: str) -> pd.DataFrame:
    rows = []
    d = ROOT / sym

    f = d / "financial_results.csv"
    if f.exists():
        fr = pd.read_csv(f, dtype=str).fillna("")
        if "broadCastDate" in fr:
            fr["d"] = pd.to_datetime(fr["broadCastDate"], errors="coerce",
                                     format="mixed", dayfirst=True).dt.normalize()
            for dt in fr["d"].dropna().unique():
                # SEBI: board meetings for results need >=5 clear days' prior
                # intimation. Measured across the feeds it lands 7-8 calendar days
                # ahead, so T-7 calendar days is the earliest it was knowable.
                rows.append({"d": dt, "type": "RESULTS",
                             "known": pd.Timestamp(dt) - pd.Timedelta(days=7)})

    f = d / "board_meetings.csv"
    if f.exists():
        bm = pd.read_csv(f, dtype=str).fillna("")
        if "bm_date" in bm and "bm_timestamp" in bm:
            md = pd.to_datetime(bm["bm_date"], errors="coerce", dayfirst=True)
            td = pd.to_datetime(bm["bm_timestamp"], errors="coerce", dayfirst=True)
            for mm, tt in zip(md, td):
                if pd.isna(mm) or pd.isna(tt):
                    continue
                # filed after the close -> you can only act the NEXT day
                k = tt.normalize() + pd.Timedelta(days=1 if tt.hour >= 15 else 0)
                rows.append({"d": mm.normalize(), "type": "BOARD_MEETING", "known": k})

    f = d / "corporate_actions.csv"
    if f.exists():
        ca = pd.read_csv(f, dtype=str).fillna("")
        if "exDate" in ca and "subject" in ca:
            ca["d"] = pd.to_datetime(ca["exDate"], errors="coerce",
                                     dayfirst=True).dt.normalize()
            div = ca[ca["subject"].str.contains("Dividend", case=False, na=False)]
            for dt in div["d"].dropna().unique():
                # the dividend is declared at the results board meeting, ~3 weeks
                # before the ex-date, so the ex-date is known well in advance
                rows.append({"d": dt, "type": "DIVIDEND",
                             "known": pd.Timestamp(dt) - pd.Timedelta(days=21)})

    f = d / "announcements.csv"
    if f.exists():
        an = pd.read_csv(f, dtype=str).fillna("")
        if "an_dt" in an and "desc" in an:
            an = an[an["desc"].isin(EB.SIGNAL_DESC)]
            ts = pd.to_datetime(an["an_dt"], errors="coerce", dayfirst=True)
            for t in ts.dropna():
                # unscheduled: knowable only once it is disseminated. After-hours
                # filings can first be traded on the NEXT close.
                k = t.normalize() + pd.Timedelta(days=1 if t.hour >= 15 else 0)
                rows.append({"d": t.normalize(), "type": "ANNOUNCEMENT", "known": k})

    if not rows:
        return pd.DataFrame(columns=["d", "type", "known"])
    ev = pd.DataFrame(rows)
    ev["d"] = pd.to_datetime(ev["d"])
    # one event per (date, type) — the feeds file the same result 2-4 times
    ev = ev.sort_values("known").drop_duplicates(subset=["d", "type"], keep="first")
    return ev.sort_values("d").reset_index(drop=True)


# ── the grid, vectorised over events ────────────────────────────────────────
def grid(p, mk, pos, known_pos, tradeable):
    """-> dict[(entry, exit)] = array of market-adjusted returns, one per event."""
    n = len(p)
    out = {}
    for ei in ENTRY:
        i_in = pos + ei
        ok_in = (i_in >= 0) & (i_in < n)
        if tradeable:
            ok_in &= (i_in >= known_pos)          # ← the whole point
        for xo in EXIT:
            if xo <= ei:
                continue
            i_out = pos + xo
            ok = ok_in & (i_out >= 0) & (i_out < n)
            if ok.sum() < 8:
                continue
            # keep the array the same length as `pos` (NaN where the trade was not
            # possible) so the train/test mask stays aligned across every cell
            v = np.full(len(pos), np.nan)
            a, b = i_in[ok], i_out[ok]
            g = p[b] / p[a] - 1 - COST
            mr = mk[b] / mk[a] - 1
            v[ok] = g - mr                        # market-adjusted, as a fraction
            out[(ei, xo)] = v
    return out


def placebo(p, mk, hold, lo, hi, n=N_PLACEBO):
    """Same holding period, random dates in the test window = the null."""
    ok = [i for i in range(len(p)) if i + hold < len(p) and lo <= i <= hi]
    if len(ok) < 30:
        return float("nan")
    random.seed(SEED)
    v = []
    for i in (random.choice(ok) for _ in range(n)):
        g = p[i + hold] / p[i] - 1 - COST
        v.append(g - (mk[i + hold] / mk[i] - 1))
    return float(np.mean(v)) * 100


def tstat(v):
    v = np.asarray(v, float)
    v = v[~np.isnan(v)]
    if len(v) < 3 or v.std(ddof=1) == 0:
        return float("nan")
    return v.mean() / (v.std(ddof=1) / len(v) ** 0.5)


def fit_and_test(g_all, is_train, p, mk, pos, hold_lo, hold_hi):
    """Pick the best cell on TRAIN, then report it on the unseen TEST events."""
    best, bscore = None, -9e9
    for k, v in g_all.items():
        tr = v[is_train]
        tr = tr[~np.isnan(tr)]
        if len(tr) < MIN_TRAIN or tr.mean() <= 0:
            continue
        win = (tr > 0).mean()
        score = (win, tr.mean())                  # highest win-rate, then avg
        if score > (bscore if isinstance(bscore, tuple) else (-9e9, -9e9)):
            bscore, best = score, k
    if best is None:
        return None
    ei, xo = best
    v = g_all[best]
    tr, te = v[is_train], v[~is_train]
    tr, te = tr[~np.isnan(tr)], te[~np.isnan(te)]
    if len(te) < 5:
        return None
    pl = placebo(p, mk, xo - ei, hold_lo, hold_hi)
    return {"buy": ei, "sell": xo, "hold": xo - ei,
            "n_train": len(tr), "train_avg": tr.mean() * 100,
            "train_win": (tr > 0).mean() * 100,
            "n_test": len(te), "test_avg": te.mean() * 100,
            "test_win": (te > 0).mean() * 100, "test_t": tstat(te),
            "placebo": pl, "edge": te.mean() * 100 - pl}


def main():
    syms = [s.upper() for s in sys.argv[1:]] or EB.discover_feed_symbols()
    mkt_ix = NBT.fetch_index()

    print(f"OPTIMAL BUY/SELL RULE PER STOCK  —  {len(syms)} stocks")
    print(f"fit on oldest {SPLIT:.0%} of events · reported on the unseen newest "
          f"{1-SPLIT:.0%} · {COST*100:.2f}% cost · market-adjusted vs NIFTY")
    print("entry T-8 .. T+2 · exit up to T+12\n")

    out = []
    for i, sym in enumerate(syms, 1):
        px = EB.fetch_prices_yahoo(sym)
        if px is None or len(px) < 200:
            continue
        ev = load_events(sym)
        if ev.empty:
            continue
        mk = mkt_ix.reindex(px.index).ffill()
        p, mkv, I = px.values.astype(float), mk.values.astype(float), px.index
        if np.isnan(mkv).any():
            mkv = pd.Series(mkv).ffill().bfill().values

        for et, g in ev.groupby("type"):
            g = g.sort_values("d")
            pos = I.searchsorted(pd.to_datetime(g["d"]).values)
            kpos = I.searchsorted(pd.to_datetime(g["known"]).values)
            keep = pos < len(I)
            pos, kpos = pos[keep], kpos[keep]
            if len(pos) < MIN_EVENTS:
                continue
            cut = int(len(pos) * SPLIT)
            is_train = np.arange(len(pos)) < cut
            lo, hi = int(pos[cut]), int(pos[-1])

            for mode, trad in (("descriptive", False), ("tradeable", True)):
                gg = grid(p, mkv, pos, kpos, trad)
                if not gg:
                    continue
                r = fit_and_test(gg, is_train, p, mkv, pos, lo, hi)
                if r:
                    out.append({"symbol": sym, "event_type": et, "mode": mode, **r})
        if i % 10 == 0 or i == len(syms):
            print(f"  {i}/{len(syms)} stocks…")

    if not out:
        raise SystemExit("nothing usable")
    df = pd.DataFrame(out)
    df.to_csv(PROC / "optimal_rules.csv", index=False)

    # ── the league table ────────────────────────────────────────────────────
    for et in sorted(df["event_type"].unique()):
        for mode in ("descriptive", "tradeable"):
            s = df[(df.event_type == et) & (df["mode"] == mode)]
            if s.empty:
                continue
            print("\n" + "=" * 108)
            tag = ("← includes entries nobody could have placed"
                   if mode == "descriptive" else "← every rule is really placeable")
            print(f"  {et}  ·  {mode.upper()}   {tag}")
            print("=" * 108)
            print(f"  {'stock':<12} {'rule':<18} {'TRAIN avg':>10} {'win':>5} │ "
                  f"{'TEST avg':>9} {'win':>5} {'t':>6} {'n':>4} │ "
                  f"{'random':>8} {'EDGE':>8}")
            print("  " + "-" * 104)
            for _, r in s.sort_values("edge", ascending=False).iterrows():
                rule = f"buy T{r.buy:+d} sell T{r.sell:+d}"
                print(f"  {r.symbol:<12} {rule:<18} {r.train_avg:>+9.2f}% "
                      f"{r.train_win:>4.0f}% │ {r.test_avg:>+8.2f}% {r.test_win:>4.0f}% "
                      f"{r.test_t:>+6.2f} {r.n_test:>4.0f} │ {r.placebo:>+7.2f}% "
                      f"{r.edge:>+7.2f}%")
            beat = (s.edge > 0).sum()
            sig = ((s.edge > 0) & (s.test_t > 2)).sum()
            print(f"\n   {beat}/{len(s)} stocks beat random dates out-of-sample.  "
                  f"{sig}/{len(s)} also have t > 2.")

    print("\n" + "=" * 108)
    print("  HOW TO READ IT")
    print("   TRAIN is where the rule was chosen — it always looks good, ignore it.")
    print("   EDGE = out-of-sample return minus what random dates earned over the")
    print("   same holding period. EDGE <= 0 means the rule is worth nothing: you")
    print("   would have done as well buying on any random day.")
    print("   Only a rule with EDGE > 0 AND t > 2 is worth a second look.")
    print(f"\n  Saved -> {PROC}\\optimal_rules.csv")


if __name__ == "__main__":
    main()
