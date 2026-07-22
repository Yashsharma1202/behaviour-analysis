"""
nifty_backtest.py
===============================================================================
HONEST BACKTEST of the behaviour-analysis rules on the NIFTY 50.

Are the dashboard's numbers real, or curve-fitting? This answers it two ways:

 1. TRAIN / TEST SPLIT
    The "best rule" (buy T-N, sell T+M) is chosen using ONLY the older 70% of
    events. It is then applied to the newer 30% of events it has never seen.
    In-sample numbers always look good; the OUT-OF-SAMPLE column is the truth.

 2. MARKET-ADJUSTED RETURNS
    Every trade's return is also reported minus the NIFTY's move over the exact
    same days. If the edge disappears once the market is stripped out, the
    "profit" was just the market drifting up, not the event.

Costs of 0.10% round-trip are charged on every trade.

    python nifty_backtest.py                 # all Nifty 50 stocks with feeds
    python nifty_backtest.py RESULTS         # one event type
    python nifty_backtest.py --split 0.6     # different train/test cut

Output: printed report + processed/nifty_backtest.csv (every test trade)
===============================================================================
"""

from __future__ import annotations

import json
import ssl
import sys
import time
import urllib.request
from pathlib import Path

import pandas as pd

import event_behaviour as EB
from download_feeds import NIFTY50_FALLBACK

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                              # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
CACHE = PROC / "price_cache"
PROC.mkdir(exist_ok=True)

COST = 0.001                 # 0.10% round-trip
SPLIT = 0.70                 # first 70% of events -> train, rest -> test
NB, NA = EB.N_BEFORE, EB.N_AFTER
MIN_TRAIN = 15               # need at least this many train events to fit a rule

_H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
_S = ssl.create_default_context()
_S.check_hostname = False
_S.verify_mode = ssl.CERT_NONE


# ---------------------------------------------------------------------------
def fetch_index() -> pd.Series:
    """NIFTY 50 index (^NSEI) daily close, cached."""
    cache = CACHE / "_NSEI.csv"
    if cache.exists():
        return pd.read_csv(cache, parse_dates=["date"]).set_index("date")["adj"].sort_index()
    url = ("https://query1.finance.yahoo.com/v8/finance/chart/%5ENSEI"
           f"?period1=1167609600&period2={int(time.time())}&interval=1d")
    req = urllib.request.Request(url, headers=_H)
    with urllib.request.urlopen(req, timeout=30, context=_S) as r:
        d = json.load(r)
    res = d["chart"]["result"][0]
    ind = res["indicators"]
    adj = ind.get("adjclose", [{}])[0].get("adjclose") or ind["quote"][0]["close"]
    s = pd.Series(adj, index=pd.to_datetime(res["timestamp"], unit="s").normalize())
    s = s.dropna().sort_index()
    s.rename("adj").rename_axis("date").to_frame().to_csv(cache)
    return s


def trades_for_rule(px, mkt, dates, db, da) -> list[dict]:
    """Apply 'buy db days before / sell da days after' to a list of event dates."""
    idx = px.index
    out = []
    for d in dates:
        pos = idx.searchsorted(pd.Timestamp(d))
        i_in, i_out = pos - db, pos + da
        if i_in < 0 or i_out >= len(idx):
            continue
        p_in, p_out = float(px.iloc[i_in]), float(px.iloc[i_out])
        if p_in <= 0:
            continue
        gross = p_out / p_in - 1
        d_in, d_out = idx[i_in], idx[i_out]
        m_in, m_out = float(mkt.get(d_in, float("nan"))), float(mkt.get(d_out, float("nan")))
        mret = (m_out / m_in - 1) if (m_in and m_in == m_in and m_out == m_out) else float("nan")
        out.append({
            "event_date": pd.Timestamp(d).date(),
            "entry": d_in.date(), "exit": d_out.date(),
            "entry_price": round(p_in, 2), "exit_price": round(p_out, 2),
            "ret_pct": (gross - COST) * 100,                 # net of cost
            "mkt_pct": mret * 100 if mret == mret else float("nan"),
            "abn_pct": (gross - COST - mret) * 100 if mret == mret else float("nan"),
        })
    return out


def fit_best(train: pd.DataFrame) -> tuple[int, int] | None:
    """Pick (days_before, days_after) on TRAIN only — same logic as the dashboard:
    highest win-rate among rules with a positive average."""
    if train.empty:
        return None
    g = train.groupby(["db", "da"]).agg(n=("ret", "size"),
                                        avg=("ret", "mean"),
                                        win=("ret", lambda s: (s > 0).mean() * 100)).reset_index()
    g = g[g["n"] >= MIN_TRAIN]
    if g.empty:
        return None
    pos = g[g["avg"] > 0]
    pool = pos if not pos.empty else g
    b = pool.sort_values(["win", "avg"], ascending=False).iloc[0]
    return int(b["db"]), int(b["da"])


def stats(d: pd.DataFrame, col: str) -> dict:
    """Per-trade stats.

    NOTE: we deliberately do NOT compound a 'total return' — these trades overlap
    (many stocks, same calendar days), so chaining them is meaningless.

    Two t-stats are reported:
      t        : naive, assumes every trade is independent  -> INFLATED, because
                 trades on the same days across 50 stocks are correlated.
      t(month) : average the trades within each calendar month first, then t-test
                 the monthly series. This absorbs the cross-sectional correlation
                 and is the number to trust.
    """
    s = d[col].dropna()
    n = len(s)
    if n == 0:
        return {"n": 0}
    sd = s.std(ddof=1)
    t = s.mean() / (sd / n ** 0.5) if sd else float("nan")

    dd = d.dropna(subset=[col]).copy()
    dd["m"] = pd.to_datetime(dd["entry"]).dt.to_period("M")
    mm = dd.groupby("m")[col].mean()
    nm = len(mm)
    tm = float("nan")
    if nm > 2:
        sdm = mm.std(ddof=1)
        if sdm and sdm == sdm:
            tm = mm.mean() / (sdm / nm ** 0.5)
    return {"n": n, "win": (s > 0).mean() * 100, "avg": s.mean(),
            "med": s.median(), "t": t, "tm": tm, "nm": nm}


def row(label, st):
    if not st.get("n"):
        return f"  {label:<26} {'—':>5}"
    tm = f"{st['tm']:+.2f}" if st["tm"] == st["tm"] else "n/a"
    return (f"  {label:<26} {st['n']:>6} {st['win']:>7.0f}% {st['avg']:>+9.2f}% "
            f"{st['med']:>+9.2f}% {st['t']:>+9.2f} {tm:>11} {st['nm']:>8}")


# ---------------------------------------------------------------------------
def main():
    args = [a for a in sys.argv[1:]]
    global SPLIT
    if "--split" in args:
        i = args.index("--split")
        SPLIT = float(args[i + 1])
        del args[i:i + 2]
    only = [a.upper() for a in args] or None

    nifty = sorted(set(NIFTY50_FALLBACK) & set(EB.discover_feed_symbols()))
    if not nifty:
        raise SystemExit("No Nifty 50 stocks with downloaded feeds. "
                         "Run: python download_feeds.py nifty50")
    print(f"NIFTY 50 BACKTEST — {len(nifty)} stocks with feeds")
    print(f"train = oldest {SPLIT:.0%} of events · test = newest {1-SPLIT:.0%} (unseen)")
    print(f"cost {COST*100:.2f}%/trade · rule searched over T-1..T-{NB} / T+1..T+{NA}\n")

    mkt = fetch_index()

    # ---- gather every (event, db, da) return, tagged train/test --------------
    recs, all_events = [], []
    for i, sym in enumerate(nifty, 1):
        px = EB.fetch_prices_yahoo(sym)
        if px is None or len(px) < NB + NA + 30:
            continue
        m = mkt.reindex(px.index).ffill()
        ev = EB.load_events_from_feeds([sym])
        for et, grp in ev.groupby("event_type"):
            if only and et not in only:
                continue
            dates = sorted(grp["event_date"].drop_duplicates())
            if len(dates) < 20:
                continue
            cut = dates[int(len(dates) * SPLIT)]
            for d in dates:
                pos = px.index.searchsorted(pd.Timestamp(d))
                for db in range(1, NB + 1):
                    i_in = pos - db
                    if i_in < 0:
                        continue
                    p_in = float(px.iloc[i_in])
                    if p_in <= 0:
                        continue
                    d_in = px.index[i_in]
                    for da in range(1, NA + 1):
                        i_out = pos + da
                        if i_out >= len(px):
                            continue
                        p_out = float(px.iloc[i_out])
                        d_out = px.index[i_out]
                        g = p_out / p_in - 1
                        m_in, m_out = m.get(d_in, float("nan")), m.get(d_out, float("nan"))
                        mr = (m_out / m_in - 1) if (m_in == m_in and m_out == m_out) else float("nan")
                        recs.append({
                            "symbol": sym, "event_type": et, "event_date": d,
                            "db": db, "da": da,
                            "ret": (g - COST) * 100,
                            "abn": (g - COST - mr) * 100 if mr == mr else float("nan"),
                            "split": "train" if d < cut else "test",
                            "entry": d_in.date(), "exit": d_out.date(),
                            "entry_price": round(p_in, 2), "exit_price": round(p_out, 2),
                        })
        if i % 10 == 0 or i == len(nifty):
            print(f"  loaded {i}/{len(nifty)} stocks…")

    if not recs:
        raise SystemExit("No usable events/prices.")
    df = pd.DataFrame(recs)

    # ---- per event type: fit on train, evaluate on test ---------------------
    print("\n" + "=" * 104)
    print(f"  {'':26} {'n':>6} {'win%':>8} {'avg':>10} {'median':>10} "
          f"{'t(naive)':>9} {'t(month)':>11} {'months':>8}")
    print("=" * 104)

    kept = []
    verdicts = []
    for et in sorted(df["event_type"].unique()):
        sub = df[df["event_type"] == et]
        tr = sub[sub["split"] == "train"]
        best = fit_best(tr)
        if not best:
            continue
        db, da = best
        tr_r = tr[(tr.db == db) & (tr.da == da)]
        te_r = sub[(sub["split"] == "test") & (sub.db == db) & (sub.da == da)]
        if te_r.empty:
            continue

        print(f"\n  {et}  —  rule fitted on train: BUY {db}d before, SELL {da}d after")
        print("  " + "-" * 100)
        print(row("IN-SAMPLE  raw", stats(tr_r, "ret")))
        print(row("IN-SAMPLE  mkt-adjusted", stats(tr_r, "abn")))
        print(row("OUT-OF-SAMPLE  raw", stats(te_r, "ret")))
        print(row("OUT-OF-SAMPLE  mkt-adj", stats(te_r, "abn")))

        oos = stats(te_r, "abn")
        oos_raw = stats(te_r, "ret")
        verdicts.append((et, db, da, oos_raw, oos))
        k = te_r.copy()
        k["rule"] = f"buy-{db}d / sell+{da}d"
        kept.append(k)

    if kept:
        out = pd.concat(kept)[["symbol", "event_type", "rule", "event_date", "entry",
                               "exit", "entry_price", "exit_price", "ret", "abn"]]
        out.columns = ["symbol", "event_type", "rule", "event_date", "entry_date",
                       "exit_date", "entry_price", "exit_price", "net_return_%",
                       "market_adjusted_%"]
        out.to_csv(PROC / "nifty_backtest.csv", index=False)

    # ---- verdict ------------------------------------------------------------
    print("\n" + "=" * 104)
    print("  VERDICT — does the edge survive on events the rule never saw?")
    print("=" * 104)
    for et, db, da, raw, adj in verdicts:
        tm = adj.get("tm", float("nan"))          # month-clustered t = the honest one
        sig = tm == tm and abs(tm) > 2
        real = adj.get("n") and adj["avg"] > 0 and sig
        drift = raw.get("n") and raw["avg"] > 0 and adj.get("avg", 0) <= 0
        if real:
            v = "REAL EDGE — still profitable out-of-sample after removing the market"
        elif drift:
            v = "JUST MARKET DRIFT — profit vanishes once the market is stripped out"
        elif raw.get("n") and raw["avg"] <= 0:
            v = "NO EDGE — loses money on unseen events"
        else:
            v = "WEAK / NOT SIGNIFICANT — edge too small to distinguish from luck"
        tms = f"{tm:+.2f}" if tm == tm else "n/a"
        print(f"   {et:<18} buy {db}d/sell {da}d  →  "
              f"raw {raw.get('avg', 0):+.2f}%, mkt-adj {adj.get('avg', 0):+.2f}%  "
              f"(t-month={tms}, n={adj.get('n', 0)}, months={adj.get('nm', 0)})")
        print(f"   {'':18} {v}")
    print("\n  HOW TO READ IT")
    print("   * IN-SAMPLE is always flattering — the rule was chosen on that data.")
    print("   * Trust the OUT-OF-SAMPLE mkt-adjusted row: unseen events, market removed.")
    print("   * Use t(month), not t(naive). Trades on 50 stocks share the same market")
    print("     days, so they are NOT independent — t(naive) is badly inflated.")
    print("   * |t(month)| > 2 ~ 95% confidence the average really isn't zero.")
    print("   * No 'total return' is shown on purpose: these trades OVERLAP, so")
    print("     compounding them would be meaningless.")
    print(f"\n  Test trades saved -> {PROC}\\nifty_backtest.csv")


if __name__ == "__main__":
    main()
