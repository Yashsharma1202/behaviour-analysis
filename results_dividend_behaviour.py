"""
results_dividend_behaviour.py
===============================================================================
WHAT DOES A STOCK ACTUALLY DO AROUND ITS QUARTERLY RESULTS AND ITS DIVIDEND?

This is a DESCRIPTIVE study, not a strategy search. The question is not "what
should I trade" but "what is the historical behaviour" — so nothing here is
selected by searching a grid for the best cell. Every number is a plain
population statistic over all events, reported with a confidence interval.

PART A — QUARTERLY RESULTS
    * the average price PATH from T-20 to T+20, raw and market-adjusted
    * the RESULT-DAY REACTION: the move from the last close before the filing
      to the first close after it. This is the actual news reaction, and it is
      measured from the first tradeable close AFTER the announcement timestamp,
      so an after-hours filing is never credited with a move nobody could act on
    * how often the stock rises, with a Wilson interval
    * the VOLATILITY KICK: realised vol in the 5 days after vs the 20 days
      before — how much a result actually shakes the stock

PART B — DIVIDEND EX-DATE DISCOUNT AND RECOVERY
    On the ex-date the stock is marked down because the cash is leaving:

        Day before ex (cum-dividend) : close 2900   <- you still get the dividend
        Ex-date                      : close 2894   <- ~Rs 5.5 knocked off

    For every dividend we measure:
      1. the actual DROP on the ex-date
      2. the DROP RATIO = actual drop / dividend amount. Theory says 1.0; the
         literature consistently finds less than 1 because of dividend taxes
      3. RECOVERY TIME — trading days until the close gets back to the
         cum-dividend close, reported both raw and market-adjusted (did the
         stock recover on its own, or did the market simply carry it up?)

WHY RAW PRICES HERE
-------------------
Yahoo's `adjclose` is dividend-adjusted: it retroactively removes the ex-date
drop, so a recovery study run on it would find nothing to recover from. This
module uses `quote.close`, which IS split/bonus-adjusted (no fake -50% from a
1:1 bonus) but NOT dividend-adjusted — exactly what is needed. Cached
separately from the adjusted series used elsewhere in the project.

RUN
    python results_dividend_behaviour.py                # all symbols with feeds
    python results_dividend_behaviour.py RELIANCE TCS
    python results_dividend_behaviour.py --max-wait 120 # recovery search window
===============================================================================
"""

from __future__ import annotations

import argparse
import json
import math
import re
import ssl
import sys
import time
import urllib.request
import warnings
from pathlib import Path
from urllib.parse import quote as urlquote

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # noqa: BLE001
    pass

import event_behaviour as EB

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
RAW_CACHE = PROC / "raw_price_cache"
RAW_CACHE.mkdir(parents=True, exist_ok=True)

PRE, POST = 20, 20            # path window around results
MAX_WAIT = 120                # trading days to look for a dividend recovery

_H = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
_S = ssl.create_default_context()


# ---------------------------------------------------------------------------
# Raw (dividend-UNadjusted) prices
# ---------------------------------------------------------------------------
def fetch_raw_prices(symbol: str) -> pd.DataFrame | None:
    """Daily OHLCV using Yahoo's `quote` block, which is split/bonus adjusted but
    NOT dividend adjusted. The ex-date drop survives, which is the whole point.

    Volume is fetched for technicals.py (OBV, volume z-score). Caches written
    before volume was added lack the column, so those are re-fetched once.
    """
    cache = RAW_CACHE / f"{symbol}.csv"
    if cache.exists():
        try:
            df = pd.read_csv(cache, parse_dates=["date"]).set_index("date").sort_index()
            if "volume" in df.columns:
                return df
        except Exception:  # noqa: BLE001
            pass
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/"
           f"{urlquote(symbol, safe='')}.NS"
           f"?period1=1104537600&period2={int(time.time())}&interval=1d")
    try:
        req = urllib.request.Request(url, headers=_H)
        with urllib.request.urlopen(req, timeout=30, context=_S) as r:
            d = json.load(r)
        res = d["chart"]["result"][0]
        q = res["indicators"]["quote"][0]
        df = pd.DataFrame({
            "open": q.get("open"), "high": q.get("high"),
            "low": q.get("low"), "close": q.get("close"),
            "volume": q.get("volume"),
        }, index=pd.to_datetime(res["timestamp"], unit="s").normalize())
        df.index.name = "date"
        df = df.dropna(subset=["close"]).sort_index()
        df.to_csv(cache)
        time.sleep(0.3)
        return df
    except Exception as e:  # noqa: BLE001
        print(f"    ! raw price fetch failed for {symbol}: {type(e).__name__}")
        return None


# ---------------------------------------------------------------------------
# Event loading
# ---------------------------------------------------------------------------
DIV_RE = re.compile(r"(?:rs\.?|re\.?|inr|₹)\s*([0-9]+(?:\.[0-9]+)?)", re.I)
BONUS_RE = re.compile(r"bonus\s*(\d+)\s*:\s*(\d+)", re.I)
SPLIT_RE = re.compile(
    r"(?:face\s*value\s*)?spl?it.*?(?:rs\.?|re\.?)\s*([0-9]+(?:\.[0-9]+)?)"
    r".*?to\s*(?:rs\.?|re\.?)\s*([0-9]+(?:\.[0-9]+)?)", re.I | re.S)


def split_bonus_factors(symbols):
    """Cumulative share-multiplication factor per symbol, by ex-date.

    Yahoo's `quote.close` is already split/bonus adjusted, but NSE reports the
    dividend amount in the share terms of the day. A Rs 51 dividend paid before
    a 10:1 split looks like a 20% yield against a split-adjusted price unless
    it is divided by 10. TATASTEEL is exactly this case.

    Returns {symbol: [(ex_date, factor), ...]}.
    """
    out = {}
    for s in symbols:
        p = ROOT / s / "corporate_actions.csv"
        if not p.exists():
            continue
        d = pd.read_csv(p, dtype=str)
        if "subject" not in d.columns or "exDate" not in d.columns:
            continue
        evs = []
        for _, r in d.iterrows():
            sub = str(r.get("subject") or "")
            low = sub.lower()
            if "debenture" in low or "rights" in low:
                continue          # not a share-count change Yahoo adjusts for
            f = 1.0
            mb = BONUS_RE.search(sub)
            if mb:
                a, b = float(mb.group(1)), float(mb.group(2))
                if b > 0:
                    f *= (a + b) / b
            ms = SPLIT_RE.search(sub)
            if ms:
                old, new = float(ms.group(1)), float(ms.group(2))
                if new > 0 and old > new:
                    f *= old / new
            if f != 1.0:
                ex = pd.to_datetime(r.get("exDate"), errors="coerce", dayfirst=True)
                if pd.notna(ex):
                    evs.append((ex, f))
        if evs:
            out[s] = sorted(evs)
    return out


def adjust_factor(sym, ex_date, factors):
    """Product of every split/bonus that happened AFTER this dividend."""
    f = 1.0
    for d0, mult in factors.get(sym, []):
        if d0 > ex_date:
            f *= mult
    return f


def load_dividends(symbols):
    """Ex-dates with the per-share amount parsed out of the subject line.

    Bonus / split / demerger rows are excluded: Yahoo's `quote.close` is already
    adjusted for those, so their 'drop' is an artefact, not a real markdown.
    """
    rows = []
    for s in symbols:
        p = ROOT / s / "corporate_actions.csv"
        if not p.exists():
            continue
        d = pd.read_csv(p, dtype=str)
        if "exDate" not in d.columns or "subject" not in d.columns:
            continue
        sub = d["subject"].fillna("")
        keep = sub.str.contains("dividend", case=False, na=False)
        # exclude anything bundled with a bonus/split on the same line
        keep &= ~sub.str.contains(r"bonus|split|demerger|spilt", case=False, na=False)
        d = d[keep]
        if d.empty:
            continue
        ex = pd.to_datetime(d["exDate"], errors="coerce", dayfirst=True)
        amt = d["subject"].str.extract(DIV_RE)[0].astype(float)
        kind = np.where(d["subject"].str.contains("interim", case=False, na=False),
                        "Interim", np.where(
            d["subject"].str.contains("final", case=False, na=False), "Final", "Other"))
        rows.append(pd.DataFrame({"symbol": s, "ex_date": ex, "amount": amt,
                                  "kind": kind, "subject": d["subject"].values}))
    if not rows:
        return pd.DataFrame()
    dv = pd.concat(rows, ignore_index=True).dropna(subset=["ex_date", "amount"])
    dv = dv[dv["amount"] > 0]
    # multiple dividends sharing an ex-date are one economic markdown
    dv = (dv.groupby(["symbol", "ex_date"], as_index=False)
            .agg(amount=("amount", "sum"), kind=("kind", "first"),
                 subject=("subject", " + ".join)))

    # put the dividend into the same share terms as the split-adjusted price
    factors = split_bonus_factors(symbols)
    dv["adj_factor"] = [adjust_factor(s, d0, factors)
                        for s, d0 in zip(dv["symbol"], dv["ex_date"])]
    dv["amount_raw"] = dv["amount"]
    dv["amount"] = dv["amount"] / dv["adj_factor"]
    return dv.sort_values(["symbol", "ex_date"]).reset_index(drop=True)


def load_results(symbols):
    """Quarterly result filings, one row per (symbol, period)."""
    rows = []
    for s in symbols:
        p = ROOT / s / "financial_results.csv"
        if not p.exists():
            continue
        d = pd.read_csv(p, dtype=str)
        if "broadCastDate" not in d.columns:
            continue
        ts = pd.to_datetime(d["broadCastDate"], format="%d-%b-%Y %H:%M:%S",
                            errors="coerce")
        ts = ts.fillna(pd.to_datetime(d["broadCastDate"], errors="coerce",
                                      dayfirst=True))
        per = (pd.to_datetime(d["toDate"], errors="coerce", dayfirst=True)
               if "toDate" in d.columns else pd.NaT)
        rows.append(pd.DataFrame({
            "symbol": s, "ts": ts, "period_end": per,
            "relating": d.get("relatingTo", pd.Series("", index=d.index)),
            "period": d.get("period", pd.Series("", index=d.index)),
        }))
    if not rows:
        return pd.DataFrame()
    r = pd.concat(rows, ignore_index=True).dropna(subset=["ts"])
    r = r[r["period"].fillna("").str.contains("Quarter", case=False, na=False)
          | r["period"].isna() | (r["period"] == "")]
    r = (r.sort_values("ts")
           .groupby(["symbol", "period_end"], as_index=False, dropna=False).first())
    return r.sort_values(["symbol", "ts"]).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Stats helpers
# ---------------------------------------------------------------------------
def wilson(wins, n, conf=0.95):
    if n == 0:
        return (np.nan, np.nan, np.nan)
    z = 1.959963984540054 if abs(conf - 0.95) < 1e-9 else 1.96
    p = wins / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return (p * 100, max(0.0, c - h) * 100, min(1.0, c + h) * 100)


def ci_mean(x, conf=0.95):
    x = np.asarray(x, dtype=float)
    x = x[np.isfinite(x)]
    if len(x) < 2:
        return (np.nan, np.nan, np.nan)
    m = x.mean()
    se = x.std(ddof=1) / math.sqrt(len(x))
    return (m, m - 1.96 * se, m + 1.96 * se)


# ---------------------------------------------------------------------------
# PART A — results behaviour
# ---------------------------------------------------------------------------
def results_behaviour(res, px_map, mkt):
    recs, paths = [], []
    for _, e in res.iterrows():
        px = px_map.get(e["symbol"])
        if px is None:
            continue
        idx = px.index
        ts = e["ts"]
        day = ts.normalize()
        # first tradeable close AFTER the filing instant
        pos = idx.searchsorted(day)
        if pos >= len(idx):
            continue
        filed_intraday = (idx[pos] == day) and (ts.hour + ts.minute / 60 < 15.5)
        p_after = pos if filed_intraday else pos + 1
        p_before = p_after - 1
        if p_before - PRE < 0 or p_after + POST >= len(idx):
            continue

        c = px["close"].to_numpy()
        before, after = c[p_before], c[p_after]
        if not (before > 0 and after > 0):
            continue
        react = (after / before - 1) * 100

        m = mkt.reindex(idx).ffill().to_numpy()
        mreact = ((m[p_after] / m[p_before] - 1) * 100
                  if np.isfinite(m[p_before]) and m[p_before] > 0 else np.nan)

        # Volatility kick. The post window must INCLUDE the reaction day --
        # that is the whole point -- and both use ddof=1, since a 5-sample
        # population std is biased low and made results look calming.
        pre_r = np.diff(c[p_before - 20:p_before + 1]) / c[p_before - 20:p_before]
        post_r = np.diff(c[p_before:p_after + 5]) / c[p_before:p_after + 4]
        v_pre = np.nanstd(pre_r, ddof=1) * 100
        v_post = np.nanstd(post_r, ddof=1) * 100 if len(post_r) > 1 else np.nan

        recs.append({
            "symbol": e["symbol"], "ts": ts, "period_end": e["period_end"],
            "relating": e["relating"],
            "close_before": before, "close_after": after,
            "reaction_pct": react,
            "reaction_abn_pct": react - mreact if np.isfinite(mreact) else np.nan,
            "vol_pre": v_pre, "vol_post": v_post,
            "vol_kick": v_post / v_pre if v_pre > 0 else np.nan,
        })
        seg = c[p_before - PRE: p_before + POST + 2]
        if len(seg) == PRE + POST + 2:
            paths.append(seg / seg[PRE] - 1.0)
    return pd.DataFrame(recs), (np.array(paths) if paths else np.empty((0, 0)))


# ---------------------------------------------------------------------------
# PART B — dividend discount and recovery
# ---------------------------------------------------------------------------
def dividend_behaviour(dv, px_map, mkt, max_wait=MAX_WAIT):
    out = []
    for _, e in dv.iterrows():
        px = px_map.get(e["symbol"])
        if px is None:
            continue
        idx = px.index
        ex = pd.Timestamp(e["ex_date"]).normalize()
        pos = idx.searchsorted(ex)
        if pos >= len(idx) or pos == 0:
            continue
        # the ex-date must be a real trading day for the drop to be measurable
        if idx[pos] != ex:
            continue
        c = px["close"].to_numpy()
        o = px["open"].to_numpy() if "open" in px.columns else c
        cum = c[pos - 1]                      # last cum-dividend close
        exc = c[pos]                          # ex-date close
        exo = o[pos] if np.isfinite(o[pos]) else exc
        if not (cum > 0 and exc > 0):
            continue

        amt = float(e["amount"])
        yld = amt / cum * 100
        drop = cum - exc
        drop_open = cum - exo

        m = mkt.reindex(idx).ffill().to_numpy()
        mret = (m[pos] / m[pos - 1] - 1) if (np.isfinite(m[pos - 1]) and m[pos - 1] > 0) else np.nan
        # what the stock "should" have done from the market alone
        drop_abn = drop - (-mret * cum if np.isfinite(mret) else 0.0)

        # ---- recovery: first close back at or above the cum-dividend close --
        hi = min(pos + max_wait, len(idx) - 1)
        fwd = c[pos:hi + 1]
        rec = np.where(fwd >= cum)[0]
        rec_days = int(rec[0]) if len(rec) else np.nan       # 0 = same day

        # market-adjusted recovery: strip the market's move out of the path so
        # we can tell a genuine recovery from one the whole market delivered
        if np.isfinite(mret):
            mm = m[pos:hi + 1]
            with np.errstate(invalid="ignore", divide="ignore"):
                adj = fwd / (mm / m[pos - 1])
            reca = np.where(adj >= cum)[0]
            rec_days_abn = int(reca[0]) if len(reca) else np.nan
        else:
            rec_days_abn = np.nan

        out.append({
            "symbol": e["symbol"], "ex_date": ex, "kind": e["kind"],
            "amount": amt, "yield_pct": yld,
            "cum_close": cum, "ex_close": exc, "ex_open": exo,
            "drop_rs": drop, "drop_pct": drop / cum * 100,
            "drop_open_rs": drop_open,
            "drop_ratio": drop / amt if amt > 0 else np.nan,
            "drop_ratio_open": drop_open / amt if amt > 0 else np.nan,
            "drop_abn_pct": drop_abn / cum * 100,
            "recovery_days": rec_days,
            "recovery_days_abn": rec_days_abn,
            "recovered": int(np.isfinite(rec_days)),
            "subject": e["subject"],
        })
    return pd.DataFrame(out)


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("symbols", nargs="*")
    ap.add_argument("--max-wait", type=int, default=MAX_WAIT)
    args = ap.parse_args()

    symbols = args.symbols or sorted(
        set(EB.discover_feed_symbols()) &
        set(__import__("download_feeds").NIFTY50_FALLBACK))

    print("=" * 78)
    print("  RESULTS & DIVIDEND BEHAVIOUR — descriptive study")
    print(f"  {len(symbols)} symbols | raw (dividend-UNadjusted) prices")
    print("=" * 78)

    print("\nFetching raw prices ...")
    px_map = {}
    for i, s in enumerate(symbols, 1):
        d = fetch_raw_prices(s)
        if d is not None and len(d) > 300:
            px_map[s] = d
    print(f"  {len(px_map)}/{len(symbols)} symbols with raw price history")

    nif = EB.fetch_prices_yahoo("^NSEI") if False else None
    import nifty_backtest as NBT
    mkt = NBT.fetch_index()

    # ================= PART A =============================================
    print("\nLoading quarterly results ...")
    res = load_results(list(px_map))
    print(f"  {len(res):,} result filings")
    rb, paths = results_behaviour(res, px_map, mkt)
    print(f"  {len(rb):,} with full price windows")

    if len(rb):
        print("\n" + "=" * 78)
        print("  PART A — BEHAVIOUR AROUND QUARTERLY RESULTS")
        print("=" * 78)
        m, lo, hi = ci_mean(rb["reaction_pct"])
        ma, loa, hia = ci_mean(rb["reaction_abn_pct"])
        up = int((rb["reaction_pct"] > 0).sum())
        w, wl, wh = wilson(up, len(rb))
        print(f"\n  RESULT-DAY REACTION (last close before -> first close after)")
        print(f"    mean          {m:+.3f}%   95% CI [{lo:+.3f}, {hi:+.3f}]")
        print(f"    market-adj    {ma:+.3f}%   95% CI [{loa:+.3f}, {hia:+.3f}]")
        print(f"    median        {rb['reaction_pct'].median():+.3f}%")
        print(f"    stock rose    {w:.1f}% of the time   "
              f"95% CI [{wl:.1f}%, {wh:.1f}%]"
              + ("   <- straddles 50%" if wl < 50 < wh else ""))
        print(f"    |move| median {rb['reaction_pct'].abs().median():.2f}%   "
              f"p90 {rb['reaction_pct'].abs().quantile(.9):.2f}%")
        print(f"    biggest up    {rb['reaction_pct'].max():+.2f}%   "
              f"biggest down {rb['reaction_pct'].min():+.2f}%")

        vk = rb["vol_kick"].replace([np.inf, -np.inf], np.nan).dropna()
        print(f"\n  VOLATILITY KICK (5d after / 20d before)")
        print(f"    median {vk.median():.2f}x   mean {vk.mean():.2f}x   "
              f"share above 1.0: {(vk > 1).mean()*100:.0f}%")
        print(f"    -> a result multiplies short-term volatility by about "
              f"{vk.median():.2f}x")

        if len(paths):
            avg = paths.mean(0) * 100
            print(f"\n  AVERAGE PRICE PATH around the result "
                  f"(T-{PRE} .. T+{POST}, indexed to the last close before)")
            days = list(range(-PRE, POST + 2))
            for d0 in (-20, -10, -5, -3, -1, 0, 1, 2, 3, 5, 10, 20):
                j = days.index(d0)
                bar = int(abs(avg[j]) * 8)
                side = " " * 22 + "|"
                mark = ("#" * min(bar, 20)).rjust(21) + "|" if avg[j] < 0 \
                    else side + "#" * min(bar, 20)
                print(f"    T{d0:+3d}  {avg[j]:+6.2f}%  {mark}")

        by = (rb.groupby("symbol")
                .agg(n=("reaction_pct", "size"),
                     mean_react=("reaction_pct", "mean"),
                     up_rate=("reaction_pct", lambda s: (s > 0).mean() * 100),
                     abs_move=("reaction_pct", lambda s: s.abs().median()))
                .sort_values("abs_move", ascending=False))
        print(f"\n  MOST / LEAST REACTIVE STOCKS (median |result-day move|)")
        for lab, sub in (("most", by.head(5)), ("least", by.tail(5))):
            for sym, r in sub.iterrows():
                print(f"    {lab:>5}  {sym:<12} n={int(r['n']):>3}  "
                      f"|move| {r['abs_move']:.2f}%  up {r['up_rate']:.0f}%  "
                      f"mean {r['mean_react']:+.2f}%")
        rb.to_csv(PROC / "results_behaviour.csv", index=False)
        by.to_csv(PROC / "results_behaviour_by_symbol.csv")

    # ================= PART B =============================================
    print("\nLoading dividends ...")
    dv = load_dividends(list(px_map))
    print(f"  {len(dv):,} dividend ex-dates")
    db = dividend_behaviour(dv, px_map, mkt, args.max_wait)
    print(f"  {len(db):,} with usable prices around the ex-date")

    if len(db):
        print("\n" + "=" * 78)
        print("  PART B — DIVIDEND EX-DATE DISCOUNT AND RECOVERY")
        print("=" * 78)
        print(f"\n  THE DISCOUNT")
        m, lo, hi = ci_mean(db["drop_pct"])
        print(f"    dividend yield   median {db['yield_pct'].median():.2f}%  "
              f"(mean {db['yield_pct'].mean():.2f}%)")
        print(f"    ex-date drop     mean {m:+.3f}%   95% CI [{lo:+.3f}, {hi:+.3f}]")
        print(f"    market-adjusted  mean {db['drop_abn_pct'].mean():+.3f}%")
        print(f"\n    DROP RATIO = actual drop / dividend amount  (theory: 1.0)")
        print(f"      Below 1.0 means the market marks the stock down by LESS")
        print(f"      than the dividend — the classic dividend-tax effect.")
        print(f"\n      {'sample':<22} {'n':>5} {'mean':>8} {'median':>8} {'95% CI':>20}")
        print("      " + "-" * 66)
        for lab, sub in (("all dividends", db),
                         ("yield > 0.5%", db[db.yield_pct > 0.5]),
                         ("yield > 1%", db[db.yield_pct > 1.0]),
                         ("yield > 2%", db[db.yield_pct > 2.0])):
            if len(sub) < 10:
                continue
            r, rl, rh = ci_mean(sub["drop_ratio"])
            print(f"      {lab:<22} {len(sub):>5} {r:>8.3f} "
                  f"{sub['drop_ratio'].median():>8.3f}   [{rl:>6.2f}, {rh:>6.2f}]")
        print(f"\n      Low-yield dividends make the ratio meaningless: a 0.3%")
        print(f"      dividend is swamped by the ~1.5% daily noise, so the")
        print(f"      denominator is tiny and the ratio explodes. Trust the")
        print(f"      yield > 1% rows.")

        fell = int((db["drop_rs"] > 0).sum())
        w, wl, wh = wilson(fell, len(db))
        print(f"\n    stock actually fell on the ex-date: {w:.1f}% of the time "
              f"95% CI [{wl:.1f}%, {wh:.1f}%]")

        print(f"\n  THE RECOVERY  (trading days back to the cum-dividend close)")
        rec = db["recovery_days"].dropna()
        reca = db["recovery_days_abn"].dropna()
        print(f"    recovered within {args.max_wait}d : "
              f"{db['recovered'].mean()*100:.1f}% of dividends "
              f"({int(db['recovered'].sum()):,} of {len(db):,})")
        if len(rec):
            print(f"    median recovery   {rec.median():.0f} trading days")
            print(f"    mean recovery     {rec.mean():.1f} days")
            print(f"    market-adjusted   median {reca.median():.0f} days"
                  if len(reca) else "")
            print(f"\n    {'within':>10} {'raw':>10} {'mkt-adj':>10}")
            print("    " + "-" * 32)
            for d0 in (0, 1, 2, 3, 5, 10, 20, 40, 60, args.max_wait):
                a = (rec <= d0).sum() / len(db) * 100
                b = (reca <= d0).sum() / len(db) * 100 if len(reca) else np.nan
                print(f"    {d0:>8}d  {a:>9.1f}% {b:>9.1f}%")

        print(f"\n  DOES A BIGGER DIVIDEND TAKE LONGER TO RECOVER?")
        db["yq"] = pd.qcut(db["yield_pct"], 4, labels=["Q1 low", "Q2", "Q3", "Q4 high"])
        g = db.groupby("yq", observed=True).agg(
            n=("yield_pct", "size"), yld=("yield_pct", "median"),
            drop=("drop_pct", "mean"), ratio=("drop_ratio", "median"),
            recov=("recovered", "mean"),
            med_days=("recovery_days", "median"))
        print(f"    {'bucket':<9} {'n':>5} {'yield':>7} {'drop':>8} "
              f"{'ratio':>7} {'recov%':>8} {'med days':>9}")
        print("    " + "-" * 56)
        for k, r in g.iterrows():
            print(f"    {str(k):<9} {int(r['n']):>5} {r['yld']:>6.2f}% "
                  f"{r['drop']:>+7.2f}% {r['ratio']:>7.2f} "
                  f"{r['recov']*100:>7.1f}% {r['med_days']:>9.0f}")

        print(f"\n  INTERIM vs FINAL")
        gk = db.groupby("kind").agg(
            n=("kind", "size"), yld=("yield_pct", "median"),
            ratio=("drop_ratio", "median"), recov=("recovered", "mean"),
            med_days=("recovery_days", "median"))
        for k, r in gk.iterrows():
            print(f"    {k:<8} n={int(r['n']):>4}  yield {r['yld']:.2f}%  "
                  f"ratio {r['ratio']:.2f}  recovered {r['recov']*100:.0f}%  "
                  f"median {r['med_days']:.0f}d")

        db.to_csv(PROC / "dividend_behaviour.csv", index=False)
        bysym = (db.groupby("symbol")
                   .agg(n=("amount", "size"), yld=("yield_pct", "median"),
                        ratio=("drop_ratio", "median"),
                        recov=("recovered", "mean"),
                        med_days=("recovery_days", "median"))
                   .sort_values("med_days"))
        bysym.to_csv(PROC / "dividend_behaviour_by_symbol.csv")
        print(f"\n  FASTEST / SLOWEST RECOVERERS (median trading days)")
        for lab, sub in (("fast", bysym.head(5)), ("slow", bysym.tail(5))):
            for sym, r in sub.iterrows():
                print(f"    {lab:>4}  {sym:<12} n={int(r['n']):>2}  "
                      f"yield {r['yld']:.2f}%  ratio {r['ratio']:.2f}  "
                      f"median {r['med_days']:.0f}d  "
                      f"recovered {r['recov']*100:.0f}%")

    print(f"\nSaved -> {PROC}\\results_behaviour.csv")
    print(f"      -> {PROC}\\dividend_behaviour.csv")
    print(f"      -> plus *_by_symbol.csv for both")


if __name__ == "__main__":
    main()
