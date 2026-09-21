"""
disclosure_timing.py
===============================================================================
DISCLOSURE METADATA AS SIGNAL — the "HOW", not the "WHAT".

Everyone parses what a company discloses. Almost nobody uses HOW it discloses:
when it files, how long it took, whether it moved the meeting. This module
tests whether that forensic layer predicts subsequent abnormal returns.

THE THREE METRICS
-----------------
 A. ATTENTION GAP  — hours between the filing timestamp and the NEXT market
    open. A Tuesday 10am filing lands in a live market (gap 0). A Friday 8pm
    filing has ~63 hours of darkness before anyone can trade it. This is the
    clean, continuous version of "bad news gets buried on Friday evening".
    Source: an_dt on every announcement, 2004-2026.

 B. REPORTING LAG  — days from period end (toDate) to the results filing
    (broadCastDate), plus its Z-SCORE against that company's own history.
    The level says "this company is slow"; the CHANGE says "this company is
    slower than it usually is", which is the part that should carry news.
    Source: financial_results.

 C. BOARD-MEETING NOTICE — days between the intimation (bm_timestamp) and the
    meeting (bm_date), and whether the meeting was RESCHEDULED
    (oriiginalMeetingDate != proposedMeetingDate).
    Source: board_meetings.

NOTE ON `difference`
--------------------
The feeds carry a `difference` column, but it is the EXCHANGE's dissemination
latency (an_dt -> exchdisstime), almost always 0-10 SECONDS. That is NSE's
system, not company behaviour, so it is reported for completeness and excluded
from the hypotheses.

HOW IT IS TESTED
----------------
Same discipline as event_stats.py, because raw returns lie:
  * market-model ABNORMAL returns (rolling OLS, T-250..T-30)
  * forward CAR at +1/+5/+10/+20 trading days from the first tradeable session
  * BMP standardised cross-sectional t-statistic
  * top-vs-bottom decile spread with a Wilson interval on the win rate
  * PERMUTATION TEST that shuffles the metric across events, holding dates
    fixed — the null is "this metric carries no information", which is exactly
    what we want to reject
  * effective sample size from per-symbol label overlap, so the t-stats are
    not inflated by 200k correlated rows

RUN
    python disclosure_timing.py                 # all symbols with feeds
    python disclosure_timing.py --perms 20000
===============================================================================
"""

from __future__ import annotations

import argparse
import math
import sys
import time
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import torch

warnings.filterwarnings("ignore")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # noqa: BLE001
    pass

import event_behaviour as EB
import event_stats as ES

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"

HORIZONS = [1, 5, 10, 20]           # forward trading days for the CAR
MKT_OPEN, MKT_CLOSE = 9.25, 15.5    # NSE hours (9:15 - 15:30)


# ---------------------------------------------------------------------------
# A. Attention gap
# ---------------------------------------------------------------------------
def attention_gap_hours(ts: pd.Series, trading_days: pd.DatetimeIndex) -> np.ndarray:
    """Hours between a filing timestamp and the next moment the market is open.

    0      -> filed into a live session, maximum attention
    ~18    -> filed after Tuesday's close, tradeable Wednesday morning
    ~63    -> filed Friday evening, nobody can act until Monday
    higher -> filed before a long holiday weekend

    Using the real NSE trading calendar means holidays are handled for free.
    """
    tset = trading_days.normalize()
    out = np.full(len(ts), np.nan)
    tvals = tset.values.astype("datetime64[D]")

    for i, t in enumerate(ts.values):
        if pd.isna(t):
            continue
        t = pd.Timestamp(t)
        day = np.datetime64(t.normalize(), "D")
        hour = t.hour + t.minute / 60.0

        # inside a live session on a trading day -> zero gap
        j = np.searchsorted(tvals, day)
        same_day_open = (j < len(tvals)) and (tvals[j] == day)
        if same_day_open and MKT_OPEN <= hour < MKT_CLOSE:
            out[i] = 0.0
            continue

        # otherwise find the next session that opens strictly after this instant
        k = j if (same_day_open and hour < MKT_OPEN) else (j + 1 if same_day_open else j)
        if k >= len(tvals):
            continue
        nxt = pd.Timestamp(tvals[k]) + pd.Timedelta(hours=MKT_OPEN)
        out[i] = max(0.0, (nxt - t).total_seconds() / 3600.0)
    return out


def timing_bucket(ts: pd.Timestamp) -> str:
    if pd.isna(ts):
        return "UNKNOWN"
    h = ts.hour + ts.minute / 60.0
    dow = ts.dayofweek
    if dow >= 5:
        return "WEEKEND"
    if MKT_OPEN <= h < MKT_CLOSE:
        return "MARKET_HOURS"
    if MKT_CLOSE <= h < 18:
        return "POST_CLOSE"
    if 18 <= h < 21:
        return "FRI_EVENING" if dow == 4 else "EVENING"
    return "LATE_NIGHT"


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------
def _dt(s, fmts=("%d-%b-%Y %H:%M:%S", "%d-%b-%Y %H:%M", "%d-%b-%Y",
                 "%Y-%m-%d %H:%M:%S")):
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    for f in fmts:
        m = out.isna() & s.notna()
        if not m.any():
            break
        out[m] = pd.to_datetime(s[m], format=f, errors="coerce")
    m = out.isna() & s.notna()
    if m.any():
        out[m] = pd.to_datetime(s[m], errors="coerce", dayfirst=True)
    return out


def load_announcements(symbols):
    rows = []
    for s in symbols:
        p = ROOT / s / "announcements.csv"
        if not p.exists():
            continue
        d = pd.read_csv(p, dtype=str)
        if "an_dt" not in d.columns:
            continue
        ts = _dt(d["an_dt"])
        rows.append(pd.DataFrame({
            "symbol": s, "ts": ts,
            "desc": d.get("desc", pd.Series("", index=d.index)),
            "exch_latency": d.get("difference", pd.Series(np.nan, index=d.index)),
        }))
    if not rows:
        return pd.DataFrame()
    a = pd.concat(rows, ignore_index=True).dropna(subset=["ts"])
    return a.sort_values("ts").reset_index(drop=True)


def load_results(symbols):
    rows = []
    for s in symbols:
        p = ROOT / s / "financial_results.csv"
        if not p.exists():
            continue
        d = pd.read_csv(p, dtype=str)
        if "broadCastDate" not in d.columns or "toDate" not in d.columns:
            continue
        bc, td = _dt(d["broadCastDate"]), _dt(d["toDate"])
        rows.append(pd.DataFrame({
            "symbol": s, "ts": bc, "period_end": td,
            "relating": d.get("relatingTo", pd.Series("", index=d.index)),
        }))
    if not rows:
        return pd.DataFrame()
    r = pd.concat(rows, ignore_index=True).dropna(subset=["ts", "period_end"])
    r["lag_days"] = (r["ts"] - r["period_end"]).dt.total_seconds() / 86400.0
    r = r[(r.lag_days > 0) & (r.lag_days < 200)]
    # one row per (symbol, period): the FIRST filing for that quarter
    r = (r.sort_values("ts").groupby(["symbol", "period_end"], as_index=False).first())

    # B. deviation from the company's own habit, using only PAST filings
    r = r.sort_values(["symbol", "ts"]).reset_index(drop=True)
    g = r.groupby("symbol")["lag_days"]
    r["lag_mean"] = g.transform(lambda x: x.shift().expanding(min_periods=3).mean())
    r["lag_std"] = g.transform(lambda x: x.shift().expanding(min_periods=3).std())
    r["lag_z"] = (r["lag_days"] - r["lag_mean"]) / r["lag_std"].replace(0, np.nan)
    return r.sort_values("ts").reset_index(drop=True)


def load_board_meetings(symbols):
    rows = []
    for s in symbols:
        p = ROOT / s / "board_meetings.csv"
        if not p.exists():
            continue
        d = pd.read_csv(p, dtype=str)
        if "bm_date" not in d.columns:
            continue
        intim = _dt(d["bm_timestamp"]) if "bm_timestamp" in d.columns else pd.NaT
        meet = _dt(d["bm_date"])
        orig = _dt(d["oriiginalMeetingDate"]) if "oriiginalMeetingDate" in d.columns else pd.NaT
        prop = _dt(d["proposedMeetingDate"]) if "proposedMeetingDate" in d.columns else pd.NaT
        rows.append(pd.DataFrame({
            "symbol": s, "ts": intim, "meet_date": meet,
            "orig": orig, "prop": prop,
            "purpose": d.get("bm_purpose", pd.Series("", index=d.index)),
        }))
    if not rows:
        return pd.DataFrame()
    b = pd.concat(rows, ignore_index=True).dropna(subset=["ts", "meet_date"])
    b["notice_days"] = (b["meet_date"] - b["ts"]).dt.total_seconds() / 86400.0
    b["rescheduled"] = (b["orig"].notna() & b["prop"].notna()
                        & (b["orig"].dt.normalize() != b["prop"].dt.normalize())).astype(int)
    b = b[(b.notice_days > -1) & (b.notice_days < 90)]
    return b.sort_values("ts").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Forward abnormal returns
# ---------------------------------------------------------------------------
def attach_forward_car(df, panel, alpha, beta, resid_sd, cal):
    """Attach forward market-model CAR at each horizon.

    Entry is the first tradeable CLOSE strictly after the filing instant, so an
    after-hours disclosure is never credited with a move that happened before
    anyone could act on it.
    """
    dev = panel.dev
    sym_to_i = {s: i for i, s in enumerate(panel.symbols)}
    df = df[df["symbol"].isin(sym_to_i)].copy()

    # first trading day whose CLOSE is after the filing
    day = df["ts"].dt.normalize().values
    hour = df["ts"].dt.hour + df["ts"].dt.minute / 60.0
    pos = cal.searchsorted(pd.DatetimeIndex(day))
    same = (pos < len(cal)) & (cal.values[np.minimum(pos, len(cal) - 1)] == day)
    # filed before this session's close -> tradeable at today's close
    entry = np.where(same & (hour.values < MKT_CLOSE), pos, pos + 1)

    lo = ES.EST_GAP + ES.EST_LEN + 5
    ok = (entry >= lo) & (entry < len(cal) - max(HORIZONS) - 2)
    df = df[ok].copy()
    entry = entry[ok]
    df["entry_idx"] = entry
    si = df["symbol"].map(sym_to_i).to_numpy()

    t_sym = torch.as_tensor(si, dtype=torch.long, device=dev)
    t_pos = torch.as_tensor(entry, dtype=torch.long, device=dev)

    Hm = max(HORIZONS)
    off = torch.arange(1, Hm + 1, device=dev)
    idx = (t_pos.unsqueeze(1) + off).clamp(0, panel.D - 1)
    r = panel.ret[t_sym.unsqueeze(1), idx]
    m = panel.mret[idx]
    a = alpha[t_sym, t_pos].unsqueeze(1)
    b = beta[t_sym, t_pos].unsqueeze(1)
    ar = r - (a + b * m)
    cum = torch.cumsum(ar, dim=1)

    for h in HORIZONS:
        df[f"car_{h}"] = (cum[:, h - 1] * 100).cpu().numpy()
    df["sigma"] = resid_sd[t_sym, t_pos].cpu().numpy()
    df["sym_i"] = si
    return df.dropna(subset=[f"car_{h}" for h in HORIZONS])


# ---------------------------------------------------------------------------
# Testing
# ---------------------------------------------------------------------------
def effective_n(df, horizon):
    """Per-symbol overlap-adjusted sample size (cross-symbol residuals are
    ~uncorrelated after the market model, so overlap only matters within a
    symbol)."""
    tot = 0.0
    for _, g in df.groupby("symbol"):
        t0 = g["entry_idx"].to_numpy()
        t1 = t0 + horizon
        u, _ = _uniq(t0, t1)
        tot += u.sum()
    return max(tot, 1.0)


def _uniq(t0, t1):
    n_days = int(t1.max()) + 2
    conc = np.zeros(n_days + 2)
    np.add.at(conc, t0, 1.0)
    np.add.at(conc, t1 + 1, -1.0)
    conc = np.maximum(np.cumsum(conc)[:n_days + 1], 1.0)
    inv = 1.0 / conc
    cinv = np.concatenate([[0.0], np.cumsum(inv)])
    return (cinv[t1 + 1] - cinv[t0]) / (t1 - t0 + 1), conc


def permutation_metric(df, metric, horizon, n_perm=10000, seed=0, top_q=0.2):
    """Null: the metric carries no information.

    Shuffles the METRIC across events while holding dates, symbols and returns
    fixed, then recomputes the top-minus-bottom spread. This isolates the
    metric's contribution from everything else about the sample.
    """
    y = df[f"car_{horizon}"].to_numpy()
    x = df[metric].to_numpy()
    ok = np.isfinite(x) & np.isfinite(y)
    x, y = x[ok], y[ok]
    n = len(x)
    if n < 100:
        return np.nan, np.nan, np.array([])
    k = max(1, int(n * top_q))

    def spread(xs):
        o = np.argsort(xs)
        return y[o[-k:]].mean() - y[o[:k]].mean()

    real = spread(x)
    rng = np.random.default_rng(seed)
    null = np.empty(n_perm)
    for i in range(n_perm):
        null[i] = spread(rng.permutation(x))
    p = float((np.abs(null) >= abs(real)).mean())
    return real, p, null


def report_metric(df, metric, label, hypothesis, n_perm=10000, top_q=0.2):
    print("\n" + "=" * 78)
    print(f"  {label}")
    print(f"  hypothesis: {hypothesis}")
    print("=" * 78)
    d = df[np.isfinite(df[metric])]
    if len(d) < 200:
        print(f"  only {len(d)} usable observations — skipped")
        return None

    print(f"  n = {len(d):,}   {metric}: "
          f"median {d[metric].median():.2f}, "
          f"p10 {d[metric].quantile(.1):.2f}, p90 {d[metric].quantile(.9):.2f}")

    out = {"metric": metric, "n": len(d)}
    print(f"\n  {'horizon':>8} {'bottom20%':>11} {'top20%':>11} {'spread':>10} "
          f"{'BMP t':>8} {'perm p':>8} {'t(eff n)':>9}")
    print("  " + "-" * 74)
    for h in HORIZONS:
        y = d[f"car_{h}"].to_numpy()
        x = d[metric].to_numpy()
        k = max(1, int(len(d) * top_q))
        o = np.argsort(x)
        bot, top = y[o[:k]].mean(), y[o[-k:]].mean()

        real, p, _ = permutation_metric(d, metric, h, n_perm=n_perm, top_q=top_q)

        # BMP: standardise each CAR by its own estimation-period sigma
        sd = d["sigma"].to_numpy() * math.sqrt(h) * 100
        sar = y / np.where(sd > 0, sd, np.nan)
        hi = sar[o[-k:]]
        lo_ = sar[o[:k]]
        diff = hi[np.isfinite(hi)].mean() - lo_[np.isfinite(lo_)].mean()
        pooled = math.sqrt(np.nanvar(hi, ddof=1) / len(hi)
                           + np.nanvar(lo_, ddof=1) / len(lo_))
        t_bmp = diff / pooled if pooled > 0 else np.nan

        ne = effective_n(d, h)
        t_eff = t_bmp * math.sqrt(ne / len(d)) if np.isfinite(t_bmp) else np.nan

        flag = " **" if (p < 0.05 and abs(t_eff) > 2) else ""
        print(f"  {h:>7}d {bot:>+10.3f}% {top:>+10.3f}% {real:>+9.3f}% "
              f"{t_bmp:>+8.2f} {p:>8.4f} {t_eff:>+9.2f}{flag}")
        out[f"spread_{h}"] = real
        out[f"t_bmp_{h}"] = t_bmp
        out[f"perm_p_{h}"] = p
        out[f"t_eff_{h}"] = t_eff
        out[f"n_eff_{h}"] = ne
    return out


def test_bucket_contrast(df, buried_buckets, n_perm=10000, seed=0):
    """Direct test of the stated hypothesis: are filings in `buried_buckets`
    followed by worse abnormal returns than everything else?

    The continuous attention-gap metric is a blunt instrument here — it puts a
    normal 18-hour overnight gap in the same ranking as a 63-hour weekend. This
    tests the categorical contrast the hypothesis actually named.

    Null: bucket membership carries no information. Shuffles the LABEL across
    filings, holding dates and returns fixed.
    """
    print("\n" + "=" * 78)
    print(f"  A2. BURIED vs VISIBLE  ({' + '.join(buried_buckets)} vs the rest)")
    print("  hypothesis: bad news is filed when nobody is watching")
    print("=" * 78)
    is_b = df["bucket"].isin(buried_buckets).to_numpy()
    n_b, n_v = int(is_b.sum()), int((~is_b).sum())
    print(f"  buried n = {n_b:,}   visible n = {n_v:,}")
    print(f"\n  {'horizon':>8} {'buried':>10} {'visible':>10} {'diff':>10} "
          f"{'perm p':>9} {'t(eff n)':>9}")
    print("  " + "-" * 62)
    rng = np.random.default_rng(seed)
    out = {"metric": "buried_bucket", "n": len(df)}
    for h in HORIZONS:
        y = df[f"car_{h}"].to_numpy()
        real = y[is_b].mean() - y[~is_b].mean()
        null = np.empty(n_perm)
        for i in range(n_perm):
            p = rng.permutation(is_b)
            null[i] = y[p].mean() - y[~p].mean()
        pv = float((np.abs(null) >= abs(real)).mean())
        sd = math.sqrt(y[is_b].var(ddof=1) / n_b + y[~is_b].var(ddof=1) / n_v)
        t = real / sd if sd > 0 else np.nan
        ne = effective_n(df, h)
        t_eff = t * math.sqrt(ne / len(df))
        flag = " **" if (pv < 0.05 and abs(t_eff) > 2) else ""
        print(f"  {h:>7}d {y[is_b].mean():>+9.3f}% {y[~is_b].mean():>+9.3f}% "
              f"{real:>+9.3f}% {pv:>9.4f} {t_eff:>+9.2f}{flag}")
        out[f"spread_{h}"] = real
        out[f"perm_p_{h}"] = pv
        out[f"t_eff_{h}"] = t_eff
    return out


def report_buckets(df, col="bucket"):
    print("\n" + "=" * 78)
    print("  FILING-TIME BUCKETS — mean forward abnormal return")
    print("=" * 78)
    g = df.groupby(col)
    print(f"  {'bucket':<14} {'n':>7} " +
          " ".join(f"{'CAR+'+str(h)+'d':>10}" for h in HORIZONS))
    print("  " + "-" * 62)
    order = ["MARKET_HOURS", "POST_CLOSE", "EVENING", "FRI_EVENING",
             "LATE_NIGHT", "WEEKEND"]
    for b in order:
        if b not in g.groups:
            continue
        s = g.get_group(b)
        vals = " ".join(f"{s['car_'+str(h)].mean():>+9.3f}%" for h in HORIZONS)
        print(f"  {b:<14} {len(s):>7,} {vals}")


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("symbols", nargs="*")
    ap.add_argument("--perms", type=int, default=10000)
    ap.add_argument("--cpu", action="store_true")
    ap.add_argument("--all", action="store_true",
                    help="every stock with feeds AND a cached price series, "
                         "instead of the Nifty-50 intersection")
    args = ap.parse_args()

    dev = ES.pick_device(args.cpu)
    print("=" * 78)
    print("  DISCLOSURE METADATA AS SIGNAL — the 'how', not the 'what'")
    print(f"  device: {ES.describe_device(dev)} | permutations: {args.perms:,}")
    print("=" * 78)

    if args.all:
        # Every stock with BOTH feeds and a cached price series. Passing ~1,500
        # symbols as arguments would blow the Windows command-line limit, so the
        # set is resolved here rather than on the command line.
        _px = ROOT / "processed" / "price_cache"
        symbols = sorted(set(EB.discover_feed_symbols()) &
                         {p.stem for p in _px.glob("*.csv") if not p.stem.startswith("_")})
    else:
        symbols = args.symbols or sorted(
            set(EB.discover_feed_symbols()) &
            set(__import__("download_feeds").NIFTY50_FALLBACK))

    panel, _ = ES.build_panel(symbols, dev)
    alpha, beta, resid_sd = ES.rolling_market_model(panel)
    cal = pd.DatetimeIndex(panel.dates)
    print(f"  market model ready ({panel.S} symbols x {panel.D} days)")

    # ---- A. attention gap ------------------------------------------------
    print("\nLoading announcements ...")
    ann = load_announcements(panel.symbols)
    print(f"  {len(ann):,} announcements "
          f"({ann.ts.min().date()} -> {ann.ts.max().date()})")
    ann["gap_h"] = attention_gap_hours(ann["ts"], cal)
    ann["bucket"] = ann["ts"].map(timing_bucket)
    ann["is_buried"] = (ann["gap_h"] > 24).astype(int)

    lat = pd.to_timedelta(ann["exch_latency"].where(
        ann["exch_latency"].str.match(r"^\d+:\d+:\d+$", na=False)),
        errors="coerce").dt.total_seconds()
    print(f"  exchange latency (`difference`): median {lat.median():.0f}s, "
          f"p99 {lat.quantile(.99):.0f}s  -> NSE system time, not company "
          f"behaviour; excluded from tests")

    print("\n  filing-time distribution:")
    vc = ann["bucket"].value_counts()
    for b, n in vc.items():
        print(f"    {b:<14} {n:>7,}  ({n/len(ann)*100:4.1f}%)")

    ann = attach_forward_car(ann, panel, alpha, beta, resid_sd, cal)
    print(f"\n  {len(ann):,} announcements with tradeable forward returns")

    report_buckets(ann)
    r1 = report_metric(ann, "gap_h", "A1. ATTENTION GAP (hours to next market open)",
                       "filings buried before a weekend/holiday are followed by "
                       "WORSE abnormal returns -> spread should be NEGATIVE",
                       n_perm=args.perms)
    r1b = test_bucket_contrast(ann, ["FRI_EVENING", "WEEKEND"], n_perm=args.perms)

    # ---- B. reporting lag -------------------------------------------------
    print("\nLoading financial results ...")
    res = load_results(panel.symbols)
    if len(res):
        res = attach_forward_car(res, panel, alpha, beta, resid_sd, cal)
        print(f"  {len(res):,} results with lag + forward returns")
        r2 = report_metric(res, "lag_days", "B1. REPORTING LAG (days from period end)",
                           "slower filers underperform -> spread NEGATIVE",
                           n_perm=args.perms)
        r3 = report_metric(res.dropna(subset=["lag_z"]), "lag_z",
                           "B2. REPORTING-LAG Z-SCORE (vs the company's own habit)",
                           "filing later THAN USUAL is the real stress signal "
                           "-> spread NEGATIVE, and stronger than B1",
                           n_perm=args.perms)
    else:
        r2 = r3 = None

    # ---- C. board meetings ------------------------------------------------
    print("\nLoading board meetings ...")
    bm = load_board_meetings(panel.symbols)
    r4 = None
    if len(bm):
        bm = attach_forward_car(bm, panel, alpha, beta, resid_sd, cal)
        print(f"  {len(bm):,} board meetings with notice + forward returns")
        n_res = int(bm["rescheduled"].sum())
        print(f"  rescheduled meetings: {n_res:,} of {len(bm):,}")
        if n_res >= 30:
            a_ = bm[bm.rescheduled == 1]
            b_ = bm[bm.rescheduled == 0]
            print("\n  RESCHEDULED vs ON-SCHEDULE board meetings:")
            for h in HORIZONS:
                print(f"    +{h:>2}d  rescheduled {a_['car_'+str(h)].mean():+7.3f}%  "
                      f"on-schedule {b_['car_'+str(h)].mean():+7.3f}%  "
                      f"diff {a_['car_'+str(h)].mean()-b_['car_'+str(h)].mean():+7.3f}%")
        r4 = report_metric(bm, "notice_days", "C. BOARD-MEETING NOTICE PERIOD",
                           "minimum legal notice signals haste -> short notice "
                           "followed by worse returns, spread POSITIVE",
                           n_perm=args.perms)

    rows = [r for r in (r1, r1b, r2, r3, r4) if r]
    if rows:
        pd.DataFrame(rows).to_csv(PROC / "disclosure_timing.csv", index=False)
        ann.to_csv(PROC / "disclosure_timing_announcements.csv", index=False)
        print(f"\nSaved -> {PROC}\\disclosure_timing.csv")

    print("\n" + "=" * 78)
    print("  HOW TO READ IT")
    print("  * spread = mean CAR of the top 20% of the metric minus the bottom 20%")
    print("  * perm p tests the METRIC, holding dates and returns fixed")
    print("  * t(eff n) rescales the t-stat by per-symbol label overlap; only")
    print("    a result with perm p < 0.05 AND |t(eff n)| > 2 is marked **")
    print("=" * 78)


if __name__ == "__main__":
    main()
