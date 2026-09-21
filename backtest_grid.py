"""
backtest_grid.py
===============================================================================
STAGE 3 — the pre-registered verdict.

Runs exactly the grid fixed in processed/preregistration_v2.json before any
result was seen: 12 configurations, chosen on 2008-2021, confirmed on 2022-2023,
and evaluated ONCE on 2024+.

    conviction   none | q40 | q70   (percentiles of DIRECTIONAL |score|, dev era)
    sizing       equal | riskparity (w proportional to 1/sigma, capped 3x)
    construction market | sectorneutral (demean by industry x month)

Costs are TIERED by liquidity — 0.20 / 0.40 / 0.80% round trip across turnover
terciles. The earlier flat 0.20% was too generous once 1,400+ stocks including
small caps entered, which is exactly where the reversal is expected to be
strongest. If tiered costs kill the result, that is the finding.

Ships only if ALL of: OOS month-clustered t > 2.0, deflated Sharpe > 0 at
n_trials=12, PBO < 0.5. Otherwise it is reported as a negative result.

    python backtest_grid.py              # dev + validation only, no OOS peek
    python backtest_grid.py --final      # ...and the single OOS evaluation
===============================================================================
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import score_feeds

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:                                 # noqa: BLE001
    pass

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"
CAR_FILE = PROC / "disclosure_timing_announcements.csv"
PREREG = PROC / "preregistration_v2.json"
CACHE = PROC / "grid_calls.parquet"

DEV_END, VAL_END = 2021, 2023
COST_TIERS = (0.20, 0.40, 0.80)                   # large / mid / small, round trip
MAX_CONCURRENT = 40
HOLD = 20


# --------------------------------------------------------------------------- #
def turnover_map() -> dict[str, pd.Series]:
    """Per-symbol rolling 60-session median traded value (close x volume).

    The pre-registration buckets by traded value in the 60 sessions BEFORE each
    filing, so this is a rolling series per symbol, not one number per stock —
    a company that was small in 2010 and large in 2024 must not be charged
    2024's costs on its 2010 filings.
    """
    out = {}
    for p in (PROC / "raw_price_cache").glob("*.csv"):
        try:
            d = pd.read_csv(p, parse_dates=["date"], usecols=["date", "close", "volume"])
        except Exception:                         # noqa: BLE001
            continue
        if d.empty or d["volume"].isna().all():
            continue
        d = d.set_index("date").sort_index()
        tv = (d["close"] * d["volume"]).rolling(60, min_periods=20).median()
        out[p.stem] = tv.dropna()
    return out


def build_calls() -> pd.DataFrame:
    """Every directional tag with its outcome, sector, sigma and cost tier."""
    if CACHE.exists():
        return pd.read_parquet(CACHE)

    car = pd.read_csv(CAR_FILE, usecols=["symbol", "ts", "car_20", "sigma"])
    car["ts"] = pd.to_datetime(car["ts"], errors="coerce")
    car = car.dropna(subset=["ts", "car_20", "sigma"])
    car = car[car["sigma"] > 1e-6]

    tmap = turnover_map()
    print(f"  turnover series for {len(tmap):,} symbols")

    rows = []
    for sym, sub in car.groupby("symbol"):
        path = ROOT / sym / "announcements.csv"
        if not path.exists():
            continue
        try:
            a = pd.read_csv(path, dtype=str).fillna("")
        except Exception:                         # noqa: BLE001
            continue
        for c in ("desc", "attchmntText", "sort_date", "smIndustry"):
            if c not in a.columns:
                a[c] = ""
        a["_ts"] = pd.to_datetime(a["sort_date"], errors="coerce", format="ISO8601")
        a = a.dropna(subset=["_ts"]).drop_duplicates(subset=["_ts"])
        cache = score_feeds.load_cache(sym)
        if not cache:
            continue

        m = sub.merge(a[["_ts", "desc", "attchmntText", "smIndustry"]],
                      left_on="ts", right_on="_ts")
        if m.empty:
            continue
        tv = tmap.get(sym)
        for _i, r in m.iterrows():
            hit = cache.get(score_feeds.text_key(r["desc"], r["attchmntText"]))
            if not hit or hit["label"] == "Neutral":
                continue
            turn = np.nan
            if tv is not None and len(tv):
                pos = tv.index.searchsorted(r["ts"]) - 1
                if pos >= 0:
                    turn = float(tv.iloc[pos])
            rows.append({
                "symbol": sym, "ts": r["ts"], "tag": hit["label"],
                "engine": hit.get("engine", "?"), "score": abs(float(hit["score"])),
                "car": float(r["car_20"]), "sigma": float(r["sigma"]),
                "industry": (str(r["smIndustry"]).strip() or "UNKNOWN"),
                "turnover": turn,
            })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["year"] = df["ts"].dt.year
    df["month"] = df["ts"].dt.to_period("M")
    df = df.sort_values("ts").reset_index(drop=True)
    df.to_parquet(CACHE, index=False)
    return df


def add_cost(df: pd.DataFrame) -> pd.DataFrame:
    """Tiered round-trip cost from turnover terciles, computed per calendar year
    so a stock is judged against its contemporaries, not against 2026 liquidity."""
    d = df.copy()
    d["cost"] = COST_TIERS[1]
    for _y, idx in d.groupby("year").groups.items():
        sub = d.loc[idx, "turnover"]
        ok = sub.dropna()
        if len(ok) < 30:
            continue
        lo, hi = ok.quantile([1 / 3, 2 / 3])
        d.loc[idx, "cost"] = np.where(sub >= hi, COST_TIERS[0],
                                      np.where(sub <= lo, COST_TIERS[2], COST_TIERS[1]))
    d.loc[d["turnover"].isna(), "cost"] = COST_TIERS[2]   # unknown -> assume worst
    return d


# --------------------------------------------------------------------------- #
def apply_config(df, conviction, sizing, construction, cuts=None):
    """One configuration -> per-call net return. `cuts` are dev-era conviction
    thresholds, passed in so later eras never re-derive their own."""
    d = df
    if conviction != "none":
        q = 0.40 if conviction == "q40" else 0.70
        if cuts is None:
            cuts = {e: g["score"].quantile(q) for e, g in d.groupby("engine")}
        keep = d["score"] >= d["engine"].map(cuts).fillna(0.0)
        d = d[keep]
    d = d.copy()

    # fade the tone: Negative -> long, Positive -> short
    d["side"] = np.where(d["tag"] == "Negative", 1, -1)
    ret = d["side"] * d["car"]

    if construction == "sectorneutral":
        # Remove the industry-month mean so a sector-wide move is not credited
        # to the filing that happened to sit inside it.
        grp = d.groupby(["industry", "month"])["car"].transform("mean")
        ret = d["side"] * (d["car"] - grp)

    d["gross"] = ret
    if sizing == "riskparity":
        w = 1.0 / d["sigma"].clip(lower=1e-6)
        w = (w / w.median()).clip(upper=3.0)
        d["w"] = w
    else:
        d["w"] = 1.0
    d["net"] = d["gross"] - d["cost"]
    return d, cuts


def monthly(d: pd.DataFrame) -> pd.Series:
    """Weighted monthly mean return, with the concurrent-position cap applied."""
    out = {}
    for m, g in d.groupby("month"):
        if len(g) > MAX_CONCURRENT:
            g = g.nlargest(MAX_CONCURRENT, "score")   # keep highest conviction
        w = g["w"].to_numpy()
        out[m] = float((g["net"].to_numpy() * w).sum() / w.sum()) if w.sum() else 0.0
    return pd.Series(out).sort_index()


def tstat(ms: pd.Series) -> float:
    if len(ms) < 3 or ms.std(ddof=1) == 0:
        return float("nan")
    return float(ms.mean() / (ms.std(ddof=1) / np.sqrt(len(ms))))


def summarise(d: pd.DataFrame) -> dict:
    ms = monthly(d)
    sr = (ms.mean() / ms.std(ddof=1) * np.sqrt(12)) if ms.std(ddof=1) > 0 else np.nan
    return {"n": len(d), "win": float((d["net"] > 0).mean() * 100),
            "mean": float(d["net"].mean()), "t": tstat(ms),
            "sharpe": float(sr), "months": len(ms), "monthly": ms}


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--final", action="store_true",
                    help="run the single pre-registered out-of-sample evaluation")
    args = ap.parse_args()

    reg = json.loads(PREREG.read_text(encoding="utf-8"))
    print("=" * 96)
    print("  STAGE 3 — PRE-REGISTERED GRID")
    print(f"  registered {reg['registered_utc']} · n_trials={reg['grid']['n_trials']}")
    print(f"  bar: {' AND '.join(reg['decision_rule']['ships_only_if_all_hold'])}")
    print("=" * 96)

    df = build_calls()
    if df.empty:
        raise SystemExit("no calls built")
    df = add_cost(df)
    print(f"  {len(df):,} directional calls · {df['symbol'].nunique():,} symbols · "
          f"{df['ts'].min().date()} to {df['ts'].max().date()}")
    print(f"  cost mix: " + "  ".join(
        f"{c}%={int((df['cost'] == c).sum()):,}" for c in COST_TIERS))

    dev = df[df.year <= DEV_END]
    val = df[(df.year > DEV_END) & (df.year <= VAL_END)]
    oos = df[df.year > VAL_END]
    print(f"  dev {len(dev):,} · val {len(val):,} · oos {len(oos):,}\n")

    CONV = ["none", "q40", "q70"]
    SIZE = ["equal", "riskparity"]
    CONS = ["market", "sectorneutral"]

    print(f"  {'#':<3}{'conviction':<12}{'sizing':<13}{'construction':<16}"
          f"{'calls':>9}{'win%':>7}{'mean':>9}{'t':>7}{'Sharpe':>8}")
    print("  " + "-" * 92)
    results, monthlies = [], []
    for c in CONV:
        for s in SIZE:
            for k in CONS:
                d, cuts = apply_config(dev, c, s, k)
                r = summarise(d)
                results.append({"conviction": c, "sizing": s, "construction": k,
                                "cuts": cuts, **{x: r[x] for x in
                                                 ("n", "win", "mean", "t", "sharpe")}})
                monthlies.append(r["monthly"])
                print(f"  {len(results):<3}{c:<12}{s:<13}{k:<16}{r['n']:>9,}"
                      f"{r['win']:>7.1f}{r['mean']:>+9.3f}{r['t']:>+7.2f}"
                      f"{r['sharpe']:>8.2f}")

    best = max(results, key=lambda r: (r["t"] if np.isfinite(r["t"]) else -9))
    print(f"\n  DEV WINNER (highest month-clustered t, as pre-registered): "
          f"{best['conviction']} / {best['sizing']} / {best['construction']}"
          f"   t={best['t']:+.2f}")

    dv, _ = apply_config(val, best["conviction"], best["sizing"],
                         best["construction"], cuts=best["cuts"])
    rv = summarise(dv)
    print(f"  VALIDATION 2022-2023: {rv['n']:,} calls  win {rv['win']:.1f}%  "
          f"mean {rv['mean']:+.3f}%  t={rv['t']:+.2f}  Sharpe {rv['sharpe']:.2f}")

    # PBO over the 12 configs, aligned on the dev months.
    from event_stats import deflated_sharpe, pbo_cscv
    idx = sorted(set().union(*[set(m.index) for m in monthlies]))
    perf = np.column_stack([m.reindex(idx).fillna(0.0).to_numpy() for m in monthlies])
    pbo = pbo_cscv(perf)
    pbo_v = pbo.get("pbo") if isinstance(pbo, dict) else pbo
    print(f"  PBO across the 12 configs: {pbo_v}")

    if not args.final:
        print("\n  Out-of-sample NOT evaluated. Re-run with --final to spend it.")
        return

    do, _ = apply_config(oos, best["conviction"], best["sizing"],
                         best["construction"], cuts=best["cuts"])
    ro = summarise(do)
    ds = deflated_sharpe(ro["monthly"].to_numpy(), n_trials=12)
    print("\n" + "=" * 96)
    print(f"  OUT-OF-SAMPLE 2024+  (evaluated once)")
    print(f"    calls {ro['n']:,}  win {ro['win']:.1f}%  mean {ro['mean']:+.3f}%  "
          f"t={ro['t']:+.2f}  Sharpe {ro['sharpe']:.2f}")
    print(f"    deflated Sharpe: sr={ds.get('sr')} dsr={ds.get('dsr')}")
    ok_t = np.isfinite(ro["t"]) and ro["t"] > 2.0
    ok_d = (ds.get("dsr") or 0) > 0
    ok_p = (pbo_v is not None) and pbo_v < 0.5
    print(f"\n    t > 2.0            {ok_t}   ({ro['t']:+.2f})")
    print(f"    deflated Sharpe>0  {ok_d}   ({ds.get('dsr')})")
    print(f"    PBO < 0.5          {ok_p}   ({pbo_v})")
    print("=" * 96)
    print("  SHIPS" if (ok_t and ok_d and ok_p) else
          "  DOES NOT SHIP — negative result, per the pre-registered rule.")
    print("=" * 96)
    (PROC / "stage3_result.json").write_text(json.dumps({
        "winner": {k: best[k] for k in ("conviction", "sizing", "construction")},
        "dev_t": best["t"], "val_t": rv["t"], "oos_t": ro["t"],
        "oos_sharpe": ro["sharpe"], "oos_win": ro["win"], "oos_n": ro["n"],
        "dsr": ds.get("dsr"), "pbo": pbo_v,
        "ships": bool(ok_t and ok_d and ok_p)}, indent=1, default=str), encoding="utf-8")


if __name__ == "__main__":
    main()
