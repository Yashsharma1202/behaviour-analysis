"""
event_ml.py
===============================================================================
TIER 2 — MACHINE LEARNING ON THE EVENT PANEL, DONE THE WAY THE LITERATURE SAYS.

event_stats.py showed the 64-cell argmax is a lucky-cell generator. This module
replaces it with a model of the whole response surface, validated with the
cross-validation machinery that overlapping financial labels actually require.

WHAT IS IMPLEMENTED
-------------------
  [9]  SURFACE MODEL, not cell selection. days_before becomes a FEATURE, so all
       8 entry offsets of every event train one model and neighbouring cells
       share information instead of competing in an argmax.  -> build_samples()
  [10] POOLED PANEL. Every event of every symbol, ~234k samples instead of the
       48 rows ml_model.py trains on.                        -> build_samples()
  [11] PURGED + EMBARGOED K-FOLD. Training rows whose label window overlaps a
       test row are dropped, plus an embargo after the test block. Without this
       the 16-day windows leak across the fold boundary.     -> PurgedKFold
  [12] COMBINATORIAL PURGED CV. C(N,k) chronology-respecting splits give a
       DISTRIBUTION of out-of-sample performance, not one number.  -> cpcv_splits()
  [13] SAMPLE UNIQUENESS WEIGHTING. Concurrency-based average uniqueness, so a
       cluster of five overlapping earnings does not count as five independent
       facts.                                               -> uniqueness_weights()
  [14] TRIPLE-BARRIER LABELS. Profit target / stop loss / time limit, whichever
       comes first — instead of an arbitrary fixed horizon.  -> triple_barrier()
  [15] MDA FEATURE IMPORTANCE on purged out-of-sample folds, replacing sklearn's
       in-sample, cardinality-biased feature_importances_.   -> mda_importance()

  Plus PBO over hyper-parameter configs, so "which model did I pick" is itself
  tested for overfitting.

Barriers are hit on the RAW price path (that is what a real stop-loss sees), but
P&L is also reported MARKET-ADJUSTED, because event_stats.py showed ~90% of the
raw move is market drift.

RUN
    python event_ml.py                     # full panel, CPCV, MDA
    python event_ml.py --pt 2.0 --sl 1.0   # barrier widths in sigma
    python event_ml.py --quick             # fewer CPCV splits
===============================================================================
"""

from __future__ import annotations

import argparse
import math
import sys
import time
import warnings
from itertools import combinations
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

NB, NA = ES.NB, ES.NA
COST = ES.COST


# ---------------------------------------------------------------------------
# [14] Triple-barrier labelling
# ---------------------------------------------------------------------------
def triple_barrier(panel, sym, t0, t_vert, sigma, pt=2.0, sl=1.0):
    """Label each sample by whichever barrier is touched first.

    sym    : [N] symbol index
    t0     : [N] entry day index (buy at this day's close)
    t_vert : [N] latest exit day (the time limit)
    sigma  : [N] daily volatility used to scale the barriers
    pt/sl  : profit-target / stop-loss width, in units of sigma*sqrt(horizon)

    Returns (ret_raw, ret_mkt, hold_days, touch) where touch is
    +1 profit target, -1 stop loss, 0 time limit.

    This removes the arbitrary fixed horizon that the 64-cell grid was
    implicitly optimising over: the exit is decided by the path, not by us.
    """
    dev = panel.dev
    N = t0.shape[0]
    Hmax = int((t_vert - t0).max().item())
    j = torch.arange(1, Hmax + 1, device=dev)                    # [H]
    idx = (t0.unsqueeze(1) + j.unsqueeze(0)).clamp(0, panel.D - 1)   # [N, H]

    # cumulative simple return along the path, from the entry close
    r = panel.ret[sym.unsqueeze(1), idx]                          # [N, H]
    m = panel.mret[idx]
    alive = (t0.unsqueeze(1) + j.unsqueeze(0)) <= t_vert.unsqueeze(1)
    r = torch.where(alive, r, torch.zeros_like(r))
    m = torch.where(alive, m, torch.zeros_like(m))
    cr = torch.cumprod(1.0 + r, dim=1) - 1.0                      # path return
    cm = torch.cumprod(1.0 + m, dim=1) - 1.0

    horizon = (t_vert - t0).clamp(min=1).to(torch.float32)
    band = sigma.unsqueeze(1) * torch.sqrt(horizon).unsqueeze(1)
    up = cr >= (pt * band)
    dn = cr <= (-sl * band)

    big = Hmax + 10
    first_up = torch.where(up & alive, j.unsqueeze(0), torch.full_like(idx, big)).min(1).values
    first_dn = torch.where(dn & alive, j.unsqueeze(0), torch.full_like(idx, big)).min(1).values
    first_vt = horizon.to(torch.long)

    hit = torch.minimum(torch.minimum(first_up, first_dn), first_vt)   # [N]
    touch = torch.where(first_up < torch.minimum(first_dn, first_vt), 1,
                        torch.where(first_dn < torch.minimum(first_up, first_vt), -1, 0))

    gather = (hit - 1).clamp(0, Hmax - 1).unsqueeze(1)
    ret_raw = cr.gather(1, gather).squeeze(1)
    ret_mkt = cm.gather(1, gather).squeeze(1)
    return ret_raw, ret_mkt, hit, touch


# ---------------------------------------------------------------------------
# [13] Sample uniqueness weighting
# ---------------------------------------------------------------------------
def uniqueness_weights(t0: np.ndarray, t1: np.ndarray, n_days: int):
    """Average uniqueness of each label (Lopez de Prado).

    Concurrency c_t = how many label windows span day t.
    Uniqueness u_i  = mean over the label's own days of 1/c_t.

    A lone event scores ~1.0; one of five overlapping earnings scores ~0.2, so
    the cluster contributes about one independent observation instead of five.
    """
    conc = np.zeros(n_days + 2, dtype=np.float64)
    np.add.at(conc, t0, 1.0)
    np.add.at(conc, t1 + 1, -1.0)
    conc = np.cumsum(conc)[:n_days + 1]
    conc = np.maximum(conc, 1.0)

    inv = 1.0 / conc
    cinv = np.concatenate([[0.0], np.cumsum(inv)])
    span = (t1 - t0 + 1).astype(np.float64)
    u = (cinv[t1 + 1] - cinv[t0]) / span
    return u, conc


# ---------------------------------------------------------------------------
# [9] + [10] Build the pooled sample matrix
# ---------------------------------------------------------------------------
FEATURES = [
    "days_before",          # the grid dimension, now a feature (item 9)
    "mom_5", "mom_20", "mom_60",
    "vol_20", "vol_60",
    "mkt_mom_20", "rel_strength_60",
    "dist_52w_high", "above_ma50",
    "beta", "resid_sd",
    "month", "quarter",
    "n_concurrent", "days_since_prev",
    "et_ANNOUNCEMENT", "et_BOARD_MEETING", "et_CORPORATE_ACTION", "et_RESULTS",
]


def build_samples(panel, types, alpha, beta, resid_sd, pt=2.0, sl=1.0):
    """One row per (event, days_before). Features are strictly as-of the entry
    day, labels come from the triple barrier that starts the day after."""
    dev = panel.dev
    P = panel.ret
    D = panel.D

    # price level (for 52w high / MA) reconstructed from returns
    logp = torch.cumsum(torch.log1p(P.clamp(min=-0.99)), dim=1)

    ev_pos, ev_sym, ev_type = panel.pos, panel.sym_idx, panel.etype
    E = ev_pos.shape[0]

    blocks = []
    for db in range(1, NB + 1):
        t0 = ev_pos - db
        t_vert = ev_pos + NA
        ok = (t0 > 300) & (t_vert < D - 1)
        blocks.append((t0[ok], ev_sym[ok], ev_type[ok], ev_pos[ok],
                       t_vert[ok], torch.full((int(ok.sum()),), db,
                                              device=dev, dtype=torch.long)))
    t0 = torch.cat([b[0] for b in blocks])
    sym = torch.cat([b[1] for b in blocks])
    etype = torch.cat([b[2] for b in blocks])
    epos = torch.cat([b[3] for b in blocks])
    tv = torch.cat([b[4] for b in blocks])
    dbv = torch.cat([b[5] for b in blocks])
    N = t0.shape[0]

    def px_ret(a, b):
        """simple return between two day indices, per sample"""
        return torch.exp(logp[sym, b] - logp[sym, a]) - 1.0

    sig = resid_sd[sym, t0]

    # ---- features, all evaluated at t0 (no look-ahead) --------------------
    f = {}
    f["days_before"] = dbv.to(torch.float32)
    f["mom_5"] = px_ret(t0 - 5, t0)
    f["mom_20"] = px_ret(t0 - 20, t0)
    f["mom_60"] = px_ret(t0 - 60, t0)

    off20 = torch.arange(-19, 1, device=dev)
    w20 = panel.ret[sym.unsqueeze(1), (t0.unsqueeze(1) + off20).clamp(0, D - 1)]
    f["vol_20"] = w20.std(dim=1)
    off60 = torch.arange(-59, 1, device=dev)
    w60 = panel.ret[sym.unsqueeze(1), (t0.unsqueeze(1) + off60).clamp(0, D - 1)]
    f["vol_60"] = w60.std(dim=1)

    mcum = torch.cumsum(torch.log1p(panel.mret.clamp(min=-0.99)), dim=0)
    f["mkt_mom_20"] = torch.exp(mcum[t0] - mcum[(t0 - 20).clamp(min=0)]) - 1.0
    mkt60 = torch.exp(mcum[t0] - mcum[(t0 - 60).clamp(min=0)]) - 1.0
    f["rel_strength_60"] = f["mom_60"] - mkt60

    off252 = torch.arange(-251, 1, device=dev)
    w252 = logp[sym.unsqueeze(1), (t0.unsqueeze(1) + off252).clamp(0, D - 1)]
    f["dist_52w_high"] = torch.exp(logp[sym, t0] - w252.max(dim=1).values) - 1.0
    off50 = torch.arange(-49, 1, device=dev)
    ma50 = logp[sym.unsqueeze(1), (t0.unsqueeze(1) + off50).clamp(0, D - 1)].mean(1)
    f["above_ma50"] = (logp[sym, t0] > ma50).to(torch.float32)

    f["beta"] = beta[sym, t0]
    f["resid_sd"] = sig

    dt = pd.DatetimeIndex(panel.dates)[t0.cpu().numpy()]
    f["month"] = torch.tensor(dt.month.values, dtype=torch.float32, device=dev)
    f["quarter"] = torch.tensor(dt.quarter.values, dtype=torch.float32, device=dev)

    for i, tn in enumerate(types):
        f[f"et_{tn}"] = (etype == i).to(torch.float32)
    for name in FEATURES:
        f.setdefault(name, torch.zeros(N, device=dev))

    # ---- [14] labels ------------------------------------------------------
    ret_raw, ret_mkt, hold, touch = triple_barrier(panel, sym, t0, tv, sig, pt, sl)
    t1 = t0 + hold
    hf = hold.to(torch.float32).clamp(min=1)

    # Full market-model adjustment (alpha AND beta), matching event_stats.py.
    # Subtracting the raw index move assumes beta=1, which leaves residual
    # market drift in the label for high-beta names.
    a_d = alpha[sym, t0]
    b_d = beta[sym, t0]
    y_raw = ret_raw - COST
    y_abn = ret_raw - (a_d * hf + b_d * ret_mkt) - COST

    # RISK-ADJUSTED target. Barriers are scaled by sigma, so a high-vol name's
    # +2sigma is a far bigger % move than a low-vol name's. Training on the raw
    # % label therefore rewards the model for simply picking volatile stocks --
    # more risk, not more skill. Dividing by the position's own sigma removes
    # that channel, which is also what vol-targeted position sizing does.
    y_norm = y_abn / (sig * torch.sqrt(hf)).clamp(min=1e-6)

    # ---- [13] weights -----------------------------------------------------
    t0n, t1n = t0.cpu().numpy(), t1.cpu().numpy()
    uniq, conc = uniqueness_weights(t0n, t1n, D)

    f["n_concurrent"] = torch.tensor(conc[t0n], dtype=torch.float32, device=dev)
    prev = np.full(N, 60.0)
    ordr = np.lexsort((t0n, sym.cpu().numpy()))
    sN = sym.cpu().numpy()
    for k in range(1, len(ordr)):
        a, b = ordr[k - 1], ordr[k]
        if sN[a] == sN[b]:
            prev[b] = min(60.0, float(t0n[b] - t0n[a]))
    f["days_since_prev"] = torch.tensor(prev, dtype=torch.float32, device=dev)

    X = torch.stack([f[c] for c in FEATURES], dim=1)
    good = (torch.isfinite(X).all(1) & torch.isfinite(y_raw)
            & torch.isfinite(y_abn) & torch.isfinite(y_norm))
    keep = good.cpu().numpy()

    df = pd.DataFrame(X[good].cpu().numpy(), columns=FEATURES)
    df["y_raw"] = y_raw[good].cpu().numpy()
    df["y_abn"] = y_abn[good].cpu().numpy()
    df["y_norm"] = y_norm[good].cpu().numpy()
    df["t0"] = t0n[keep]
    df["t1"] = t1n[keep]
    df["hold"] = hold[good].cpu().numpy()
    df["touch"] = touch[good].cpu().numpy()
    df["w"] = uniq[keep]
    df["symbol"] = [panel.symbols[i] for i in sym[good].cpu().numpy()]
    df["event_type"] = [types[i] for i in etype[good].cpu().numpy()]
    return df.sort_values("t0").reset_index(drop=True)


# ---------------------------------------------------------------------------
# [11] Purged + embargoed K-fold
# ---------------------------------------------------------------------------
def purge_train(t0, t1, test_lo, test_hi, embargo):
    """Indices of training rows that do NOT overlap the test window.

    Drops any row whose label span [t0,t1] intersects the test span, plus rows
    starting inside the embargo band after it. This is what stops the 16-day
    label windows leaking the answer across the fold boundary.
    """
    overlap = (t1 >= test_lo) & (t0 <= test_hi)
    emb = (t0 > test_hi) & (t0 <= test_hi + embargo)
    return ~(overlap | emb)


def purged_kfold(df, n_splits=6, embargo_frac=0.01):
    t0 = df["t0"].to_numpy()
    t1 = df["t1"].to_numpy()
    n = len(df)
    embargo = int((t0.max() - t0.min()) * embargo_frac)
    bounds = np.linspace(0, n, n_splits + 1).astype(int)
    for i in range(n_splits):
        te = np.zeros(n, dtype=bool)
        te[bounds[i]:bounds[i + 1]] = True
        lo, hi = t0[te].min(), t1[te].max()
        tr = purge_train(t0, t1, lo, hi, embargo) & ~te
        yield np.where(tr)[0], np.where(te)[0]


# ---------------------------------------------------------------------------
# [12] Combinatorial Purged CV
# ---------------------------------------------------------------------------
def cpcv_splits(df, n_groups=8, k_test=2, embargo_frac=0.01, max_splits=None):
    """C(n_groups, k_test) chronology-respecting train/test partitions, each one
    purged and embargoed. Every period ends up tested many times, which yields a
    DISTRIBUTION of out-of-sample performance instead of a single number."""
    t0 = df["t0"].to_numpy()
    t1 = df["t1"].to_numpy()
    n = len(df)
    embargo = int((t0.max() - t0.min()) * embargo_frac)
    edges = np.linspace(0, n, n_groups + 1).astype(int)
    groups = [np.arange(edges[i], edges[i + 1]) for i in range(n_groups)]
    combos = list(combinations(range(n_groups), k_test))
    if max_splits and len(combos) > max_splits:
        rng = np.random.default_rng(0)
        combos = [combos[i] for i in rng.choice(len(combos), max_splits, replace=False)]
    for c in combos:
        te = np.concatenate([groups[i] for i in c])
        mask_te = np.zeros(n, dtype=bool)
        mask_te[te] = True
        tr = np.ones(n, dtype=bool)
        for i in c:                       # purge around EACH test block
            lo, hi = t0[groups[i]].min(), t1[groups[i]].max()
            tr &= purge_train(t0, t1, lo, hi, embargo)
        tr &= ~mask_te
        yield np.where(tr)[0], te, c


# ---------------------------------------------------------------------------
# Model
# ---------------------------------------------------------------------------
def make_model(cfg):
    import xgboost as xgb
    return xgb.XGBRegressor(
        device="cuda" if torch.cuda.is_available() else "cpu",
        tree_method="hist",
        n_estimators=cfg.get("n_estimators", 300),
        max_depth=cfg.get("max_depth", 4),
        learning_rate=cfg.get("learning_rate", 0.03),
        subsample=0.8, colsample_bytree=0.8,
        min_child_weight=cfg.get("min_child_weight", 50),
        reg_lambda=cfg.get("reg_lambda", 2.0),
        verbosity=0, n_jobs=0,
    )


def rank_ic(pred, y, w=None):
    """Weighted Spearman rank IC — the standard score for return prediction.
    Accuracy is useless here because the base rate is ~50%."""
    if len(pred) < 10:
        return np.nan
    pr = pd.Series(pred).rank().to_numpy()
    yr = pd.Series(y).rank().to_numpy()
    if w is None:
        w = np.ones_like(pr)
    pr = pr - np.average(pr, weights=w)
    yr = yr - np.average(yr, weights=w)
    den = math.sqrt(np.average(pr ** 2, weights=w) * np.average(yr ** 2, weights=w))
    return float(np.average(pr * yr, weights=w) / den) if den > 0 else np.nan


def strategy_return(pred, y, w, top_frac=0.2):
    """P&L of actually trading the model: take the top-quintile predictions."""
    if len(pred) < 20:
        return np.nan, np.nan, 0
    k = max(1, int(len(pred) * top_frac))
    sel = np.argsort(-pred)[:k]
    r, ww = y[sel], w[sel]
    mu = float(np.average(r, weights=ww))
    sd = float(math.sqrt(np.average((r - mu) ** 2, weights=ww))) or np.nan
    return mu, (mu / sd if sd and sd == sd else np.nan), k


# ---------------------------------------------------------------------------
# [15] MDA feature importance
# ---------------------------------------------------------------------------
def mda_importance(df, target="y_abn", n_splits=6, seed=0, n_repeat=1):
    """Mean Decrease Accuracy on PURGED out-of-sample folds.

    Replaces sklearn's feature_importances_ (mean decrease impurity), which is
    computed in-sample and biased toward high-cardinality features. Here each
    feature is shuffled in the TEST fold only and we measure how much the
    out-of-sample rank IC degrades.
    """
    rng = np.random.default_rng(seed)
    X = df[FEATURES].to_numpy(dtype=np.float32)
    y = df[target].to_numpy(dtype=np.float32)
    w = df["w"].to_numpy(dtype=np.float32)

    base_scores, imp = [], {f: [] for f in FEATURES}
    for tr, te in purged_kfold(df, n_splits=n_splits):
        if len(tr) < 500 or len(te) < 200:
            continue
        m = make_model({}).fit(X[tr], y[tr], sample_weight=w[tr])
        p = m.predict(X[te])
        base = rank_ic(p, y[te], w[te])
        base_scores.append(base)
        for j, fname in enumerate(FEATURES):
            drops = []
            for _ in range(n_repeat):
                Xp = X[te].copy()
                rng.shuffle(Xp[:, j])
                drops.append(base - rank_ic(m.predict(Xp), y[te], w[te]))
            imp[fname].append(float(np.mean(drops)))
    out = pd.DataFrame({
        "feature": FEATURES,
        "mda": [np.mean(imp[f]) if imp[f] else np.nan for f in FEATURES],
        "mda_std": [np.std(imp[f]) if imp[f] else np.nan for f in FEATURES],
    }).sort_values("mda", ascending=False).reset_index(drop=True)
    return out, float(np.mean(base_scores)) if base_scores else np.nan


# ---------------------------------------------------------------------------
# PBO over hyper-parameter configs
# ---------------------------------------------------------------------------
CONFIGS = [
    {"max_depth": 3, "learning_rate": 0.03, "n_estimators": 200},
    {"max_depth": 4, "learning_rate": 0.03, "n_estimators": 300},
    {"max_depth": 5, "learning_rate": 0.03, "n_estimators": 300},
    {"max_depth": 4, "learning_rate": 0.10, "n_estimators": 150},
    {"max_depth": 6, "learning_rate": 0.05, "n_estimators": 400},
    {"max_depth": 3, "learning_rate": 0.10, "n_estimators": 100},
    {"max_depth": 8, "learning_rate": 0.05, "n_estimators": 500},
    {"max_depth": 2, "learning_rate": 0.05, "n_estimators": 250},
]


def run_cpcv(df, target="y_norm", n_groups=8, k_test=2, max_splits=None,
             pnl_col="y_abn"):
    """Fit every config on every CPCV split -> [n_splits, n_configs] OOS matrix.

    The model is TRAINED on `target` (risk-adjusted by default) but P&L is always
    scored on `pnl_col`, the realised market-adjusted % return, so the reported
    return is what you would actually have earned.
    """
    X = df[FEATURES].to_numpy(dtype=np.float32)
    y = df[target].to_numpy(dtype=np.float32)
    pnl = df[pnl_col].to_numpy(dtype=np.float32)
    w = df["w"].to_numpy(dtype=np.float32)

    splits = list(cpcv_splits(df, n_groups, k_test, max_splits=max_splits))
    print(f"  {len(splits)} CPCV splits x {len(CONFIGS)} configs "
          f"= {len(splits)*len(CONFIGS)} GPU fits")
    ic = np.full((len(splits), len(CONFIGS)), np.nan)
    ret = np.full((len(splits), len(CONFIGS)), np.nan)
    shp = np.full((len(splits), len(CONFIGS)), np.nan)
    t0 = time.time()
    for si, (tr, te, c) in enumerate(splits):
        if len(tr) < 1000 or len(te) < 500:
            continue
        for ci, cfg in enumerate(CONFIGS):
            m = make_model(cfg).fit(X[tr], y[tr], sample_weight=w[tr])
            p = m.predict(X[te])
            ic[si, ci] = rank_ic(p, y[te], w[te])
            mu, sr, _ = strategy_return(p, pnl[te], w[te])   # realised % P&L
            ret[si, ci], shp[si, ci] = mu, sr
        if (si + 1) % 5 == 0 or si == len(splits) - 1:
            el = time.time() - t0
            print(f"    split {si+1}/{len(splits)}  ({el:.0f}s elapsed)")
    return ic, ret, shp, splits


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("symbols", nargs="*")
    ap.add_argument("--pt", type=float, default=2.0, help="profit target (sigma)")
    ap.add_argument("--sl", type=float, default=1.0, help="stop loss (sigma)")
    ap.add_argument("--groups", type=int, default=8)
    ap.add_argument("--ktest", type=int, default=2)
    ap.add_argument("--target", default="y_norm",
                    choices=["y_norm", "y_abn", "y_raw"],
                    help="y_norm = risk-adjusted (default, honest); "
                         "y_abn = market-adjusted pct; y_raw = raw pct")
    ap.add_argument("--placebo", action="store_true",
                    help="replace every event date with a RANDOM date on the "
                         "same symbol. If the signal survives, it is a generic "
                         "cross-sectional factor, not an event effect.")
    ap.add_argument("--quick", action="store_true")
    ap.add_argument("--cpu", action="store_true")
    args = ap.parse_args()

    dev = ES.pick_device(args.cpu)
    print("=" * 78)
    print("  TIER 2 — ML ON THE EVENT PANEL")
    print(f"  device: {ES.describe_device(dev)}")
    print(f"  triple barrier: +{args.pt}s profit / -{args.sl}s stop / "
          f"{NA}d limit | cost {COST*100:.2f}%")
    print("=" * 78)

    symbols = args.symbols or sorted(
        set(EB.discover_feed_symbols()) &
        set(__import__("download_feeds").NIFTY50_FALLBACK))
    panel, types = ES.build_panel(symbols, dev)

    if args.placebo:
        # PLACEBO: keep every symbol, event type and count, but move each event
        # to a random date. Anything that survives this is not about events.
        g = torch.Generator(device=dev if dev.type == "cuda" else "cpu")
        g.manual_seed(12345)
        lo = ES.EST_GAP + ES.EST_LEN + NB + 5
        hi = panel.D - NA - 2
        panel.pos = torch.randint(lo, hi, panel.pos.shape, generator=g, device=dev)
        print("\n  *** PLACEBO MODE: event dates replaced with random dates ***")

    alpha, beta, resid_sd = ES.rolling_market_model(panel)

    print("\n[9][10][13][14] Building pooled sample matrix ...")
    t0 = time.time()
    df = build_samples(panel, types, alpha, beta, resid_sd, args.pt, args.sl)
    print(f"  {len(df):,} samples x {len(FEATURES)} features "
          f"({time.time()-t0:.1f}s)")
    print(f"  = {df['symbol'].nunique()} symbols, "
          f"{len(df)//NB:,} events x {NB} entry offsets")
    tt = df["touch"].value_counts()
    print(f"  barrier touched: profit {int(tt.get(1,0)):,} | "
          f"stop {int(tt.get(-1,0)):,} | time limit {int(tt.get(0,0)):,}")
    print(f"  mean holding: {df['hold'].mean():.1f} days "
          f"(fixed-grid was always {NB+NA})")
    print(f"\n[13] Sample uniqueness: mean {df['w'].mean():.3f}, "
          f"median {df['w'].median():.3f}")
    print(f"     -> {len(df):,} rows carry the weight of "
          f"~{df['w'].sum():,.0f} independent observations")

    print(f"\n  Label mean  raw, net of cost            : "
          f"{df['y_raw'].mean()*100:+.4f}%")
    print(f"  Label mean  market-model adjusted       : "
          f"{df['y_abn'].mean()*100:+.4f}%   "
          f"(median {df['y_abn'].median()*100:+.4f}%)")
    print(f"  Label mean  RISK-adjusted (sigma units) : "
          f"{df['y_norm'].mean():+.4f}")
    print(f"  training target = {args.target}  |  P&L scored on y_abn")

    df.to_csv(PROC / "event_ml_samples.csv", index=False)

    # ---- [11] purged K-fold baseline ------------------------------------
    print("\n[11] Purged + embargoed K-fold ...")
    X = df[FEATURES].to_numpy(dtype=np.float32)
    y = df[args.target].to_numpy(dtype=np.float32)
    w = df["w"].to_numpy(dtype=np.float32)
    naive_ic, purged_ic = [], []
    for i, (tr, te) in enumerate(purged_kfold(df, n_splits=6)):
        n_all = len(df) - len(te)
        m = make_model({}).fit(X[tr], y[tr], sample_weight=w[tr])
        purged_ic.append(rank_ic(m.predict(X[te]), y[te], w[te]))
        # deliberately UNPURGED comparison, to show the leak
        tr2 = np.setdiff1d(np.arange(len(df)), te)
        m2 = make_model({}).fit(X[tr2], y[tr2], sample_weight=w[tr2])
        naive_ic.append(rank_ic(m2.predict(X[te]), y[te], w[te]))
        print(f"   fold {i+1}: train {len(tr):,} purged "
              f"(dropped {n_all-len(tr):,} overlapping) | "
              f"IC purged {purged_ic[-1]:+.4f} vs unpurged {naive_ic[-1]:+.4f}")
    print(f"\n   mean OOS rank IC  — purged  : {np.nanmean(purged_ic):+.4f}")
    print(f"   mean OOS rank IC  — UNpurged: {np.nanmean(naive_ic):+.4f}"
          f"   <- the leak, if positive gap")

    # ---- [12] CPCV --------------------------------------------------------
    print("\n[12] Combinatorial Purged CV ...")
    ic, ret, shp, splits = run_cpcv(
        df, args.target, args.groups, args.ktest,
        max_splits=6 if args.quick else None, pnl_col="y_abn")

    best_cfg = int(np.nanargmax(np.nanmean(ic, axis=0)))
    print(f"\n   OOS rank IC by config (mean over {len(splits)} splits):")
    for ci, cfg in enumerate(CONFIGS):
        star = " <- best" if ci == best_cfg else ""
        print(f"     depth {cfg['max_depth']} lr {cfg['learning_rate']:.2f} "
              f"n {cfg['n_estimators']:>3}: IC {np.nanmean(ic[:,ci]):+.4f} "
              f"+/- {np.nanstd(ic[:,ci]):.4f} | "
              f"top-quintile ret {np.nanmean(ret[:,ci])*100:+.4f}%{star}")

    d = ic[:, best_cfg]
    d = d[np.isfinite(d)]
    print(f"\n   Best config OOS IC distribution: mean {d.mean():+.4f}, "
          f"5th pct {np.percentile(d,5):+.4f}, 95th {np.percentile(d,95):+.4f}")
    print(f"   P(IC > 0) across splits = {(d>0).mean():.3f}")

    r = ret[:, best_cfg]
    r = r[np.isfinite(r)]
    print(f"   Top-quintile OOS return: mean {r.mean()*100:+.4f}%, "
          f"P(>0) = {(r>0).mean():.3f}")

    # ---- PBO --------------------------------------------------------------
    valid = np.isfinite(ic).all(axis=1)
    if valid.sum() >= 6:
        pbo = ES.pbo_cscv(ic[valid])
        print(f"\n   Probability of Backtest Overfitting over "
              f"{len(CONFIGS)} configs: {pbo['pbo']:.3f}")

    ds = ES.deflated_sharpe(r, n_trials=len(CONFIGS) * max(len(splits), 1))
    print(f"   Deflated Sharpe of the selected config: {ds['dsr']:.4f} "
          f"(SR {ds['sr']:+.4f} vs null max {ds['sr0']:+.4f})")

    # ---- [15] MDA ---------------------------------------------------------
    print("\n[15] MDA feature importance (purged OOS folds) ...")
    mda, base = mda_importance(df, args.target, n_splits=6)
    print(f"   baseline OOS rank IC: {base:+.4f}")
    print(f"   {'feature':22} {'MDA':>9} {'std':>9}")
    print("   " + "-" * 42)
    for _, row in mda.head(12).iterrows():
        bar = "#" * max(0, int(row["mda"] * 2000))
        print(f"   {row['feature']:22} {row['mda']:+9.5f} "
              f"{row['mda_std']:9.5f}  {bar}")
    mda.to_csv(PROC / "event_ml_mda.csv", index=False)

    pd.DataFrame({
        "config": [str(c) for c in CONFIGS],
        "oos_ic_mean": np.nanmean(ic, axis=0),
        "oos_ic_std": np.nanstd(ic, axis=0),
        "oos_ret_mean": np.nanmean(ret, axis=0),
    }).to_csv(PROC / "event_ml_cpcv.csv", index=False)

    # ---- effective sample size ------------------------------------------
    # [13] said it: 232k overlapping rows are NOT 232k facts. Judge the IC
    # against the number of INDEPENDENT observations, not the row count, or the
    # t-statistic is inflated by three orders of magnitude.
    n_eff = float(df["w"].sum())
    t_eff = d.mean() * math.sqrt(max(n_eff, 1.0))
    t_naive = d.mean() * math.sqrt(len(df))
    print("\n" + "=" * 78)
    print(f"  EFFECTIVE SAMPLE SIZE (from [13] uniqueness weights)")
    print(f"    rows                       : {len(df):,}")
    print(f"    independent observations   : {n_eff:,.0f}")
    print(f"    IC t-stat on ROW count     : {t_naive:+.2f}   <- badly inflated")
    print(f"    IC t-stat on EFFECTIVE n   : {t_eff:+.2f}   <- the honest one")
    print(f"    Harvey-Liu multiple-testing bar for a new finding: |t| > 3.0")
    print("=" * 78)

    if d.mean() <= 0 or (d > 0).mean() <= 0.6:
        verdict = "NO USABLE SIGNAL — OOS IC indistinguishable from zero"
    elif abs(t_eff) > 3.0:
        verdict = "SIGNAL — survives CPCV, PBO and the effective-n t-test"
    elif abs(t_eff) > 2.0:
        verdict = ("PROMISING BUT NOT ESTABLISHED — consistent OOS IC, but "
                   f"|t|={abs(t_eff):.2f} on effective n is below the 3.0 bar")
    else:
        verdict = (f"NOT ESTABLISHED — OOS IC is consistent but |t|={abs(t_eff):.2f} "
                   "on effective n; the row count was doing the persuading")
    print(f"  VERDICT: {verdict}")
    print("=" * 78)
    print(f"\nSaved -> {PROC}\\event_ml_samples.csv")
    print(f"      -> {PROC}\\event_ml_cpcv.csv")
    print(f"      -> {PROC}\\event_ml_mda.csv")


if __name__ == "__main__":
    main()
