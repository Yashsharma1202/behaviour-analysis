"""
event_stats.py
===============================================================================
GPU-ACCELERATED STATISTICAL VALIDATION OF THE "BUY BEFORE / SELL AFTER" RULES.

event_behaviour.py measures raw returns and reports the best of a 64-cell grid.
That procedure manufactures significance: the expected maximum t-statistic over
N independent pure-noise cells is

    E[max t] ~ (1-g)*Z^-1(1 - 1/N) + g*Z^-1(1 - 1/(N*e)),    g = 0.5772

which for N=64 is ~2.4 — comfortably past the usual "significant" bar of 2.0,
with zero real edge. This module replaces that with the tests the event-study
and data-snooping literature actually requires.

WHAT IT COMPUTES
----------------
 1. ABNORMAL returns from a market model estimated on a clean pre-event window
    (T-250..T-30), not raw returns.                       -> market_model()
 2. BMP standardised cross-sectional test — valid under event-induced variance,
    which earnings always cause.                          -> bmp_test()
 3. Kolari-Pynnonen adjustment for cross-sectional correlation, because 50
    stocks report in the same 3-week window.              -> bmp_test(kp=True)
 4. Corrado rank test — non-parametric, survives fat tails. -> corrado_test()
 5. PERMUTATION TEST OF THE WHOLE GRID SEARCH. Shuffles every event to a random
    date and re-runs all 64 cells, 10,000x, to get the distribution of "best
    cell under pure noise". The real winner must beat that. -> permutation_test()
 6. White's Reality Check / Hansen's SPA / Romano-Wolf stepdown via the
    stationary bootstrap.                                 -> reality_check()
 7. Deflated Sharpe Ratio + Probability of Backtest Overfitting (CSCV).
                                                          -> deflated_sharpe(), pbo_cscv()
 8. Wilson score intervals on every win rate.             -> wilson_ci()

Everything heavy runs on the GPU as batched tensors. The rolling market model is
computed ONCE per symbol via cumulative sums, so a permutation is just a gather —
which is what makes 10,000 grid re-searches take seconds instead of days.

RUN
    python event_stats.py                    # all Nifty-50 stocks with feeds
    python event_stats.py RELIANCE TCS       # specific symbols
    python event_stats.py --perms 20000      # more permutations
    python event_stats.py --cpu              # force CPU (works, just slower)
===============================================================================
"""

from __future__ import annotations

import argparse
import math
import sys
import time
import warnings
from pathlib import Path

warnings.filterwarnings("ignore", category=UserWarning)

import numpy as np
import pandas as pd

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # noqa: BLE001
    pass

import torch

import event_behaviour as EB

ROOT = Path(__file__).resolve().parent
PROC = ROOT / "processed"

NB, NA = EB.N_BEFORE, EB.N_AFTER      # 8 x 8 grid, same as the dashboard
EST_LEN = 221                          # estimation window length (T-250..T-30)
EST_GAP = 30                           # gap between estimation window and event
COST = 0.0025                          # realistic Indian round-trip (0.25%)
EULER = 0.5772156649015329


# ---------------------------------------------------------------------------
# Device
# ---------------------------------------------------------------------------
def pick_device(force_cpu=False) -> torch.device:
    if force_cpu or not torch.cuda.is_available():
        return torch.device("cpu")
    return torch.device("cuda")


def describe_device(dev: torch.device) -> str:
    if dev.type == "cpu":
        return "CPU"
    p = torch.cuda.get_device_properties(0)
    return (f"{torch.cuda.get_device_name(0)}  sm_{p.major}{p.minor}  "
            f"{p.total_memory / 1e9:.1f} GB")


# ---------------------------------------------------------------------------
# Panel construction
# ---------------------------------------------------------------------------
class Panel:
    """Aligned price/return panel plus an event list, all on one device.

    prices are aligned to a single trading calendar (the union of all symbols'
    dates) so that a day index means the same thing for every symbol and for the
    market index.
    """

    def __init__(self, ret, mret, valid, sym_idx, pos, etype, symbols, dates, dev):
        self.ret = ret            # [S, D] simple returns, 0 where missing
        self.mret = mret          # [D]    market returns
        self.valid = valid        # [S, D] bool, True where the symbol has data
        self.sym_idx = sym_idx    # [E]
        self.pos = pos            # [E] day index of the event
        self.etype = etype        # [E] int code
        self.symbols = symbols
        self.dates = dates
        self.dev = dev
        self.S, self.D = ret.shape


def load_market() -> pd.Series:
    cache = PROC / "price_cache" / "_NSEI.csv"
    if cache.exists():
        s = pd.read_csv(cache, parse_dates=["date"]).set_index("date")["adj"]
        return s.sort_index()
    import nifty_backtest as NBT
    return NBT.fetch_index()


def build_panel(symbols: list[str], dev: torch.device) -> tuple[Panel, list[str]]:
    print(f"Loading prices for {len(symbols)} symbols ...")
    px_map, kept = {}, []
    for s in symbols:
        p = EB.fetch_prices_yahoo(s)
        if p is not None and len(p) > EST_LEN + EST_GAP + NB + NA + 10:
            px_map[s] = p
            kept.append(s)
    if not kept:
        raise SystemExit("No usable price history.")

    mkt = load_market()
    # One calendar for everything: the market index's trading days.
    cal = mkt.index.sort_values().unique()
    date_pos = pd.Series(np.arange(len(cal)), index=cal)
    D = len(cal)
    S = len(kept)

    prices = np.full((S, D), np.nan, dtype=np.float64)
    for i, s in enumerate(kept):
        aligned = px_map[s].reindex(cal)
        prices[i] = aligned.to_numpy(dtype=np.float64)

    mp = mkt.reindex(cal).to_numpy(dtype=np.float64)

    # simple returns; missing -> 0 and masked out via `valid`
    with np.errstate(invalid="ignore", divide="ignore"):
        ret = np.zeros_like(prices)
        ret[:, 1:] = prices[:, 1:] / prices[:, :-1] - 1.0
        mret = np.zeros(D)
        mret[1:] = mp[1:] / mp[:-1] - 1.0

    valid = np.isfinite(ret) & np.isfinite(prices)
    ret = np.where(valid, ret, 0.0)
    mret = np.nan_to_num(mret)

    print(f"  panel: {S} symbols x {D} trading days "
          f"({cal[0].date()} -> {cal[-1].date()})")

    # ---- events -----------------------------------------------------------
    print("Loading events from NSE feeds ...")
    ev = EB.load_events_from_feeds(kept)
    ev = ev[ev["symbol"].isin(kept)]
    sym_to_i = {s: i for i, s in enumerate(kept)}
    types = sorted(ev["event_type"].unique())
    type_to_i = {t: i for i, t in enumerate(types)}

    day = np.array(date_pos.reindex(ev["event_date"].values).to_numpy(),
                   dtype=np.float64, copy=True)
    # events falling on a non-trading day -> next trading day
    miss = ~np.isfinite(day)
    if miss.any():
        nxt = cal.searchsorted(pd.DatetimeIndex(ev["event_date"].values[miss]))
        day[miss] = np.where(nxt < D, nxt, -1)
    day = day.astype(np.int64)

    si = ev["symbol"].map(sym_to_i).to_numpy()
    ti = ev["event_type"].map(type_to_i).to_numpy()

    # keep only events with a full estimation window before and exit window after
    lo = EST_GAP + EST_LEN + NB
    ok = (day >= lo) & (day < D - NA - 1)
    si, day, ti = si[ok], day[ok], ti[ok]
    print(f"  events: {len(ev):,} loaded -> {len(day):,} usable "
          f"(need {lo} days of history before, {NA + 1} after)")

    t = lambda a, dt: torch.as_tensor(a, dtype=dt, device=dev)  # noqa: E731
    panel = Panel(t(ret, torch.float32), t(mret, torch.float32),
                  t(valid, torch.bool), t(si, torch.long), t(day, torch.long),
                  t(ti, torch.long), kept, cal, dev)
    return panel, types


# ---------------------------------------------------------------------------
# 1. Market model -> abnormal returns
# ---------------------------------------------------------------------------
def rolling_market_model(panel: Panel):
    """Rolling OLS of stock returns on market returns, per symbol per day.

    For day t the estimation window is [t-EST_GAP-EST_LEN+1, t-EST_GAP], i.e.
    it ENDS 30 trading days before t so the event itself never contaminates the
    parameters. Computed for every (symbol, day) at once with cumulative sums —
    O(S*D) instead of O(S*D*EST_LEN) — which is what makes the permutation test
    affordable later, since a permuted event is then only a gather.

    Returns alpha [S,D], beta [S,D], resid_sd [S,D], all NaN where the window
    is incomplete.
    """
    r, m = panel.ret, panel.mret.unsqueeze(0).expand_as(panel.ret)
    v = panel.valid.to(r.dtype)

    def cs(x):
        z = torch.zeros(x.shape[0], 1, device=x.device, dtype=x.dtype)
        return torch.cat([z, torch.cumsum(x, dim=1)], dim=1)   # [S, D+1]

    CS_y, CS_x = cs(r * v), cs(m * v)
    CS_xy, CS_xx = cs(r * m * v), cs(m * m * v)
    CS_yy, CS_n = cs(r * r * v), cs(v)

    D = panel.D
    t_idx = torch.arange(D, device=r.device)
    b = t_idx - EST_GAP                     # window end (inclusive)
    a = b - EST_LEN + 1                     # window start (inclusive)
    good = a >= 0

    ai = a.clamp(min=0)
    bi = b.clamp(min=0)

    def wsum(CS):
        return CS[:, bi + 1] - CS[:, ai]

    n = wsum(CS_n)
    sy, sx = wsum(CS_y), wsum(CS_x)
    sxy, sxx, syy = wsum(CS_xy), wsum(CS_xx), wsum(CS_yy)

    n = torch.where(n < EST_LEN * 0.6, torch.nan, n)     # demand a mostly-full window
    mean_x, mean_y = sx / n, sy / n
    cov = sxy / n - mean_x * mean_y
    var = sxx / n - mean_x * mean_x
    beta = cov / torch.where(var.abs() < 1e-12, torch.nan, var)
    alpha = mean_y - beta * mean_x

    # residual variance = Var(y) - beta^2 * Var(x), with the n/(n-2) correction
    var_y = syy / n - mean_y * mean_y
    resid_var = (var_y - beta * beta * var) * (n / (n - 2))
    resid_sd = torch.sqrt(resid_var.clamp(min=1e-12))

    bad = ~good.unsqueeze(0).expand_as(beta)
    beta = beta.masked_fill(bad, torch.nan)
    alpha = alpha.masked_fill(bad, torch.nan)
    resid_sd = resid_sd.masked_fill(bad, torch.nan)
    return alpha, beta, resid_sd


def abnormal_windows(panel, alpha, beta, resid_sd, pos=None, sym=None):
    """Abnormal returns over [pos-NB, pos+NA] for each event -> [E, NB+NA+1].

    One alpha/beta per event (estimated pre-event) applied across the whole
    window, which is standard event-study practice.
    """
    if pos is None:
        pos, sym = panel.pos, panel.sym_idx
    off = torch.arange(-NB, NA + 1, device=panel.dev)         # [W]
    idx = pos.unsqueeze(1) + off.unsqueeze(0)                 # [E, W]
    idx = idx.clamp(0, panel.D - 1)

    r = panel.ret[sym.unsqueeze(1), idx]                      # [E, W]
    m = panel.mret[idx]                                       # [E, W]
    a = alpha[sym, pos].unsqueeze(1)
    b = beta[sym, pos].unsqueeze(1)
    ar = r - (a + b * m)
    sd = resid_sd[sym, pos]
    return ar, sd, r                                          # abnormal, sd, RAW


def car_grid(ar: torch.Tensor) -> torch.Tensor:
    """Cumulative abnormal return for every (days_before, days_after) pair.

    ar : [E, W] with W = NB+NA+1, column j = day (pos-NB+j).
    out: [E, NB, NA] where out[e, i, j] = CAR of buying (i+1) days before and
         selling (j+1) days after -- i.e. the sum of ARs over that holding period.
    """
    E = ar.shape[0]
    z = torch.zeros(E, 1, device=ar.device, dtype=ar.dtype)
    cum = torch.cat([z, torch.cumsum(ar, dim=1)], dim=1)      # [E, W+1]
    db = torch.arange(1, NB + 1, device=ar.device)
    da = torch.arange(1, NA + 1, device=ar.device)
    j_in = (NB - db)                                          # entry column
    j_out = (NB + da)                                         # exit column
    return cum[:, j_out + 1].unsqueeze(1) - cum[:, j_in + 1].unsqueeze(2)


# ---------------------------------------------------------------------------
# 2/3. BMP standardised cross-sectional test (+ Kolari-Pynnonen)
# ---------------------------------------------------------------------------
def bmp_test(car: torch.Tensor, resid_sd: torch.Tensor, horizon: torch.Tensor,
             rbar: float = 0.0):
    """Boehmer-Musumeci-Poulsen standardised cross-sectional test.

    Each event's CAR is first standardised by its own estimation-period residual
    volatility (scaled by sqrt of the holding length). The test then uses the
    CROSS-SECTIONAL variance of those standardised values, which is what keeps
    it valid when the event itself inflates volatility -- exactly the earnings
    case, where a plain t-test is misspecified.

    rbar > 0 applies the Kolari-Pynnonen correction for cross-sectional
    correlation of abnormal returns (clustered event dates):

        t_KP = t_BMP * sqrt( (1 - rbar) / (1 + (N-1)*rbar) )

    car      : [E, NB, NA]
    resid_sd : [E]
    horizon  : [NB, NA] number of trading days held (for the sqrt-time scaling)
    """
    denom = resid_sd.view(-1, 1, 1) * torch.sqrt(horizon).unsqueeze(0)
    sar = car / denom                                          # standardised
    n = torch.isfinite(sar).sum(0).clamp(min=1)
    mean = torch.nanmean(sar, dim=0)
    var = torch.nanmean((sar - mean) ** 2, dim=0) * (n / (n - 1).clamp(min=1))
    t_bmp = mean / torch.sqrt((var / n).clamp(min=1e-20))
    if rbar and rbar > 0:
        nf = n.to(torch.float32)
        adj = math.sqrt(max(1e-9, 1.0 - rbar)) / torch.sqrt(1.0 + (nf - 1.0) * rbar)
        return t_bmp, t_bmp * adj
    return t_bmp, t_bmp


def average_cross_correlation(panel, alpha, beta, sample=400, seed=0):
    """Mean pairwise correlation of estimation-period abnormal returns.

    This is the `rbar` that Kolari-Pynnonen needs. Estimated from a random
    sample of symbol pairs over a common recent window.
    """
    g = torch.Generator(device="cpu").manual_seed(seed)
    S, D = panel.S, panel.D
    end = D - NA - 2
    start = max(0, end - EST_LEN)
    idx = torch.arange(start, end, device=panel.dev)
    m = panel.mret[idx].unsqueeze(0)
    r = panel.ret[:, idx]
    a = alpha[:, end].unsqueeze(1)
    b = beta[:, end].unsqueeze(1)
    ar = r - (a + b * m)
    ok = torch.isfinite(ar).all(1)
    ar = ar[ok]
    if ar.shape[0] < 3:
        return 0.0
    ar = ar - ar.mean(1, keepdim=True)
    sd = ar.norm(dim=1, keepdim=True).clamp(min=1e-12)
    C = (ar / sd) @ (ar / sd).T
    n = C.shape[0]
    off = ~torch.eye(n, dtype=torch.bool, device=C.device)
    return float(C[off].mean())


# ---------------------------------------------------------------------------
# 4. Corrado rank test
# ---------------------------------------------------------------------------
def corrado_test(panel, alpha, beta, pos=None, sym=None):
    """Corrado (1989) non-parametric rank test on the event-day abnormal return.

    Ranks each event's abnormal returns across its own estimation + event period
    and asks whether the event-window rank sits away from the middle. Uses only
    ranks, so fat tails, skew and event-induced variance don't break it.

    Pass pos/sym to test one event type; defaults to every event in the panel.
    """
    if pos is None:
        pos, sym = panel.pos, panel.sym_idx
    L = EST_LEN
    off = torch.arange(-(EST_GAP + L) + 1, NA + 1, device=panel.dev)
    idx = (pos.unsqueeze(1) + off.unsqueeze(0)).clamp(0, panel.D - 1)
    r = panel.ret[sym.unsqueeze(1), idx]
    m = panel.mret[idx]
    a = alpha[sym, pos].unsqueeze(1)
    b = beta[sym, pos].unsqueeze(1)
    ar = r - (a + b * m)                                       # [E, T]

    finite = torch.isfinite(ar)
    ar = torch.where(finite, ar, torch.zeros_like(ar))
    T = ar.shape[1]
    order = ar.argsort(dim=1)
    ranks = torch.empty_like(order)
    ar_ = torch.arange(T, device=ar.device).expand_as(order)
    ranks.scatter_(1, order, ar_)
    K = (ranks.to(torch.float32) + 1.0) / (T + 1.0)            # scaled ranks

    event_col = T - NA - 1                                     # the event day
    Ke = K[:, event_col]
    E = Ke.shape[0]
    dev_ = K - 0.5
    s_k = torch.sqrt((dev_ ** 2).mean(1).mean() + 1e-20)       # rank sd
    t_corrado = (Ke - 0.5).mean() / (s_k / math.sqrt(E))
    return float(t_corrado), E


# ---------------------------------------------------------------------------
# 5. Permutation test of the whole grid search
# ---------------------------------------------------------------------------
def permutation_test(panel, alpha, beta, resid_sd, real_best: float,
                     n_perm=10000, batch=250, seed=0, net_cost=True):
    """Distribution of the BEST CELL under the null of no event effect.

    Each permutation reassigns every event to a RANDOM date on its own symbol's
    calendar (so per-symbol count and volatility regime are preserved) and then
    re-runs the entire 64-cell search. The maximum over the grid is recorded.

    Comparing the real winner against this distribution is the empirical form
    of White's Reality Check: it prices in the fact that you searched 64 cells,
    with no distributional assumptions at all.
    """
    dev = panel.dev
    E = panel.pos.shape[0]
    sym = panel.sym_idx
    lo = EST_GAP + EST_LEN + NB
    hi = panel.D - NA - 2
    horizon = _horizon(dev)

    g = torch.Generator(device=dev if dev.type == "cuda" else "cpu")
    g.manual_seed(seed)

    off = torch.arange(-NB, NA + 1, device=dev)
    db = torch.arange(1, NB + 1, device=dev)
    da = torch.arange(1, NA + 1, device=dev)
    j_in, j_out = (NB - db) + 1, (NB + da) + 1     # +1 for the cumsum's leading 0

    maxima = torch.empty(n_perm, device=dev)
    done = 0
    t0 = time.time()
    while done < n_perm:
        B = min(batch, n_perm - done)
        rnd = torch.randint(lo, hi, (B, E), generator=g, device=dev)   # [B, E]
        idx = (rnd.unsqueeze(2) + off).clamp(0, panel.D - 1)           # [B, E, W]

        r = panel.ret[sym.view(1, -1, 1), idx]
        m = panel.mret[idx]
        a = alpha[sym.unsqueeze(0), rnd].unsqueeze(2)
        b = beta[sym.unsqueeze(0), rnd].unsqueeze(2)
        ar = r - (a + b * m)                                           # [B, E, W]

        z = torch.zeros(B, E, 1, device=dev, dtype=ar.dtype)
        cum = torch.cat([z, torch.cumsum(ar, dim=2)], dim=2)           # [B, E, W+1]

        # The grid mean is LINEAR in the cumulative AR, so average over events
        # FIRST and never materialise the [B, E, NB, NA] tensor at all. That is
        # what turns 10,000 grid re-searches into a few seconds.
        mcum = torch.nanmean(cum, dim=1)                               # [B, W+1]
        mean = (mcum[:, j_out].unsqueeze(1) - mcum[:, j_in].unsqueeze(2)) * 100.0
        if net_cost:
            mean = mean - COST * 100.0
        maxima[done:done + B] = torch.nan_to_num(
            mean.reshape(B, -1), nan=-1e9).max(dim=1).values

        done += B
        if done % max(batch * 4, 1) < B or done == n_perm:
            el = time.time() - t0
            print(f"    {done:,}/{n_perm:,} permutations  "
                  f"({done / max(el, 1e-9):,.0f}/s)")
    maxima = maxima.cpu().numpy()
    p = float((maxima >= real_best).mean())
    return maxima, p


def _horizon(dev):
    db = torch.arange(1, NB + 1, device=dev, dtype=torch.float32)
    da = torch.arange(1, NA + 1, device=dev, dtype=torch.float32)
    return db.unsqueeze(1) + da.unsqueeze(0)      # trading days held


# ---------------------------------------------------------------------------
# 6. White's Reality Check / Hansen's SPA / Romano-Wolf
# ---------------------------------------------------------------------------
def stationary_bootstrap_idx(n, B, mean_block=20, gen=None, dev="cpu"):
    """Politis-Romano stationary bootstrap indices [B, n]."""
    p = 1.0 / mean_block
    idx = torch.empty(B, n, dtype=torch.long, device=dev)
    cur = torch.randint(0, n, (B,), generator=gen, device=dev)
    for t in range(n):
        idx[:, t] = cur
        newstart = torch.rand(B, generator=gen, device=dev) < p
        cur = torch.where(newstart,
                          torch.randint(0, n, (B,), generator=gen, device=dev),
                          (cur + 1) % n)
    return idx


def reality_check(perf: torch.Tensor, n_boot=2000, mean_block=20, seed=0):
    """White's Reality Check + Hansen's SPA + Romano-Wolf stepdown.

    perf : [E, K] per-event performance of each of K rules (already net of costs
           and market-adjusted). Null: no rule beats zero.

    Returns dict with p_white, p_spa, and the set of rules surviving Romano-Wolf.
    """
    dev = perf.device
    E, K = perf.shape
    g = torch.Generator(device=dev if dev.type == "cuda" else "cpu")
    g.manual_seed(seed)

    fbar = torch.nanmean(perf, dim=0)                       # [K]
    sd = torch.sqrt(torch.nanmean((perf - fbar) ** 2, dim=0).clamp(min=1e-20))
    V = math.sqrt(E) * fbar
    V_max = float(V.max())

    idx = stationary_bootstrap_idx(E, n_boot, mean_block, g, dev)
    boot = perf[idx]                                        # [B, E, K]
    bmean = torch.nanmean(boot, dim=1)                      # [B, K]

    # --- White's Reality Check: recentre on the full mean -------------------
    Vb_white = math.sqrt(E) * (bmean - fbar.unsqueeze(0))
    p_white = float((Vb_white.max(dim=1).values >= V_max).float().mean())

    # --- Hansen's SPA: recentre only rules that are not badly bad -----------
    thresh = -sd * math.sqrt(2.0 * math.log(math.log(max(E, 3)))) / math.sqrt(E)
    keep = (fbar >= thresh).float()
    Vb_spa = math.sqrt(E) * (bmean - (fbar * keep).unsqueeze(0))
    p_spa = float((Vb_spa.max(dim=1).values >= V_max).float().mean())

    # --- Romano-Wolf stepdown: which rules individually survive -------------
    surviving, remaining = [], list(range(K))
    Vb = math.sqrt(E) * (bmean - fbar.unsqueeze(0))
    while remaining:
        sub = torch.tensor(remaining, device=dev)
        crit = torch.quantile(Vb[:, sub].max(dim=1).values, 0.95)
        stats = V[sub]
        win = (stats > crit).nonzero().flatten()
        if win.numel() == 0:
            break
        for w in win.tolist():
            surviving.append(remaining[w])
        remaining = [r for i, r in enumerate(remaining) if i not in set(win.tolist())]
    return {"p_white": p_white, "p_spa": p_spa,
            "rw_survivors": surviving, "V_max": V_max}


# ---------------------------------------------------------------------------
# 7. Deflated Sharpe Ratio + PBO
# ---------------------------------------------------------------------------
def _norm_cdf(x):
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))


def _norm_ppf(p):
    # Acklam's inverse normal CDF approximation
    a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02,
         1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00]
    b = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02,
         6.680131188771972e+01, -1.328068155288572e+01]
    c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00,
         -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00]
    d = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00,
         3.754408661907416e+00]
    pl, ph = 0.02425, 1 - 0.02425
    if p < pl:
        q = math.sqrt(-2 * math.log(p))
        return (((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / \
               ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1)
    if p > ph:
        q = math.sqrt(-2 * math.log(1 - p))
        return -(((((c[0]*q+c[1])*q+c[2])*q+c[3])*q+c[4])*q+c[5]) / \
                ((((d[0]*q+d[1])*q+d[2])*q+d[3])*q+1)
    q = p - 0.5
    r = q * q
    return (((((a[0]*r+a[1])*r+a[2])*r+a[3])*r+a[4])*r+a[5])*q / \
           (((((b[0]*r+b[1])*r+b[2])*r+b[3])*r+b[4])*r+1)


def expected_max_sharpe(n_trials: int, sd_sharpe: float = 1.0) -> float:
    """E[max SR] over n_trials independent strategies with zero true edge.
    Bailey & Lopez de Prado (2014). This is the bar a search must clear."""
    N = max(int(n_trials), 2)
    return sd_sharpe * ((1 - EULER) * _norm_ppf(1 - 1.0 / N)
                        + EULER * _norm_ppf(1 - 1.0 / (N * math.e)))


def deflated_sharpe(returns: np.ndarray, n_trials: int, sd_sharpe: float | None = None):
    """Deflated Sharpe Ratio: probability the observed SR is genuinely > 0 once
    you account for how many trials produced it, the sample length, skew and
    kurtosis (Bailey & Lopez de Prado 2014)."""
    r = np.asarray(returns, dtype=np.float64)
    r = r[np.isfinite(r)]
    T = len(r)
    if T < 8:
        return {"sr": np.nan, "sr0": np.nan, "dsr": np.nan, "T": T}
    mu, sd = r.mean(), r.std(ddof=1)
    if sd <= 0:
        return {"sr": np.nan, "sr0": np.nan, "dsr": np.nan, "T": T}
    sr = mu / sd
    z = (r - mu) / sd
    skew = float((z ** 3).mean())
    kurt = float((z ** 4).mean())
    if sd_sharpe is None:
        sd_sharpe = 1.0 / math.sqrt(max(T - 1, 1))
    sr0 = expected_max_sharpe(n_trials, sd_sharpe)
    denom = math.sqrt(max(1e-12, 1 - skew * sr + (kurt - 1) / 4.0 * sr * sr))
    dsr = _norm_cdf(((sr - sr0) * math.sqrt(max(T - 1, 1))) / denom)
    return {"sr": sr, "sr0": sr0, "dsr": dsr, "T": T,
            "skew": skew, "kurt": kurt}


def pbo_cscv(perf: np.ndarray, n_split: int = 12):
    """Probability of Backtest Overfitting via Combinatorial Symmetric CV
    (Bailey, Borwein, Lopez de Prado & Zhu).

    perf : [T, K] per-period performance of K rules.
    Splits the timeline into S blocks, takes every half as in-sample, picks the
    IS winner, and records its OUT-OF-SAMPLE rank. PBO = P(OOS rank below median).
    """
    from itertools import combinations
    T, K = perf.shape
    S = min(n_split, max(4, (T // 8) * 2))
    S -= S % 2
    if S < 4 or K < 2:
        return {"pbo": np.nan, "n_comb": 0}
    edges = np.linspace(0, T, S + 1).astype(int)
    blocks = [perf[edges[i]:edges[i + 1]] for i in range(S)]
    combos = list(combinations(range(S), S // 2))
    if len(combos) > 2000:
        rng = np.random.default_rng(0)
        combos = [combos[i] for i in rng.choice(len(combos), 2000, replace=False)]
    logits = []
    for c in combos:
        cset = set(c)
        IS = np.vstack([blocks[i] for i in range(S) if i in cset])
        OOS = np.vstack([blocks[i] for i in range(S) if i not in cset])
        with np.errstate(invalid="ignore", divide="ignore"):
            is_sr = np.nanmean(IS, 0) / (np.nanstd(IS, 0) + 1e-12)
            oos_sr = np.nanmean(OOS, 0) / (np.nanstd(OOS, 0) + 1e-12)
        if not np.isfinite(is_sr).any():
            continue
        star = int(np.nanargmax(is_sr))
        rank = (np.argsort(np.argsort(oos_sr))[star] + 1) / (K + 1)
        rank = min(max(rank, 1e-6), 1 - 1e-6)
        logits.append(math.log(rank / (1 - rank)))
    if not logits:
        return {"pbo": np.nan, "n_comb": 0}
    logits = np.array(logits)
    return {"pbo": float((logits <= 0).mean()), "n_comb": len(logits),
            "median_logit": float(np.median(logits))}


# ---------------------------------------------------------------------------
# 8. Wilson score interval
# ---------------------------------------------------------------------------
def wilson_ci(wins: int, n: int, conf: float = 0.95):
    """Wilson score interval for a win rate -- the honest version of the
    dashboard's bare '62%'."""
    if n == 0:
        return (np.nan, np.nan, np.nan)
    z = _norm_ppf(1 - (1 - conf) / 2)
    p = wins / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return (p * 100, max(0.0, c - h) * 100, min(1.0, c + h) * 100)


# ---------------------------------------------------------------------------
# Report
# ---------------------------------------------------------------------------
def analyse(panel, types, n_perm, out_csv=True):
    dev = panel.dev
    print("\nEstimating rolling market model (T-250..T-30) ...")
    t0 = time.time()
    alpha, beta, resid_sd = rolling_market_model(panel)
    torch.cuda.synchronize() if dev.type == "cuda" else None
    print(f"  done in {time.time() - t0:.2f}s "
          f"({panel.S} symbols x {panel.D} days, all windows at once)")

    rbar = average_cross_correlation(panel, alpha, beta)
    print(f"  mean cross-correlation of abnormal returns (rbar) = {rbar:+.4f}")

    horizon = _horizon(dev)
    rows = []
    summary = []

    for ti, tname in enumerate(types):
        mask = panel.etype == ti
        n_ev = int(mask.sum())
        if n_ev < 30:
            print(f"\n[{tname}] only {n_ev} events — skipped")
            continue

        sub_pos, sub_sym = panel.pos[mask], panel.sym_idx[mask]
        ar, sd, rawr = abnormal_windows(panel, alpha, beta, resid_sd,
                                        sub_pos, sub_sym)
        keep = torch.isfinite(ar).all(1) & torch.isfinite(sd) & (sd > 0)
        ar, sd, rawr = ar[keep], sd[keep], rawr[keep]
        sub_pos = sub_pos[keep]
        E = ar.shape[0]
        if E < 30:
            print(f"\n[{tname}] only {E} usable events — skipped")
            continue

        car = car_grid(ar)                                  # [E, NB, NA]
        net = car - COST                                    # charge costs
        mean_pct = torch.nanmean(net, dim=0) * 100
        win = (net > 0).to(torch.float32)
        win_rate = win.mean(0) * 100
        wins_n = win.sum(0)

        t_bmp, t_kp = bmp_test(car, sd, horizon, rbar=max(rbar, 0.0))

        # --- the cell the CURRENT dashboard would pick ----------------------
        # TRUE raw return -- what the dashboard actually reports: no market
        # model, no costs. This is the number the whole comparison hinges on.
        raw_mean = torch.nanmean(car_grid(rawr), dim=0) * 100
        abn_gross = torch.nanmean(car, dim=0) * 100
        flat_raw = raw_mean.flatten()
        i_raw = int(torch.nan_to_num(flat_raw, nan=-1e9).argmax())
        db_raw, da_raw = i_raw // NA + 1, i_raw % NA + 1

        # --- best on abnormal, net of costs --------------------------------
        flat = mean_pct.flatten()
        i_best = int(torch.nan_to_num(flat, nan=-1e9).argmax())
        db_b, da_b = i_best // NA + 1, i_best % NA + 1
        best_val = float(flat[i_best])

        print("\n" + "=" * 78)
        print(f"  {tname}   ({E:,} usable events, {panel.S} symbols)")
        print("=" * 78)
        print(f"  [1] What the dashboard reports (RAW, gross) : "
              f"buy {db_raw}d / sell {da_raw}d  ->  {float(flat_raw[i_raw]):+.3f}%")
        print(f"      same cell, market-model ABNORMAL         : "
              f"{float(abn_gross.flatten()[i_raw]):+.3f}%   "
              f"(market drift removed: "
              f"{float(flat_raw[i_raw]) - float(abn_gross.flatten()[i_raw]):+.3f}pp)")
        print(f"      best ABNORMAL cell, net of {COST*100:.2f}% cost  : "
              f"buy {db_b}d / sell {da_b}d  ->  {best_val:+.3f}%")

        wr, lo_, hi_ = wilson_ci(int(wins_n.flatten()[i_best]), E)
        print(f"\n  [8] Win rate at that cell : {wr:.1f}%   "
              f"Wilson 95% CI [{lo_:.1f}%, {hi_:.1f}%]"
              + ("   <- straddles 50%" if lo_ < 50 < hi_ else ""))

        print(f"\n  [2] BMP standardised cross-sectional t : "
              f"{float(t_bmp.flatten()[i_best]):+.3f}")
        print(f"  [3] Kolari-Pynnonen adjusted t (rbar={rbar:+.4f}) : "
              f"{float(t_kp.flatten()[i_best]):+.3f}")

        t_cor, n_cor = corrado_test(panel, alpha, beta, sub_pos, sub_sym[keep])
        print(f"  [4] Corrado rank test t (event day, this type)  : {t_cor:+.3f}")

        # --- 5. permutation test of the search ------------------------------
        print(f"\n  [5] Permutation test of the whole 64-cell search "
              f"({n_perm:,} shuffles) ...")
        sub_panel = Panel(panel.ret, panel.mret, panel.valid, sub_sym[keep],
                          sub_pos, panel.etype[mask][keep], panel.symbols,
                          panel.dates, dev)
        maxima, p_perm = permutation_test(sub_panel, alpha, beta, resid_sd,
                                          real_best=best_val, n_perm=n_perm)
        q95 = float(np.percentile(maxima, 95))
        print(f"      best cell under NOISE : mean {maxima.mean():+.3f}%,  "
              f"95th pct {q95:+.3f}%,  max {maxima.max():+.3f}%")
        print(f"      your best cell        : {best_val:+.3f}%")
        print(f"      permutation p-value   : {p_perm:.4f}"
              f"   {'** beats the noise ceiling' if p_perm < 0.05 else '<- INSIDE the noise distribution'}")

        # --- 6. Reality Check / SPA / Romano-Wolf ---------------------------
        perf = net.reshape(E, NB * NA)
        rc = reality_check(perf)
        print(f"\n  [6] White's Reality Check p : {rc['p_white']:.4f}")
        print(f"      Hansen's SPA p          : {rc['p_spa']:.4f}")
        print(f"      Romano-Wolf survivors   : {len(rc['rw_survivors'])} of {NB*NA} rules")

        # --- 7. DSR + PBO ----------------------------------------------------
        best_ret = perf[:, i_best].cpu().numpy()
        ds = deflated_sharpe(best_ret, n_trials=NB * NA)
        print(f"\n  [7] Sharpe (per event)      : {ds['sr']:+.4f}")
        print(f"      E[max Sharpe] under null : {ds['sr0']:+.4f}   (64 trials)")
        print(f"      Deflated Sharpe Ratio    : {ds['dsr']:.4f}"
              f"   {'** survives' if ds['dsr'] > 0.95 else '<- does NOT survive deflation'}")
        order = torch.argsort(sub_pos)
        pbo = pbo_cscv(perf[order].cpu().numpy())
        print(f"      Probability of Backtest Overfitting : {pbo['pbo']:.3f}"
              f"   ({pbo['n_comb']} splits)")

        verdict = ("REAL — survives permutation, deflation and reality check"
                   if (p_perm < 0.05 and ds["dsr"] > 0.95 and rc["p_white"] < 0.05)
                   else "NOT SUPPORTED — consistent with a lucky cell in a 64-cell search")
        print(f"\n  VERDICT: {verdict}")
        summary.append({
            "event_type": tname, "n_events": E,
            "dash_db": db_raw, "dash_da": da_raw,
            "dash_raw_pct": float(flat_raw[i_raw]),
            "best_db": db_b, "best_da": da_b, "best_abn_net_pct": best_val,
            "win_rate": wr, "wilson_lo": lo_, "wilson_hi": hi_,
            "t_bmp": float(t_bmp.flatten()[i_best]),
            "t_kp": float(t_kp.flatten()[i_best]),
            "t_corrado": t_cor,
            "perm_p": p_perm, "perm_noise_95pct": q95,
            "p_white": rc["p_white"], "p_spa": rc["p_spa"],
            "rw_survivors": len(rc["rw_survivors"]),
            "sharpe": ds["sr"], "sharpe_null": ds["sr0"], "dsr": ds["dsr"],
            "pbo": pbo["pbo"], "verdict": verdict,
        })

        for i in range(NB):
            for j in range(NA):
                rows.append({
                    "event_type": tname, "days_before": i + 1, "days_after": j + 1,
                    "n_events": E,
                    "abn_net_pct": float(mean_pct[i, j]),
                    "raw_pct": float(raw_mean[i, j]),
                    "win_rate_pct": float(win_rate[i, j]),
                    "t_bmp": float(t_bmp[i, j]), "t_kp": float(t_kp[i, j]),
                })

    if summary and out_csv:
        pd.DataFrame(summary).to_csv(PROC / "event_stats_summary.csv", index=False)
        pd.DataFrame(rows).to_csv(PROC / "event_stats_grid.csv", index=False)
        print(f"\nSaved -> {PROC}\\event_stats_summary.csv")
        print(f"      -> {PROC}\\event_stats_grid.csv")
    return summary


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("symbols", nargs="*")
    ap.add_argument("--perms", type=int, default=10000)
    ap.add_argument("--cpu", action="store_true")
    args = ap.parse_args()

    dev = pick_device(args.cpu)
    print("=" * 78)
    print("  STATISTICAL VALIDATION OF THE EVENT GRID")
    print(f"  device: {describe_device(dev)}")
    print(f"  grid: {NB}x{NA} = {NB*NA} rules | estimation window T-{EST_GAP+EST_LEN}..T-{EST_GAP}")
    print(f"  cost: {COST*100:.2f}% round-trip | permutations: {args.perms:,}")
    print("=" * 78)
    print(f"\n  Reference: E[max t] over {NB*NA} pure-noise cells "
          f"~ {expected_max_sharpe(NB*NA):.2f} sigma")
    print("  -> a 'best of 64' result must clear that bar before it means anything.\n")

    symbols = args.symbols or sorted(
        set(EB.discover_feed_symbols()) &
        set(__import__("download_feeds").NIFTY50_FALLBACK))
    panel, types = build_panel(symbols, dev)
    analyse(panel, types, args.perms)


if __name__ == "__main__":
    main()
