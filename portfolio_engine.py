"""
Portfolio engine for the November Top-N walk-forward strategy (capital curve, utilisation, CAGR, Sharpe ...).
Rules (all stated in the Excel ReadMe):
  * each test year 2013-2025, for each calendar date d (Nov 1..30) the Top-N stocks/rules are chosen from PRIOR years only
    (same method as the validated walk-forward: >=10 prior events, positive avg, >=15 eligible stocks, rank win% then avg%)
  * each selected date is a 'sleeve': start-of-year equity is split equally across the K sleeves, and each sleeve equally across its N stocks
  * a position is opened at the close of T-n and closed at the close of T+m (trading days); LONG or SHORT, cash-settled P&L, 0.15% cost per position
  * daily mark-to-market equity on the NIFTY trading calendar; idle cash earns 0 (also reported at the risk-free rate)
Date policies: ALL (all 30 dates), HINDSIGHT (dates whose full-period WF result passes TRADE test - biased), EXPANDING (date qualifies in year Y if its WF
results over earlier years, >=5 of them, pass the TRADE test - honest, test window 2018-2025).
"""
import sys, glob, os
sys.path.insert(0, 'D:/behaviour analysis'); sys.stdout.reconfigure(encoding='utf-8')
import numpy as np, pandas as pd
import build_calendar_top_n as t
from statistics import NormalDist
b = t.b
ROOT = 'D:/behaviour analysis/'
COST, TRAIN_MIN, FIRST, LAST = 0.15, 10, 2013, 2025
RF = 0.065; CAP0 = 1_000_000.0
SIZES = (1, 3, 5, 10)

def load_prices():
    base = {}
    for f in glob.glob(ROOT + 'Stock_Data_213_Adjusted/daily_ohlc/*.csv'):
        s = os.path.basename(f)[:-4]
        if s in b.INDICES or s == 'SENSEX': continue
        d = pd.read_csv(f, usecols=['Date', 'Adj Close']); d['Date'] = pd.to_datetime(d['Date'])
        base[s] = d.sort_values('Date').drop_duplicates('Date').set_index('Date')['Adj Close']
    if 'TATAMOTORS' in base and 'TMPV' in base: del base['TATAMOTORS']
    return {s: p[p.index >= '2003-01-01'] for s, p in base.items() if len(p[p.index >= '2003-01-01']) >= 250}

def build_picks(PX):
    """picks[(d,Y)] = list (top 10) of dict(sym, side(+1/-1), n, m, gross, entry, exit, train_win, train_avg)"""
    picks = {}
    for d in range(1, 31):
        per = {}
        for s, p in PX.items():
            sm = b.samples(p, 11, d, s, f'{d:02d}-Nov')
            if sm: per[s] = sm
        for Y in range(FIRST, LAST + 1):
            cand = []
            for s, (ys, R, cals, idx) in per.items():
                iy = np.where(ys == Y)[0]
                if len(iy) == 0: continue
                tr = ys < Y
                if tr.sum() < TRAIN_MIN: continue
                k, win, avg, sd = b.best(R[tr])
                if avg <= 0: continue
                cand.append((win, avg, s, k, sd, R[iy[0], k] * sd, idx[iy[0]]))
            if len(cand) < 15: continue
            cand.sort(key=lambda x: (-x[0], -x[1]))
            lst = []
            for win, avg, s, k, sd, g, a in cand[:10]:
                n, m = k // 8 + 1, k % 8 + 1; dts = PX[s].index
                lst.append(dict(sym=s, side=int(sd), n=n, m=m, gross=float(g), entry=dts[a - n], exit=dts[a + m], train_win=float(win), train_avg=float(avg)))
            picks[(d, Y)] = lst
        print('picks built for', d, flush=True)
    return picks

def wf_returns(picks, N, cost=COST):
    """net return per (d,Y) of the equal-weight Top-N basket (gross mean - cost)"""
    return {(d, Y): np.mean([p['gross'] for p in l[:N]]) - cost for (d, Y), l in picks.items()}

def qualifies(hist, min_n=5):
    if len(hist) < min_n: return False
    h = np.array(hist); return h.mean() >= 1.0 and (h > 0).mean() * 100 >= 62

def select_dates(picks, N, policy, Y):
    r = wf_returns(picks, N)
    if policy == 'ALL': return [d for d in range(1, 31) if (d, Y) in picks]
    if policy == 'HINDSIGHT':
        return [d for d in range(1, 31) if qualifies([r[(d, y)] for y in range(FIRST, LAST + 1) if (d, y) in r], min_n=1) and (d, Y) in picks]
    if policy == 'EXPANDING':
        return [d for d in range(1, 31) if (d, Y) in picks and qualifies([r[(d, y)] for y in range(FIRST, Y) if (d, y) in r])]
    raise ValueError(policy)

def make_calendar(PX_full):
    nif = pd.read_csv(ROOT + 'Stock_Data_213/daily_ohlc/NIFTY.csv', usecols=['Date', 'Close']); nif['Date'] = pd.to_datetime(nif['Date'])
    nif = nif.sort_values('Date').drop_duplicates('Date').set_index('Date')['Close']
    cnt = pd.concat([pd.Series(1, index=p.index) for p in PX_full.values()]).groupby(level=0).sum()
    days = cnt[cnt >= 20].index.union(nif.index)
    days = days[(days >= f'{FIRST - 1}-12-31') & (days <= f'{LAST}-12-31')]
    return nif.reindex(days).ffill()

_path_cache = {}
def pos_path(PX, G, p):
    key = (p['sym'], p['entry'], p['exit'])
    if key in _path_cache: return _path_cache[key]
    ge = int(np.searchsorted(G, p['entry'])); gx = int(np.searchsorted(G, p['exit'], side='right') - 1)
    s = PX[p['sym']].loc[p['entry']:p['exit']]
    vals = s.reindex(G[ge:gx + 1], method='ffill')
    r = (vals.values / s.iloc[0] - 1.0)
    _path_cache[key] = (ge, gx, r); return _path_cache[key]

def simulate(PX, G, picks, N, policy, start_year, cost=COST, scale=1.0):
    """returns dict: equity (Series), deployed (Series), trades (DataFrame), sleeves per year"""
    Gd = G.index.values; nG = len(Gd)
    years = list(range(start_year, LAST + 1))
    first_i = int(np.searchsorted(Gd, np.datetime64(f'{start_year}-01-01')))
    eq = np.full(nG, np.nan); dep = np.zeros(nG); E = CAP0; trades = []; sleeve_log = []
    prev_end = first_i - 1
    for Y in years:
        y0 = int(np.searchsorted(Gd, np.datetime64(f'{Y}-01-01'))); y1 = int(np.searchsorted(Gd, np.datetime64(f'{Y + 1}-01-01'))) - 1
        dates_ = select_dates(picks, N, policy, Y); K = len(dates_)
        sleeve_log.append((Y, K, dates_))
        pnl = np.zeros(y1 - y0 + 1)
        for d in dates_:
            for rank, p in enumerate(picks[(d, Y)][:N], 1):
                stake = E / (K * N) * scale
                ge, gx, r = pos_path(PX, G.index, p)
                seg = stake * p['side'] * r
                pnl[ge - y0:gx - y0 + 1] += seg                                   # mark-to-market while held
                cost_amt = stake * cost / 100.0
                pnl[gx - y0:] += -cost_amt                                         # cost at exit
                pnl[gx - y0 + 1:] += seg[-1]                                       # realised gross P&L carried forward after exit
                dep[ge:gx + 1] += stake
                trades.append(dict(Strategy=f'{policy}-Top{N}', Year=Y, Sleeve=f'{d:02d}-Nov', Rank=rank, Stock=p['sym'], Side='LONG' if p['side'] == 1 else 'SHORT', Rule=f'T-{p["n"]} to T+{p["m"]}',
                                   Entry=pd.Timestamp(Gd[ge]), Exit=pd.Timestamp(Gd[gx]), Stake=stake, GrossPct=p['gross'], NetPct=p['gross'] - cost,
                                   PnL=stake * (p['gross'] - cost) / 100.0, HoldDays=gx - ge + 1))
        eq[y0:y1 + 1] = E + pnl
        E = E + pnl[-1]
    mask = ~np.isnan(eq); idx = np.where(mask)[0]
    ser = pd.Series(eq[idx], index=G.index[idx]); depser = pd.Series(dep[idx], index=G.index[idx])
    return dict(equity=ser, deployed=depser, trades=pd.DataFrame(trades), sleeves=sleeve_log, name=f'{policy}-Top{N}', start_year=start_year, scale=scale)

# ------------------------- metrics -------------------------
def max_dd(eq):
    peak = eq.cummax(); dd = eq / peak - 1; return dd
def dd_episodes(eq, top=5):
    dd = max_dd(eq); eps = []; in_dd = False
    for dt, v in dd.items():
        if v < 0 and not in_dd: in_dd = True; start = dt; trough = dt; tv = v
        elif v < 0 and in_dd and v < tv: trough = dt; tv = v
        elif v == 0 and in_dd: eps.append((start, trough, dt, tv)); in_dd = False
    if in_dd: eps.append((start, trough, None, tv))
    eps.sort(key=lambda x: x[3]); return eps[:top]

def metrics(res, nifty, rf=RF):
    eq = res['equity']; dep = res['deployed']; tr = res['trades']
    r = eq.pct_change(); r.iloc[0] = eq.iloc[0] / CAP0 - 1
    yrs = (eq.index[-1] - eq.index[0]).days / 365.25 if (eq.index[-1] - eq.index[0]).days > 0 else np.nan
    start_e = CAP0; end_e = eq.iloc[-1]; cagr = (end_e / start_e) ** (1 / yrs) - 1
    n = len(r); rf_d = rf / 252
    sd_d = r.std(ddof=1); vol = sd_d * np.sqrt(252); mean_ex = (r - rf_d).mean()
    sharpe = mean_ex / sd_d * np.sqrt(252) if sd_d > 0 else np.nan
    dside = np.sqrt(np.mean(np.minimum(0, r - rf_d) ** 2)); sortino = mean_ex / dside * np.sqrt(252) if dside > 0 else np.nan
    dd = max_dd(eq); mdd = dd.min(); calmar = cagr / abs(mdd) if mdd < 0 else np.nan
    # drawdown duration (longest time below a previous peak) and recovery of the max drawdown
    below = (dd < 0).astype(int).values; best = cur = 0; best_end = 0
    for i, v in enumerate(below):
        cur = cur + 1 if v else 0
        if cur > best: best = cur; best_end = i
    longest_dd_days = (eq.index[best_end] - eq.index[best_end - best + 1]).days if best else 0
    ulcer = np.sqrt(np.mean((dd * 100) ** 2))
    util = dep / eq; inmkt = dep > 0
    ret_in = r[inmkt]
    var95 = -np.percentile(ret_in, 5) * 100 if len(ret_in) > 20 else np.nan; cvar95 = -ret_in[ret_in <= np.percentile(ret_in, 5)].mean() * 100 if len(ret_in) > 20 else np.nan
    pnl_d = eq.diff().fillna(eq.iloc[0] - CAP0)
    avg_dep = dep.mean(); rod_ann = pnl_d.sum() / (dep.sum() / 252) / 1.0     # ₹ P&L per ₹ deployed-year -> annual return on deployed capital
    # same account but idle cash earns the risk-free rate (liquid fund)
    r_alt = r + (1 - util.clip(upper=1)) * rf_d; eq_alt = CAP0 * (1 + r_alt).cumprod()
    cagr_alt = (eq_alt.iloc[-1] / CAP0) ** (1 / yrs) - 1; sd_alt = r_alt.std(ddof=1); ex_alt = (r_alt - rf_d).mean()
    dsd_alt = np.sqrt(np.mean(np.minimum(0, r_alt - rf_d) ** 2)); dd_alt = max_dd(eq_alt).min()
    # monthly / yearly
    me = eq.resample('ME').last(); mr = me.pct_change(); mr.iloc[0] = me.iloc[0] / CAP0 - 1
    mon_in = dep.resample('ME').mean() > 0
    ye = eq.groupby(eq.index.year).last(); yr = ye.pct_change(); yr.iloc[0] = ye.iloc[0] / CAP0 - 1
    # versus nifty
    nr = nifty.reindex(eq.index).pct_change().fillna(0.0); nr.iloc[0] = 0.0
    nifty_cagr = (nifty.reindex(eq.index).iloc[-1] / nifty.reindex(eq.index).iloc[0]) ** (1 / yrs) - 1
    cov = np.cov(r.values, nr.values); beta = cov[0, 1] / cov[1, 1]; alpha_ann = (r.mean() - beta * nr.mean()) * 252
    corr = np.corrcoef(r.values, nr.values)[0, 1]
    up = nr > 0; dn = nr < 0
    upcap = r[up].mean() / nr[up].mean() if nr[up].mean() != 0 else np.nan; dncap = r[dn].mean() / nr[dn].mean() if nr[dn].mean() != 0 else np.nan
    act = r - nr; ir = act.mean() / act.std(ddof=1) * np.sqrt(252) if act.std(ddof=1) > 0 else np.nan
    nmdd = max_dd(nifty.reindex(eq.index)).min(); nvol = nr.std(ddof=1) * np.sqrt(252)
    nsharpe = ((nr - rf_d).mean() / nr.std(ddof=1)) * np.sqrt(252)
    # trades
    wins = tr[tr.NetPct > 0]; loss = tr[tr.NetPct <= 0]
    pf = wins.PnL.sum() / abs(loss.PnL.sum()) if len(loss) and loss.PnL.sum() != 0 else np.nan
    payoff = wins.NetPct.mean() / abs(loss.NetPct.mean()) if len(wins) and len(loss) else np.nan
    wr = len(wins) / len(tr) if len(tr) else np.nan
    kelly = wr - (1 - wr) / payoff if payoff and not np.isnan(payoff) else np.nan
    seq = (tr.sort_values(['Exit', 'Sleeve']).NetPct > 0).values; wstreak = lstreak = c1 = c2 = 0
    for v in seq:
        c1 = c1 + 1 if v else 0; c2 = c2 + 1 if not v else 0; wstreak = max(wstreak, c1); lstreak = max(lstreak, c2)
    top3 = tr.nlargest(3, 'PnL').PnL.sum() / tr.PnL.sum() * 100 if tr.PnL.sum() > 0 else np.nan
    # yearly bootstrap for CAGR
    rng = np.random.default_rng(7); yv = yr.values; B = []
    for _ in range(5000):
        s = rng.choice(yv, len(yv), replace=True); B.append(np.prod(1 + s) ** (1 / len(yv)) - 1)
    B = np.array(B)
    tstat_y = yv.mean() / (yv.std(ddof=1) / np.sqrt(len(yv))) if len(yv) > 2 and yv.std(ddof=1) > 0 else np.nan
    # PSR vs 0 (daily)
    srd = r.mean() / sd_d if sd_d > 0 else np.nan
    sk = ((r - r.mean()) ** 3).mean() / sd_d ** 3 if sd_d > 0 else 0; ku = ((r - r.mean()) ** 4).mean() / sd_d ** 4 if sd_d > 0 else 3
    return dict(
        Start=eq.index[0].date(), End=eq.index[-1].date(), Years=round(yrs, 2), StartCapital=CAP0, EndCapital=end_e, TotalReturnPct=(end_e / CAP0 - 1) * 100, CAGRPct=cagr * 100,
        RfCash_EndCapital=eq_alt.iloc[-1], RfCash_CAGRPct=cagr_alt * 100, RfCash_SharpeAnn=ex_alt / sd_alt * np.sqrt(252) if sd_alt > 0 else np.nan,
        RfCash_SortinoAnn=ex_alt / dsd_alt * np.sqrt(252) if dsd_alt > 0 else np.nan, RfCash_MaxDrawdownPct=dd_alt * 100, RfCash_VolatilityAnnPct=sd_alt * np.sqrt(252) * 100,
        RfCash_CalmarRatio=cagr_alt / abs(dd_alt) if dd_alt < 0 else np.nan, RfCash_ExcessCAGRvsNiftyPct=(cagr_alt - nifty_cagr) * 100,
        NiftyCAGRPct=nifty_cagr * 100, ExcessCAGRvsNiftyPct=(cagr - nifty_cagr) * 100,
        AvgFundUtilisedPct=util.mean() * 100, PeakFundUtilisedPct=util.max() * 100, AvgUtilisedWhenInvestedPct=util[inmkt].mean() * 100 if inmkt.any() else np.nan,
        DaysInMarketPct=inmkt.mean() * 100, IdleCashPct=100 - util.mean() * 100, AvgRupeesDeployed=avg_dep, ReturnOnDeployedCapitalAnnPct=rod_ann * 100,
        ReturnOnMargin20pct_Scenario=rod_ann * 100 / 0.20,
        TradesPerYear=len(tr) / len(yv), TotalTrades=len(tr), AvgHoldDays=tr.HoldDays.mean(), AvgOpenPositions=(tr.assign(x=1).groupby('Year').size().mean()),
        VolatilityAnnPct=vol * 100, SharpeAnn=sharpe, SortinoAnn=sortino, CalmarRatio=calmar, MaxDrawdownPct=mdd * 100, LongestDrawdownDays=longest_dd_days, UlcerIndex=ulcer,
        VaR95_DailyWhenInvestedPct=var95, CVaR95_DailyWhenInvestedPct=cvar95, BestDayPct=r.max() * 100, WorstDayPct=r.min() * 100,
        BestMonthPct=mr.max() * 100, WorstMonthPct=mr.min() * 100, PositiveMonthsInvestedPct=(mr[mon_in.reindex(mr.index).fillna(False)] > 0).mean() * 100,
        BestYearPct=yv.max() * 100, WorstYearPct=yv.min() * 100, PositiveYearsPct=(yv > 0).mean() * 100,
        Beta=beta, AlphaAnnPct=alpha_ann * 100, CorrelationNifty=corr, UpCapture=upcap, DownCapture=dncap, InformationRatio=ir,
        NiftyVolPct=nvol * 100, NiftyMaxDDPct=nmdd * 100, NiftySharpe=nsharpe,
        TradeWinPct=wr * 100, AvgTradeNetPct=tr.NetPct.mean(), MedianTradeNetPct=tr.NetPct.median(), AvgWinPct=wins.NetPct.mean(), AvgLossPct=loss.NetPct.mean(), PayoffRatio=payoff, ProfitFactor=pf,
        ExpectancyPer100=tr.NetPct.mean(), KellyFraction=kelly, LongestWinStreakTrades=wstreak, LongestLossStreakTrades=lstreak, Top3TradesShareOfProfitPct=top3,
        TotalCostRs=(tr.Stake * COST / 100).sum(), CostDragPctOfGrossProfit=((tr.Stake * COST / 100).sum() / (tr.Stake * tr.GrossPct / 100).sum() * 100) if (tr.Stake * tr.GrossPct / 100).sum() > 0 else np.nan,
        YearlyReturnTstat=tstat_y, CAGR_Boot5th_Pct=np.percentile(B, 5) * 100, CAGR_Boot95th_Pct=np.percentile(B, 95) * 100, ProbCAGRpositive_Pct=(B > 0).mean() * 100,
        _srd=srd, _skew=sk, _kurt=ku, _n=n)

def add_dsr(rows, n_trials):
    """Deflated Sharpe (Bailey & Lopez de Prado): probability the true Sharpe exceeds the best Sharpe expected from n_trials unskilled trials."""
    nd = NormalDist(); srs = np.array([r['_srd'] for r in rows]); sd_sr = srs.std(ddof=1); g = 0.5772156649
    sr0 = sd_sr * ((1 - g) * nd.inv_cdf(1 - 1 / n_trials) + g * nd.inv_cdf(1 - 1 / (n_trials * np.e)))
    for r in rows:
        sr, n, sk, ku = r['_srd'], r['_n'], r['_skew'], r['_kurt']
        denom = np.sqrt(max(1e-12, 1 - sk * sr + (ku - 1) / 4 * sr ** 2))
        r['DeflatedSharpeProb_Pct'] = nd.cdf((sr - sr0) * np.sqrt(n - 1) / denom) * 100
        r['ProbSharpePositive_PSR_Pct'] = nd.cdf(sr * np.sqrt(n - 1) / denom) * 100
    return sr0
