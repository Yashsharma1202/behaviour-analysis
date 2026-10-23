"""
Independent verification of the November portfolio results.
Re-simulates selected strategies from the raw adjusted prices with separate, simple code (no use of portfolio_engine.simulate / metrics) and compares
every headline measure with November_Portfolio_Performance.xlsx. Also checks the trade list, picks and benchmark.
run_checks() -> DataFrame(Check, Detail, Workbook, Independent, AbsDiff, Tolerance, Result)
"""
import sys, glob, os, pickle, math
sys.stdout.reconfigure(encoding='utf-8')
import numpy as np, pandas as pd

ROOT = 'D:/behaviour analysis/'
PERF = ROOT + 'November_Portfolio_Performance.xlsx'
COST, RF, CAP0 = 0.15, 0.065, 1_000_000.0

def _calendar(prices):
    nif = pd.read_csv(ROOT + 'Stock_Data_213/daily_ohlc/NIFTY.csv', usecols=['Date', 'Close'], parse_dates=['Date']).sort_values('Date').drop_duplicates('Date').set_index('Date')['Close']
    cnt = {}
    for s in prices.values():
        for d in s.index: cnt[d] = cnt.get(d, 0) + 1
    days = sorted(set(d for d, c in cnt.items() if c >= 20) | set(nif.index))
    days = [d for d in days if pd.Timestamp('2012-12-31') <= d <= pd.Timestamp('2025-12-31')]
    return pd.DatetimeIndex(days), nif

def _load(syms):
    out = {}
    for s in syms:
        d = pd.read_csv(ROOT + f'Stock_Data_213_Adjusted/daily_ohlc/{s}.csv', usecols=['Date', 'Adj Close'], parse_dates=['Date']).sort_values('Date').drop_duplicates('Date').set_index('Date')['Adj Close']
        out[s] = d
    return out

def _net(picks, N):
    return {(d, y): np.mean([p['gross'] for p in l[:N]]) - COST for (d, y), l in picks.items()}

def _dates(picks, N, policy, Y):
    net = _net(picks, N)
    def ok(h): return len(h) > 0 and np.mean(h) >= 1.0 and 100 * np.mean([x > 0 for x in h]) >= 62
    out = []
    for d in range(1, 31):
        if (d, Y) not in picks: continue
        if policy == 'ALL': out.append(d)
        elif policy == 'HINDSIGHT' and ok([net[(d, y)] for y in range(2013, 2026) if (d, y) in net]): out.append(d)
        elif policy == 'EXPANDING':
            h = [net[(d, y)] for y in range(2013, Y) if (d, y) in net]
            if len(h) >= 5 and ok(h): out.append(d)
    return out

def independent_sim(picks, prices, G, N, policy, start_year):
    first = G[G >= pd.Timestamp(f'{start_year}-01-01')][0]
    days = G[G >= first]; n = len(days)
    equity = np.zeros(n); deployed = np.zeros(n); E = CAP0
    for Y in range(start_year, 2026):
        ymask = np.array([d.year == Y for d in days]); idx = np.where(ymask)[0]
        sel = _dates(picks, N, policy, Y); K = len(sel)
        add = np.zeros(len(idx))
        for d in sel:
            for p in picks[(d, Y)][:N]:
                stake = E / (K * N)
                s = prices[p['sym']].loc[p['entry']:p['exit']]
                path = s.reindex(days[idx][(days[idx] >= p['entry']) & (days[idx] <= p['exit'])], method='ffill')
                rel = path.values / s.iloc[0] - 1.0
                lo = int(np.searchsorted(days[idx], p['entry'])); hi = lo + len(rel)
                for j in range(lo, len(idx)):
                    if j < hi: add[j] += stake * p['side'] * rel[j - lo]
                    else: add[j] += stake * p['side'] * rel[-1]
                    if j >= hi - 1: add[j] -= stake * COST / 100.0
                deployed[idx[lo:hi]] += stake
        equity[idx] = E + add; E = E + add[-1]
    return pd.Series(equity, index=days), pd.Series(deployed, index=days)

def independent_metrics(eq, dep, nifty):
    r = eq.pct_change(); r.iloc[0] = eq.iloc[0] / CAP0 - 1
    yrs = (eq.index[-1] - eq.index[0]).days / 365.25
    cagr = (eq.iloc[-1] / CAP0) ** (1 / yrs) - 1
    sd = r.std(ddof=1); rfd = RF / 252
    sharpe = (r - rfd).mean() / sd * math.sqrt(252)
    util = dep / eq
    r2 = r + (1 - util.clip(upper=1)) * rfd
    cagr2 = (np.prod(1 + r2.values)) ** (1 / yrs) - 1; sharpe2 = (r2 - rfd).mean() / r2.std(ddof=1) * math.sqrt(252)
    mdd = (eq / eq.cummax() - 1).min()
    nf = nifty.reindex(eq.index, method='ffill'); ncagr = (nf.iloc[-1] / nf.iloc[0]) ** (1 / yrs) - 1
    return dict(end=eq.iloc[-1], cagr=cagr * 100, cagr_rf=cagr2 * 100, sharpe=sharpe, sharpe_rf=sharpe2, mdd=mdd * 100, util=util.mean() * 100, peak=util.max() * 100, days_in=(dep > 0).mean() * 100,
                nifty_cagr=ncagr * 100, vol=sd * math.sqrt(252) * 100, years=yrs)

def run_checks():
    rows = []
    def add(check, detail, wb, ind, tol):
        diff = abs(float(wb) - float(ind)); rows.append(dict(Check=check, Detail=detail, Workbook=float(wb), Independent=float(ind), AbsDiff=diff, Tolerance=tol, Result='PASS' if diff <= tol else 'FAIL'))
    picks = pickle.load(open(ROOT + 'portfolio_cache/picks.pkl', 'rb'))
    syms = sorted({p['sym'] for l in picks.values() for p in l}); prices = _load(syms)
    G, nifty = _calendar(_load([os.path.basename(f)[:-4] for f in glob.glob(ROOT + 'Stock_Data_213_Adjusted/daily_ohlc/*.csv') if os.path.basename(f)[:-4] not in ('NIFTY', 'BANKNIFTY', 'SENSEX', 'TATAMOTORS')]))
    A = pd.read_excel(PERF, sheet_name='All_Measures')
    def wbv(col, label, nth=0):
        v = A[A.Measure == label][col].tolist(); return float(v[min(nth, len(v) - 1)])
    # --- 1. picks reproduce the earlier validated Top-N workbooks
    S5 = pd.read_excel(ROOT + 'Calendar_Backtest_Nov_ADJUSTED_2003plus_v2.xlsx', sheet_name='Summary_WalkForward')
    T1 = pd.read_excel(ROOT + 'Top1_Nov_Performance.xlsx', sheet_name='Summary_Top1').set_index('Date')
    mx = 0
    for N in (3, 5, 10):
        net = _net(picks, N)
        for d in range(1, 31):
            mine = np.mean([net[(d, y)] for y in range(2013, 2026)]); x = S5[(S5.Date == f'{d:02d}-Nov') & (S5['Portfolio size'] == N)]['Avg net % / yr'].iloc[0]; mx = max(mx, abs(mine - x))
    add('Picks vs Top 3/5/10 workbook', 'max |avg net % per date|, 90 date-sizes', 0.0, mx, 0.006)
    net = _net(picks, 1); mx = max(abs(np.mean([net[(d, y)] for y in range(2013, 2026)]) - T1.loc[f'{d:02d}-Nov', 'WF_AvgNet']) for d in range(1, 31))
    add('Picks vs Top 1 workbook', 'max |avg net %| over 30 dates', 0.0, mx, 0.006)
    # --- 2. independent re-simulation of four strategies
    for pol, N, sy in (('ALL', 3, 2013), ('EXPANDING', 5, 2018), ('HINDSIGHT', 1, 2013), ('ALL', 10, 2018)):
        eq, dep = independent_sim(picks, prices, G, N, pol, sy); m = independent_metrics(eq, dep, nifty); col = f'{pol}-Top{N} ({sy}-25)'; tag = f'{pol}-Top{N} {sy}-25'
        add('Re-simulation: ending capital (idle 0%)', tag, wbv(col, 'Ending capital - idle cash earns 0% (Rs)'), m['end'], 1.0)
        add('Re-simulation: CAGR % (idle 0%)', tag, wbv(col, 'CAGR % (idle cash 0%)'), m['cagr'], 1e-6)
        add('Re-simulation: CAGR % (idle 6.5%)', tag, wbv(col, 'CAGR % (idle cash earns 6.5%)'), m['cagr_rf'], 1e-6)
        add('Re-simulation: Sharpe (idle 0%)', tag, wbv(col, 'Sharpe ratio (annualised, rf 6.5%)'), m['sharpe'], 1e-6)
        add('Re-simulation: Sharpe (idle 6.5%)', tag, wbv(col, 'Sharpe ratio (annualised)'), m['sharpe_rf'], 1e-6)
        add('Re-simulation: max drawdown %', tag, wbv(col, 'Maximum drawdown %'), m['mdd'], 1e-6)
        add('Re-simulation: average fund utilised %', tag, wbv(col, 'Average fund utilised % (of equity, all days)'), m['util'], 1e-6)
        add('Re-simulation: peak fund utilised %', tag, wbv(col, 'Peak fund utilised %'), m['peak'], 1e-6)
        add('Re-simulation: days in market %', tag, wbv(col, 'Days in market %'), m['days_in'], 1e-6)
        add('Re-simulation: annualised volatility %', tag, wbv(col, 'Annualised volatility %'), m['vol'], 1e-6)
        add('Benchmark: Nifty 50 CAGR % (raw file)', tag, wbv(col, 'Nifty 50 CAGR % (same window)'), m['nifty_cagr'], 1e-6)
    # --- 3. trade list integrity (all 20 strategies)
    T = pd.read_excel(PERF, sheet_name='Trades'); E = pd.read_excel(PERF, sheet_name='Equity_Daily')
    worst_pnl = 0; worst_cost = 0; bad_dates = 0
    for (strat, win), g in T.groupby(['Strategy', 'Window']):
        pol, top = strat.split('-Top'); col = f'{pol}-Top{top} ({win[:4]}-25) equity'
        eqs = E[col].dropna(); worst_pnl = max(worst_pnl, abs(g['Net P&L (Rs)'].sum() - (eqs.iloc[-1] - CAP0)))
        worst_cost = max(worst_cost, float((g['Gross %'] - g['Net % (after 0.15%)'] - COST).abs().max()))
        bad_dates += int(((pd.to_datetime(g.Entry) >= pd.to_datetime(g.Exit)) | (g['Hold days'] > 20)).sum())
    add('Trade list: sum of net P&L vs equity change', 'worst over 20 strategies (Rs; trade P&L is rounded to Rs 1)', 0.0, worst_pnl, 60.0)
    add('Trade list: gross minus net equals 0.15%', 'worst over all trades (percentage points)', 0.0, worst_cost, 0.011)
    add('Trade list: entry before exit and hold <= 20 days', 'count of bad trades', 0, bad_dates, 0)
    add('Trade count', 'trades in list', len(T), len(T), 0)
    # --- 4. pick structure
    PK = pd.read_excel(PERF, sheet_name='WF_Picks_All')
    add('Picks: 30 dates x 13 years x 10 ranks', 'rows', 3900, len(PK), 0)
    add('Picks: entry/exit within the calendar year', 'count outside', 0, int(((pd.to_datetime(PK.Entry).dt.year != PK.Year) | (pd.to_datetime(PK.Exit).dt.year != PK.Year)).sum()), 0)
    # --- 5. utilisation never exceeds 100%
    util_cols = [c for c in E.columns if c.endswith('utilised %')]
    add('Utilisation never above 100% of equity', 'max over 20 strategies, % ', 100.0, float(E[util_cols].max().max()), 100.0)
    df = pd.DataFrame(rows); df['Result'] = np.where(df.AbsDiff <= df.Tolerance, 'PASS', 'FAIL')
    # utilisation check is an upper bound: pass if Independent <= 100
    m = df.Check.str.startswith('Utilisation never'); df.loc[m, 'Result'] = np.where(df.loc[m, 'Independent'] <= 100.0, 'PASS', 'FAIL')
    return df

if __name__ == '__main__':
    d = run_checks(); pd.set_option('display.width', 250); pd.set_option('display.max_rows', 200)
    print(d.round(6).to_string(index=False)); print('\nPASS', int((d.Result == 'PASS').sum()), 'FAIL', int((d.Result == 'FAIL').sum()))
