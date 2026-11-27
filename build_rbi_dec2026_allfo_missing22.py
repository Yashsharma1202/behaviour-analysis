"""
ADDITIVE: computes the RBI 04-Dec-2026 8x8 windows for the 22 F&O stocks that the original All-F&O build dropped
(7 had no price rows in the parquet -> read from Stock_Data_213/new_fo_additions_20Nov2026; 15 are recent listings with
<30 RBI events -> computed with a lower event floor and flagged LOW SAMPLE). Same grid/selection logic as build_rbi_dec2026_allfo_index.py.
Writes RBI_Dec2026_AllFO_Missing22.json only.
"""
import json, glob, os
import numpy as np, pandas as pd
ROOT = 'D:/behaviour analysis/'; T_UP = pd.Timestamp('2026-12-04'); K = 8; MIN_EV = 30; MIN_LOW = 3
ev = {pd.Timestamp(e['eff_date']) for e in json.load(open(ROOT + 'dashboard_data/rbi_policy_behaviour.json'))['all_events'] if e['category'] != 'OTHER'}
ev |= {pd.Timestamp(c['date_str']) for c in json.load(open(ROOT + 'processed/rbi_mpc_calendar.json')) if c['status'] == 'confirmed'}
ev = sorted(d for d in ev if d < T_UP)
px = pd.read_parquet(ROOT + 'Stock_Data_213/all_stocks_daily_ohlc.parquet', columns=['Symbol', 'Date', 'Close', 'Adj Close'])
extra = pd.concat([pd.read_csv(f, usecols=['Symbol', 'Date', 'Close', 'Adj Close']) for f in glob.glob(ROOT + 'Stock_Data_213/new_fo_additions_20Nov2026/*.csv')])
px = pd.concat([px, extra]).dropna(subset=['Close'])
px['Date'] = pd.to_datetime(px['Date']).dt.normalize(); px = px.drop_duplicates(['Symbol', 'Date'])
A = px.pivot(index='Date', columns='Symbol', values='Adj Close').sort_index(); R = px.pivot(index='Date', columns='Symbol', values='Close').sort_index()
base = px[~px.Symbol.isin(extra.Symbol.unique())]
cal = A.index[(A[[c for c in A.columns if c not in set(extra.Symbol)]].notna().sum(1) / (A.shape[1] - extra.Symbol.nunique())) >= 0.30]
A, R = A.loc[cal], R.loc[cal]
have = {r['symbol'] for r in json.load(open(ROOT + 'dashboard_data/event_dashboard_data.json'))['rbi_policy'][0]['all_stocks']}
fo = pd.read_csv(ROOT + 'FO_Stocks_Updated_20Nov2026.csv'); lots = fo.set_index('symbol').lot_size_nov26
miss = [s for s in fo.symbol if s not in have]
idx = []
for d in ev:
    t = int(cal.searchsorted(d, 'left'))
    if t < len(cal) and (cal[t] - d).days <= 4 and t - K >= 0 and t + K < len(cal): idx.append(t)
def cube(sym):
    a = A[sym].values; C = np.full((len(idx), K, K), np.nan)
    for k, t in enumerate(idx):
        for i in range(K):
            e = a[t - i - 1]
            for j in range(K):
                x = a[t + j + 1]
                if np.isfinite(e) and np.isfinite(x) and e > 0 and abs(x / e - 1) < 0.6: C[k, i, j] = (x / e - 1) * 100
    return C
out = []
for s in miss:
    if s not in A.columns: out.append(dict(symbol=s, status='NO PRICE DATA')); continue
    C = cube(s); v = ~np.isnan(C); n = v.sum(0); nmax = int(n.max())
    if nmax < MIN_LOW: out.append(dict(symbol=s, status=f'NO WINDOW ({nmax} events)')); continue
    floor = MIN_EV if nmax >= MIN_EV else nmax
    with np.errstate(all='ignore'):
        wl = ((C > 0) & v).sum(0) / n; ws = ((C < 0) & v).sum(0) / n; av = np.nanmean(C, 0)
    sl = np.where(n >= floor, wl + 1e-6 * av, -1); ss = np.where(n >= floor, ws - 1e-6 * av, -1)
    li = np.unravel_index(np.argmax(sl), sl.shape); sj = np.unravel_index(np.argmax(ss), ss.shape)
    if sl[li] >= ss[sj]: side, (a_, b_), wr, ar, nn = 'LONG', li, wl[li] * 100, av[li], int(n[li])
    else: side, (a_, b_), wr, ar, nn = 'SHORT', sj, ws[sj] * 100, -av[sj], int(n[sj])
    lc = float(R[s].dropna().iloc[-1]); lot = lots.get(s)
    out.append(dict(symbol=s, status='OK' if floor == MIN_EV else 'LOW SAMPLE', direction=side, window=f'T-{a_+1} to T+{b_+1}',
                    win_rate=round(float(wr), 1), events=nn, win_ratio=f'{round(wr * nn / 100)} Wins / {nn} Events', avg_ret=round(float(ar), 3),
                    lot=None if pd.isna(lot) else int(lot), margin=None if pd.isna(lot) else round(lot * lc * 0.2)))
json.dump(out, open(ROOT + 'RBI_Dec2026_AllFO_Missing22.json', 'w'), indent=1)
print(len(miss), 'missing;', pd.Series([o['status'].split(' (')[0] for o in out]).value_counts().to_dict(), '| events', len(idx))
for o in out: print(o)
