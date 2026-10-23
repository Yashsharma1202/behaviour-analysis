"""
Runs the November Top-N portfolio simulation end to end and caches the results for the Excel / report builders.
  1. loads adjusted prices, builds the walk-forward picks (30 dates x 2013-2025, Top 10 each)
  2. simulates 20 strategies (ALL / HINDSIGHT / EXPANDING date policies x Top 1,3,5,10), the 'scaled' variants and the cost-sensitivity runs
  3. computes every performance measure and writes portfolio_cache/{picks,sims}.pkl
Usage: python run_portfolio_pipeline.py     (takes about 1 minute)
"""
import sys, pickle, time, pathlib
sys.path.insert(0, 'D:/behaviour analysis'); sys.stdout.reconfigure(encoding='utf-8')
import numpy as np, pandas as pd
import portfolio_engine as pe

CACHE = pathlib.Path('D:/behaviour analysis/portfolio_cache'); CACHE.mkdir(exist_ok=True)
t0 = time.time()
PX = pe.load_prices(); print('stocks', len(PX))
picks = pe.build_picks(PX)
pickle.dump(picks, open(CACHE / 'picks.pkl', 'wb')); print('picks', len(picks))
G = pe.make_calendar(PX); print('calendar', G.index[0].date(), G.index[-1].date(), len(G))
specs = []
for N in pe.SIZES: specs += [('ALL', N, 2013), ('HINDSIGHT', N, 2013), ('EXPANDING', N, 2018), ('ALL', N, 2018), ('HINDSIGHT', N, 2018)]
sims, scaled, sens = {}, {}, {}
for pol, N, sy in specs:
    r = pe.simulate(PX, G, picks, N, pol, sy); sims[(pol, N, sy)] = r
    peak = (r['deployed'] / r['equity']).max(); sc = 0.98 / peak
    scaled[(pol, N, sy)] = pe.simulate(PX, G, picks, N, pol, sy, scale=sc); scaled[(pol, N, sy)]['scale'] = sc
for pol, N, sy in [('ALL', n, 2013) for n in pe.SIZES] + [('EXPANDING', n, 2018) for n in pe.SIZES] + [('HINDSIGHT', n, 2013) for n in pe.SIZES]:
    for c in (0.30, 0.50): sens[(pol, N, sy, c)] = pe.simulate(PX, G, picks, N, pol, sy, cost=c)
bad = 0
for k, r in list(sims.items()) + list(scaled.items()):
    tr = r['trades']; E = pe.CAP0
    for Y in range(k[2], 2026):
        ye = r['equity'][r['equity'].index.year == Y].iloc[-1]; E += tr[tr.Year == Y].PnL.sum()
        if abs(ye - E) > 0.5: bad += 1
print('integrity: year-end equity vs sum of trade P&L mismatches =', bad)
assert bad == 0
def mt(d):
    out = []
    for k, r in d.items():
        m = pe.metrics(r, G); m.update(Strategy=f'{k[0]}-Top{k[1]}', Policy=k[0], N=k[1], Window=f'{k[2]}-2025')
        if len(k) == 4: m['Cost'] = k[3]
        out.append(m)
    return out
rows, rows_sc, rows_se = mt(sims), mt(scaled), mt(sens)
for rs in (rows, rows_sc): pe.add_dsr(rs, 20)
pickle.dump((sims, scaled, sens, rows, rows_sc, rows_se, G), open(CACHE / 'sims.pkl', 'wb'))
print('done in', round(time.time() - t0), 's ->', CACHE)
