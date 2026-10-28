"""Validates Diwali_Past_Events_Backtest_Data_8x8.json: strict JSON, structure, dates, internal consistency, and a full re-computation of Book A from the file alone vs the report numbers."""
import sys, json, pickle, math
sys.stdout.reconfigure(encoding='utf-8')
ROOT = 'D:/behaviour analysis/'
def reject(c): raise ValueError('non-finite constant ' + c)
doc = json.load(open(ROOT + 'Diwali_Past_Events_Backtest_Data_8x8.json', encoding='utf-8'), parse_constant=reject)   # fails on NaN / Infinity
errs = []; chk = 0
def ok(c, msg):
    global chk; chk += 1
    if not c: errs.append(msg)
import datetime as dt
D = lambda s: dt.date.fromisoformat(s)
ev = doc['events']; YEARS = doc['meta']['years']
ok(len(ev) == 26 and [e['year'] for e in ev] == list(range(2000, 2026)) == YEARS, 'events should be 2000..2025')
OFFS = [f'T-{n}' for n in range(8, 0, -1)] + [f'T+{m}' for m in range(1, 9)]
for e in ev + [doc['event_2026']]:
    ds = [D(x['date']) for x in e['sessions']]; T = D(e['anchor_date'])
    ok([x['offset'] for x in e['sessions']] == OFFS, f"{e['year']} offsets")
    ok(all(a < b for a, b in zip(ds, ds[1:])), f"{e['year']} sessions not increasing")
    ok(ds[7] < T < ds[8], f"{e['year']} T not between T-1 and T+1")
    ok(all(d.weekday() < 5 for d in ds), f"{e['year']} weekend session")
    ok((ds[8] - ds[7]).days <= 8 and (T - ds[7]).days <= 7, f"{e['year']} gap around T too large")
ok(doc['event_2026']['anchor_date'] == '2026-11-08' and '2026-11-10' not in [x['date'] for x in doc['event_2026']['sessions']], '2026 calendar')
st = doc['stocks']; ok(len(st) == 50 and doc['meta']['stock_count'] == 50, 'stock count')
nvalid = 0; ngrid = 0; cons_bad = 0; cons = 0
for s, v in st.items():
    ok(v['lot_size'] > 0 and v['last_close'] and v['last_close'] > 0, f'{s} lot/close')
    ok(set(v['years']) == {str(y) for y in YEARS}, f'{s} years')
    for y, yy in v['years'].items():
        adj, rw, g = yy['adj_close'], yy['raw_close'], yy['long_return_pct_grid']
        ok(len(adj) == 16 and len(rw) == 16, f'{s} {y} path length'); ok(len(g) == 8 and all(len(r) == 8 for r in g), f'{s} {y} grid shape')
        ok(all(a is None or a > 0 for a in adj) and all(a is None or a > 0 for a in rw), f'{s} {y} non-positive price')
        has = any(c is not None for r in g for c in r); ok(has == yy['has_valid_window'], f'{s} {y} has_valid flag'); nvalid += has
        for li in range(8):
            for hi in range(8):
                c = g[li][hi]
                if c is None: continue
                ngrid += 1; e_, x_ = adj[8 - (li + 1)], adj[7 + (hi + 1)]
                if e_ is None or x_ is None: cons_bad += 1; continue
                cons += 1
                if abs((x_ / e_ - 1) * 100 - c) > 2e-3: cons_bad += 1
ok(cons_bad == 0, f'{cons_bad} grid cells disagree with adj_close paths (of {cons})')
# recompute Book A from the file only
R = pickle.load(open(ROOT + 'portfolio_cache/diwali_report_8x8.pkl', 'rb')); MA = R['M']['A']
trades = []
for p in doc['positions_2026']:
    s = st[p['symbol']]; cv = s['contract_value_1lot']; sg = 1 if p['direction'] == 'LONG' else -1
    for y in YEARS:
        c = s['years'][str(y)]['long_return_pct_grid'][p['lead'] - 1][p['hold'] - 1]
        if c is not None: trades.append((y, sg * c, sg * c / 100 * cv))
n = len(trades); wr = sum(t[1] > 0 for t in trades) / n * 100; net = sum(t[2] for t in trades)
ok(n == MA['trades'], f"trade count {n} vs report {MA['trades']}"); ok(abs(wr - MA['wr']) < 0.01, f"win rate {wr:.2f} vs {MA['wr']:.2f}"); ok(abs(net - MA['net']) < 1, f"net {net:,.0f} vs {MA['net']:,.0f}")
ok(len(doc['positions_2026']) == len(R['POS']) and all(p['lead'] <= 8 and p['hold'] <= 8 for p in doc['positions_2026']), 'positions count / grid bounds'); ok(all(p['symbol'] in st for p in doc['positions_2026']), 'position symbol missing')
ok(all(D(p['entry_date_2026']) >= D('2026-10-28') for p in doc['positions_2026']), 'a 2026 entry date is before 28-Oct')
ok(all(D(p['entry_date_2026']) < D('2026-11-08') < D(p['exit_date_2026']) for p in doc['positions_2026']), '2026 entry/exit not around T')
fut = [e['nifty_futures_T-5_to_T+4'] for e in ev]; ok(all(f is not None for f in fut), 'index futures trade missing for a year')
print(f'checks run: {chk}   valid stock-years: {nvalid}   grid cells verified: {ngrid}')
print(f'Book A recomputed from the JSON only: {n} trades, win rate {wr:.2f}%, net Rs {net:,.0f}  (report: {MA["trades"]}, {MA["wr"]:.2f}%, {MA["net"]:,.0f})')
if errs: print('ERRORS:', len(errs)); [print(' -', e) for e in errs[:30]]; sys.exit(1)
print('ALL CHECKS PASSED - 0 errors')
