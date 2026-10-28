import sys, json, datetime as dt
sys.stdout.reconfigure(encoding='utf-8')
def rej(c): raise ValueError('non-finite ' + c)
doc = json.load(open('D:/behaviour analysis/Christmas_FO_All_Stocks_Data_8x8.json', encoding='utf-8'), parse_constant=rej)
errs = []; n = 0
def ok(c, m):
    global n; n += 1
    if not c: errs.append(m)
D = dt.date.fromisoformat; YEARS = doc['meta']['years']; ev = doc['events']
ok(len(ev) == 26 and [e['year'] for e in ev] == list(range(2000, 2026)), 'events')
for e in ev + [doc['event_2026']]:
    ds = [D(x['date']) for x in e['sessions']]; T = D(e['anchor_date'])
    ok(len(ds) == 16 and all(a < b for a, b in zip(ds, ds[1:])) and ds[7] < T < ds[8] and all(d.weekday() < 5 for d in ds), f"event {e['year']}")
st = doc['stocks']; ok(len(st) == doc['meta']['stock_count'] == 213, f'stock count {len(st)}'); ok(sum(v['is_nifty50'] for v in st.values()) == 50, 'nifty count')
cells = bad = sel_n = 0; side_ct = {}; lowc = 0; nosel = 0
for s, v in st.items():
    ok(set(v['years']) == {str(y) for y in YEARS}, f'{s} years'); ok((v['lot_size'] is None) == v['lot_size_missing'] and (v['lot_size'] is None or v['lot_size'] > 0), f'{s} lot')
    ok('decision' not in v and 'conviction' not in json.dumps(v['selected_position'] or {}), f'{s} contains a decision field')
    grids = {}
    for y, yy in v['years'].items():
        a, g = yy['adj_close'], yy['long_return_pct_grid']; grids[int(y)] = g
        ok(len(a) == 16 and len(yy['raw_close']) == 16 and len(g) == 8 and all(len(r) == 8 for r in g), f'{s} {y} shape')
        ok(all(x is None or x > 0 for x in a), f'{s} {y} price')
        for li in range(8):
            for hi in range(8):
                c = g[li][hi]
                if c is None: continue
                cells += 1; e_, x_ = a[7 - li], a[8 + hi]
                if e_ is None or x_ is None or abs((x_ / e_ - 1) * 100 - c) > 2e-3: bad += 1
    sp = v['selected_position']
    if sp is None:
        nosel += 1; ok(v['max_valid_years_any_window'] < 7, f'{s} no pick but >=7 years'); continue
    sel_n += 1; side_ct[sp['side']] = side_ct.get(sp['side'], 0) + 1; sg = -1 if sp['side'] == 'SHORT' else 1
    vals = [sg * grids[y][sp['lead'] - 1][sp['hold'] - 1] for y in sorted(grids) if grids[y][sp['lead'] - 1][sp['hold'] - 1] is not None]
    ok(len(vals) == sp['valid_years'] == len(sp['yearly_trades']) and sp['valid_years'] >= 7 and sp['lead'] <= 8 and sp['hold'] <= 8, f'{s} valid years / grid bounds')
    ok(abs(sum(x > 0 for x in vals) / len(vals) * 100 - sp['win_rate_pct']) < 0.01, f'{s} win rate recompute')
    ok(abs(sum(vals) / len(vals) - sp['avg_signed_return_pct']) < 2e-3, f'{s} avg recompute')
    ok(all(t['result'] == ('WIN' if t['signed_return_pct'] > 0 else 'LOSS') for t in sp['yearly_trades']), f'{s} result flags')
    d26 = sp['dates_2026']; ok(D(d26['entry_date']) < D('2026-12-25') < D(d26['exit_date']), f'{s} 2026 dates')
    ok(sp['lead'] == int(sp['window'].split()[0][2:]) and sp['hold'] == int(sp['window'].split()[-1][2:]), f'{s} window label')
ok(bad == 0, f'{bad} grid cells disagree with adj paths of {cells}')
print('lot size missing:', [s for s,x in st.items() if x['lot_size_missing']], '| index instruments:', [s for s,x in st.items() if x['instrument_type']=='INDEX'])
print(f'checks {n}, grid cells verified {cells}, stocks with a rule-based pick {sel_n}, without {nosel}, sides {side_ct}')
if errs: print('ERRORS', len(errs)); [print(' -', e) for e in errs[:25]]; sys.exit(1)
print('ALL CHECKS PASSED - 0 errors')
