"""
Diwali_Past_Events_Backtest_Data.json - everything needed to backtest the Diwali strategy on every past Diwali (2000-2025) + the 2026 calendar.
Per event: Laxmi Pujan date T and the 20 sessions T-10..T-1, T+1..T+10 (Muhurat session excluded, market-day calendar).
Per Nifty-50 stock and event: adj close + raw close for those 20 sessions, and the 10x10 LONG-return grid (lead x hold) with the
v4 validity rules already applied (null = window not valid that year).  Strict JSON (no NaN / Infinity).
Additive only: writes the JSON and diwali_backtest_data_validate.py reads it back.
"""
import sys, json, pickle
sys.stdout.reconfigure(encoding='utf-8')
ROOT = 'D:/behaviour analysis/'
src = open(ROOT + 'build_diwali_christmas_perstock_reports_v4.py', encoding='utf-8').read()
head = src.split('# ---------------------------------------------------------------- per-stock analysis')[0].replace('json.dump(EVENT_WINDOW', '(lambda *a, **k: None)(EVENT_WINDOW').replace(r'open(ROOT + r"\Diwali_Christmas_event_windows_v4.json", "w")', 'open(__import__("os").devnull, "w")')  # DEVNULLFIX: never truncate the tracked windows file
ns = {}; exec(compile(head, 'v4', 'exec'), ns)
np, pd = ns['np'], ns['pd']
data = ns['data']['Diwali']; YEARS = ns['YEARS']; ANCH = ns['ANCH']; CAL = ns['CAL']; lots = ns['lots']; meta = ns['meta']; rep = ns['rep']; hol = ns['hol']
NIFTY50 = ns['NIFTY50']; N50 = sorted(s for s in NIFTY50 if s in data)
R = pickle.load(open(ROOT + 'portfolio_cache/diwali_report_8x8.pkl', 'rb'))
px = pd.read_parquet(ROOT + 'Stock_Data_213/all_stocks_daily_ohlc.parquet', columns=['Symbol', 'Date', 'Close']).dropna()
raw = {s: g.drop_duplicates('Date').set_index('Date').Close for s, g in px.groupby('Symbol') if s in N50}
muh = [np.datetime64(ANCH['Diwali'][y], 'D') for y in YEARS]
cal = CAL[~np.isin(CAL, muh)]                                           # market days, Muhurat sessions removed (same rule as the scan)
OFFS = [f'T-{n}' for n in range(8, 0, -1)] + [f'T+{m}' for m in range(1, 9)]
iso = lambda d: pd.Timestamp(d).strftime('%Y-%m-%d')
fl = lambda v, nd=6: None if v is None or (isinstance(v, float) and not np.isfinite(v)) else round(float(v), nd)

events = []; sess_idx = {}
for yi, y in enumerate(YEARS):
    T = np.datetime64(ANCH['Diwali'][y], 'D'); ip = int(np.searchsorted(cal, T, 'left')) - 1; iq = int(np.searchsorted(cal, T, 'right'))
    idx = list(range(ip - 7, ip + 1)) + list(range(iq, iq + 8)); sess_idx[y] = idx
    d = [iso(cal[i]) for i in idx]
    events.append(dict(year=y, anchor_date=iso(T), anchor_weekday=pd.Timestamp(T).strftime('%A'), anchor_is_market_day=bool(T in CAL),
                       sessions=[dict(offset=o, date=dt) for o, dt in zip(OFFS, d)]))
# index futures (independent check), keyed by year
fh = R['fh']; idx_tr = {t['year']: t for t in fh['trades']}
for e in events:
    t = idx_tr.get(e['year'])
    e['nifty_futures_T-5_to_T+4'] = None if not t else dict(entry_date=pd.to_datetime(t['entry_date'][:11], format='%d-%b-%Y').strftime('%Y-%m-%d'), entry_price=t['entry_price'],
                                                           exit_date=pd.to_datetime(t['exit_date'][:11], format='%d-%b-%Y').strftime('%Y-%m-%d'), exit_price=t['exit_price'], ret_pct=t['ret_pct'], pnl_rs_1lot=t['pnl_rs'])

dash = json.load(open(ROOT + 'dashboard_data/event_dashboard_data.json', encoding='utf-8'))
sector = {s['symbol']: s.get('sector', '') for s in [x for x in dash['holidays'] if x['id'] == 'diwali_2026'][0]['stocks']}
stocks = {}
for s in N50:
    D = data[s]; amap = {pd.Timestamp(d).strftime('%Y-%m-%d'): float(a) for d, a in zip(D['d'], D['a'])}; rs = raw[s]
    rmap = {pd.Timestamp(k).strftime('%Y-%m-%d'): float(v) for k, v in rs.items()}
    yrs = {}
    for yi, y in enumerate(YEARS):
        dts = [x['date'] for x in events[yi]['sessions']]
        adj = [fl(amap.get(d), 10) for d in dts]; rw = [fl(rmap.get(d), 2) for d in dts]
        grid = [[fl(D['R'][yi, li, hi], 6) for hi in range(8)] for li in range(8)]
        has = any(v is not None for row in grid for v in row)
        yrs[str(y)] = dict(has_valid_window=has, adj_close=adj, raw_close=rw, long_return_pct_grid=grid)
    first = pd.Timestamp(D['d'][0]).strftime('%Y-%m-%d')
    stocks[s] = dict(name=meta.get(s, {}).get('name', ''), sector=sector.get(s, ''), lot_size=int(lots[s]), last_close=fl(float(rep.LatestClose[s]), 2), first_data_date=first,
                     contract_value_1lot=fl(R['notional'][s], 2), margin_20pct_1lot=fl(0.2 * R['notional'][s], 2), years=yrs)

# 2026 calendar + positions
closed = set(hol[hol.Is_Trading_Holiday == 'YES'].Holiday_Date); closed.add(pd.Timestamp('2026-11-10'))
T26 = pd.Timestamp('2026-11-08'); c26 = [d for d in pd.bdate_range(T26 - pd.Timedelta(days=40), T26 + pd.Timedelta(days=40)) if d not in closed]
pre26 = [d for d in c26 if d < T26][-8:]; post26 = [d for d in c26 if d > T26][:8]
ev26 = dict(year=2026, anchor_date='2026-11-08', anchor_weekday='Sunday', anchor_is_market_day=False, note='Laxmi Pujan / Muhurat on a Sunday; 2026-11-10 (Balipratipada) added as a market holiday',
            sessions=[dict(offset=o, date=iso(d)) for o, d in zip(OFFS, pre26 + post26)])
POS = []
for p in R['POS']:
    lead, hold = int(p['lead']), int(p['hold'])
    POS.append(dict(symbol=p['symbol'], name=p['name'], direction=p['direction'], window=f'T-{lead} to T+{hold}', lead=lead, hold=hold, entry_date_2026=iso(pre26[-lead]), exit_date_2026=iso(post26[hold - 1]),
                    win_rate_26y_pct=p['wr'], avg_signed_ret_pct=p['avg'], walk_forward_win_rate_pct=p['wf'], walk_forward_trades=p['wf_n'], conviction=p['conv'], re_picked=bool(p['repick']),
                    sector=p['sector'], lot_size=int(lots[p['symbol']])))
doc = dict(
    meta=dict(name='Diwali past-events backtest data', generated='2026-10-27', scope='Nifty 50 (50 stocks) x 26 past Diwali events (2000-2025) + 2026 calendar',
              conventions=dict(anchor='Laxmi Pujan / Muhurat date T (hand-corrected list, verified against price data); the Muhurat session itself is NOT a trading day',
                               offsets='T-1 = last market session before T, T+1 = first after T. PROJECT 8x8 GRID. sessions[] is ordered T-8..T-1, T+1..T+8 (index 0..15). Entry at close of T-N = sessions[8-N]; exit at close of T+M = sessions[7+M]',
                               return_formula='LONG return % = (adj_close[exit]/adj_close[entry]-1)*100 ; SHORT = -LONG',
                               grid='long_return_pct_grid[N-1][M-1] (8x8, N and M = 1..8) = LONG % return for entry T-N, exit T+M. null = window invalid that year (stock not listed, >7 days from T, >40 days span, or a >35% daily move)',
                               calendar='Market day = at least 30% of listed stocks traded volume that day (removes placeholder rows on holidays)',
                               prices='adj_close drives returns (split/dividend safe); raw_close is the actually quoted close for ledgers; null where the stock has no data that day',
                               sizing='contract_value_1lot = lot_size x last_close (today); margin = 20% of it; same Rs weight for every year',
                               cost_assumption_pct_round_trip=0.15),
              years=YEARS, stock_count=len(stocks), event_count=len(events), source_files=['build_diwali_christmas_perstock_reports_v4.py', 'Stock_Data_213/all_stocks_daily_ohlc.parquet',
              'dashboard_data/nifty_futures_holiday_behaviour.json', 'Diwali_2026_Nifty50_Position_List_8x8.json']),
    events=events, stocks=stocks, event_2026=ev26, positions_2026=POS)
out = ROOT + 'Diwali_Past_Events_Backtest_Data_8x8.json'
json.dump(doc, open(out, 'w', encoding='utf-8'), ensure_ascii=False, allow_nan=False, separators=(',', ':'))
import os; print('written', out, f'{os.path.getsize(out)/1e6:.2f} MB', len(events), 'events', len(stocks), 'stocks', len(POS), 'positions')
