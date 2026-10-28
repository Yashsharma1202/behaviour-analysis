"""
Diwali_FO_All_Stocks_Data.json - ALL F&O stocks (213 incl. Nifty 50), all past Diwali data + each stock's best position by the scan rule + key details.
NO trade / no-trade decisions: no conviction labels, no filters. Every stock is listed with its facts; selection stats are reported as numbers only.
Per stock: lot / margin / data span, 26 yearly records (20 session adj+raw closes, 10x10 LONG-return grid), best LONG window, best SHORT window,
rule-selected position (higher win rate, tie flagged), walk-forward stats, per-year trade log at the selected window, 2026 entry/exit dates.
Additive only: writes this JSON; fo_diwali_json_validate.py re-checks it.
"""
import sys, json, os
sys.stdout.reconfigure(encoding='utf-8')
ROOT = 'D:/behaviour analysis/'
src = open(ROOT + 'build_diwali_christmas_perstock_reports_v4.py', encoding='utf-8').read()
head = src.split('# ---------------------------------------------------------------- per-stock analysis')[0].replace('json.dump(EVENT_WINDOW', '(lambda *a, **k: None)(EVENT_WINDOW').replace(r'open(ROOT + r"\Diwali_Christmas_event_windows_v4.json", "w")', 'open(__import__("os").devnull, "w")')  # DEVNULLFIX: never truncate the tracked windows file
ns = {}; exec(compile(head, 'v4', 'exec'), ns)
np, pd = ns['np'], ns['pd']
data = ns['data']['Christmas']; YEARS = ns['YEARS']; ANCH = ns['ANCH']; CAL = ns['CAL']; lots = ns['lots']; meta = ns['meta']; rep = ns['rep']; hol = ns['hol']
select = ns['select']; MIN_YEARS = ns['MIN_YEARS']; NIFTY50 = ns['NIFTY50']
SYMS = sorted(s for s in data)
px = pd.read_parquet(ROOT + 'Stock_Data_213/all_stocks_daily_ohlc.parquet', columns=['Symbol', 'Date', 'Close']).dropna()
raw = {s: g.drop_duplicates('Date').set_index('Date').Close for s, g in px.groupby('Symbol')}
dash = json.load(open(ROOT + 'dashboard_data/event_dashboard_data.json', encoding='utf-8'))
sector = {s['symbol']: s.get('sector', '') for s in [x for x in dash['holidays'] if x['id'] == 'diwali_2026'][0]['stocks']}
cal = CAL                                   # Christmas: no special session to remove; 25-Dec is a holiday so it is not in the market calendar
OFFS = [f'T-{n}' for n in range(8, 0, -1)] + [f'T+{m}' for m in range(1, 9)]
iso = lambda d: pd.Timestamp(d).strftime('%Y-%m-%d')
def fl(v, nd=6):
    if v is None: return None
    v = float(v); return round(v, nd) if np.isfinite(v) else None

events = []
for y in YEARS:
    T = np.datetime64(ANCH['Christmas'][y], 'D'); ip = int(np.searchsorted(cal, T, 'left')) - 1; iq = int(np.searchsorted(cal, T, 'right'))
    idx = list(range(ip - 7, ip + 1)) + list(range(iq, iq + 8))
    events.append(dict(year=y, anchor_date=iso(T), anchor_weekday=pd.Timestamp(T).strftime('%A'), sessions=[dict(offset=o, date=iso(cal[i])) for o, i in zip(OFFS, idx)]))
closed = set(hol[hol.Is_Trading_Holiday == 'YES'].Holiday_Date); closed.add(pd.Timestamp('2026-12-25'))
T26 = pd.Timestamp('2026-12-25'); c26 = [d for d in pd.bdate_range(T26 - pd.Timedelta(days=40), T26 + pd.Timedelta(days=40)) if d not in closed]
pre26 = [d for d in c26 if d < T26][-8:]; post26 = [d for d in c26 if d > T26][:8]
ev26 = dict(year=2026, anchor_date='2026-12-25', anchor_weekday='Friday', note='Christmas Day, NSE closed', sessions=[dict(offset=o, date=iso(d)) for o, d in zip(OFFS, pre26 + post26)])
TODAY = pd.Timestamp('2026-10-27')

def best_side(C, sg, minyrs):
    """best (lead, hold) for one side by win rate then avg signed return, needs >= minyrs valid years; returns dict or None"""
    v = ~np.isnan(C); n = v.sum(0)
    with np.errstate(all='ignore'):
        w = ((sg * C > 0) & v).sum(0) / n; avg = np.nanmean(sg * C, 0)
    sc = np.where(n >= minyrs, w + 1e-6 * avg, -1)
    if sc.max() < 0: return None
    li, hi = np.unravel_index(np.argmax(sc), sc.shape); x = sg * C[:, li, hi]; x = x[~np.isnan(x)]
    return dict(window=f'T-{li + 1} to T+{hi + 1}', lead=int(li + 1), hold=int(hi + 1), valid_years=int(n[li, hi]), win_rate_pct=fl(w[li, hi] * 100, 2), avg_signed_return_pct=fl(avg[li, hi], 3),
                median_signed_return_pct=fl(np.median(x), 3), best_year_pct=fl(x.max(), 3), worst_year_pct=fl(x.min(), 3))

def LOT(s):
    v = lots.get(s, np.nan); return int(v) if np.isfinite(v) else None
def CV(s):
    l, c = LOT(s), float(rep.LatestClose.get(s, np.nan))
    return fl(l * c, 2) if l and np.isfinite(c) else None
stocks = {}
for s in SYMS:
    D = data[s]; C = D['R'][:, :8, :8]; sg_ = None
    amap = {pd.Timestamp(d).strftime('%Y-%m-%d'): float(a) for d, a in zip(D['d'], D['a'])}; rmap = {pd.Timestamp(k).strftime('%Y-%m-%d'): float(v) for k, v in raw[s].items()}
    yrs = {}
    for yi, y in enumerate(YEARS):
        dts = [x['date'] for x in events[yi]['sessions']]
        yrs[str(y)] = dict(adj_close=[fl(amap.get(d), 10) for d in dts], raw_close=[fl(rmap.get(d), 2) for d in dts], long_return_pct_grid=[[fl(C[yi, li, hi]) for hi in range(8)] for li in range(8)])
    maxn = int((~np.isnan(C)).sum(0).max())
    bl, bs = best_side(C, 1, MIN_YEARS), best_side(C, -1, MIN_YEARS)
    sel = None; note = ''
    if bl is None and bs is None:
        note = f'fewer than {MIN_YEARS} valid Christmas years in any window (max {maxn}); no rule-based pick'
        low = [best_side(C, 1, min(3, maxn)) if maxn >= 3 else None, best_side(C, -1, min(3, maxn)) if maxn >= 3 else None]
        lo = max([x for x in low if x], key=lambda x: x['win_rate_pct'], default=None)
        low_side = None if lo is None else ('LONG' if low[0] is lo else 'SHORT')
        low_pick = None if lo is None else dict(side=low_side, **lo, note='low-sample indication only (3+ years)')
    else:
        low_pick = None
        gl = bl['win_rate_pct'] if bl else -1; gs = bs['win_rate_pct'] if bs else -1
        side = 'LONG' if gl > gs else ('SHORT' if gs > gl else 'TIE')
        use = bl if side in ('LONG', 'TIE') else bs; sgn = 1 if side in ('LONG', 'TIE') else -1
        sel = dict(side=side, side_rule='higher best-window win rate; exact tie reported as TIE and the LONG window is shown', **use)
        # per-year trade log at the selected window
        trades = []; li, hi = use['lead'] - 1, use['hold'] - 1
        for yi, y in enumerate(YEARS):
            r = C[yi, li, hi]
            if np.isnan(r): continue
            dts = [x['date'] for x in events[yi]['sessions']]; ed, xd = dts[8 - use['lead']], dts[7 + use['hold']]
            trades.append(dict(year=y, entry_date=ed, exit_date=xd, entry_raw_close=fl(rmap.get(ed), 2), exit_raw_close=fl(rmap.get(xd), 2), signed_return_pct=fl(sgn * r, 3), result='WIN' if sgn * r > 0 else 'LOSS'))
        sel['yearly_trades'] = trades
        l10 = [t['signed_return_pct'] for t in trades][-10:]; sel['last_10_win_rate_pct'] = fl(np.mean([v > 0 for v in l10]) * 100, 1) if l10 else None
        # walk-forward (pick redone each year from earlier years; same rule as the scan)
        oos = []
        for yi in range(MIN_YEARS, len(YEARS)):
            p = select(C[:yi])
            if p is None: continue
            r = (1 if p[0] == 'LONG' else -1) * C[yi, p[1], p[2]]
            if not np.isnan(r): oos.append(float(r))
        oos = np.array(oos)
        sel['walk_forward'] = dict(trades=int(len(oos)), win_rate_pct=fl((oos > 0).mean() * 100, 1) if len(oos) else None, avg_return_pct=fl(oos.mean(), 3) if len(oos) else None)
        e26, x26 = pre26[-use['lead']], post26[use['hold'] - 1]
        sel['dates_2026'] = dict(entry_date=iso(e26), exit_date=iso(x26), entry_date_already_passed=bool(e26 <= TODAY), note='as of 2026-10-27; entry at close of T-N, exit at close of T+M')
        sel['lot_size'] = LOT(s); sel['contract_value_1lot'] = CV(s); sel['margin_20pct_1lot'] = None if CV(s) is None else fl(0.2 * CV(s), 2)
        sel['futures_action'] = 'BUY FUTURES' if side in ('LONG', 'TIE') else 'SELL FUTURES'; sel['option_play'] = '1% ITM CALL' if side in ('LONG', 'TIE') else '1% ITM PUT'
    last = float(rep.LatestClose.get(s, np.nan))
    stocks[s] = dict(symbol=s, instrument_type='INDEX' if s.endswith('NIFTY') else 'STOCK', lot_size_missing=LOT(s) is None, last_close_missing=not np.isfinite(last), name=meta.get(s, {}).get('name', ''), is_nifty50=s in NIFTY50, sector=sector.get(s), lot_size=LOT(s), last_close=fl(last, 2), contract_value_1lot=CV(s),
                     margin_20pct_1lot=None if CV(s) is None else fl(0.2 * CV(s), 2), first_data_date=pd.Timestamp(D['d'][0]).strftime('%Y-%m-%d'), last_data_date=pd.Timestamp(D['d'][-1]).strftime('%Y-%m-%d'),
                     max_valid_years_any_window=maxn, valid_years_by_year={str(y): bool(not np.isnan(C[yi]).all()) for yi, y in enumerate(YEARS)},
                     best_long=bl, best_short=bs, selected_position=sel, low_sample_indication=low_pick, selection_note=note, years=yrs)

doc = dict(meta=dict(name='Christmas - all F&O stocks: past data, position taking and key details (project 8x8 grid)', generated='2026-10-27', stock_count=len(stocks), nifty50_count=sum(v['is_nifty50'] for v in stocks.values()),
                     event_count=len(events), years=YEARS,
                     disclaimer='No trade / no-trade decision is made in this file. selected_position is simply the scan rule output (best win-rate window); stats are facts, not recommendations. Windows were chosen on the same history they are scored on (in-sample); walk_forward is the out-of-sample view.',
                     conventions=dict(anchor='Christmas Day 25-Dec (T); NSE is closed on T',
                                      offsets='sessions[] ordered T-8..T-1, T+1..T+8 (index 0..15). Entry = close of T-N = sessions[8-N]; exit = close of T+M = sessions[7+M]',
                                      return_formula='LONG % = (adj_close[exit]/adj_close[entry]-1)*100; SHORT = -LONG; grid[N-1][M-1] = LONG % for entry T-N, exit T+M, null = invalid window that year',
                                      scan='PROJECT 8x8 GRID: lead 1..8 x hold 1..8 x LONG/SHORT; needs >= 7 valid years; pick by win rate then avg signed return; windows with >35% daily move, >40 calendar-day span or no session within 7 days of T are invalid',
                                      calendar='market day = >= 30% of listed stocks traded volume; prices: adj_close for returns, raw_close as quoted',
                                      sizing='contract_value_1lot = lot_size x last_close (today); margin = 20%'),
                     source_files=['build_diwali_christmas_perstock_reports_v4.py', 'Stock_Data_213/all_stocks_daily_ohlc.parquet', 'Stock_Data_213/lot_sizes.csv', 'dashboard_data/fo_stocks_211.json']),
           events=events, event_2026=ev26, stocks=stocks)
out = ROOT + 'Christmas_FO_All_Stocks_Data_8x8.json'
json.dump(doc, open(out, 'w', encoding='utf-8'), ensure_ascii=False, allow_nan=False, separators=(',', ':'))
print('written', out, f'{os.path.getsize(out)/1e6:.2f} MB', len(stocks), 'stocks')
