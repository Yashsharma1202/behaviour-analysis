"""Diwali 8x8: position list JSON (all 50 Nifty, status + analysis), low-history experimental JSON, and the Diwali master JSON. Reads portfolio_cache/diwali_report_8x8.pkl. Additive only."""
import sys, json, pickle
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
ROOT = 'D:/behaviour analysis/'
R = pickle.load(open(ROOT + 'portfolio_cache/diwali_report_8x8.pkl', 'rb'))
pre, post = R['pre26'], R['post26']; lots = R['lots']; nt = R['notional']
iso = lambda d: pd.Timestamp(d).strftime('%Y-%m-%d'); lng = lambda d: pd.Timestamp(d).strftime('%d-%b-%Y (%A)')
stocks = []
for rank, r in enumerate(sorted(R['ALL50'], key=lambda x: (x['status'] != 'TRADE', -(x['analysis']['wr'] if x['analysis'] else -1))), 1):
    a = r['analysis']; row = dict(rank=rank, symbol=r['symbol'], name=r['name'], sector=r['sector'], status=r['status'], lot_size=int(lots[r['symbol']]) if r['symbol'] in lots else None)
    if a:
        e, x = pre[-a['lead']], post[a['hold'] - 1]
        row.update(direction=a['direction'], window=f"T-{a['lead']} to T+{a['hold']}", lead=a['lead'], hold=a['hold'], entry_date=lng(e), exit_date=lng(x), entry_date_iso=iso(e), exit_date_iso=iso(x),
                   valid_years=a['n'], win_rate_pct=a['wr'], avg_signed_return_pct=a['avg'], median_pct=a['median'], best_year_pct=a['best'], worst_year_pct=a['worst'], last_10_win_rate_pct=a['last10'],
                   walk_forward_win_rate_pct=a['wf'], walk_forward_trades=a['wf_n'], conviction=a['conv'],
                   contract_value_1lot=round(nt[r['symbol']], 2), margin_20pct_1lot=round(0.2 * nt[r['symbol']], 2),
                   futures_action='BUY FUTURES' if a['direction'] == 'LONG' else 'SELL FUTURES', option_play='1% ITM CALL' if a['direction'] == 'LONG' else '1% ITM PUT')
    else:
        row.update(conviction='DATA INSUFFICIENT', note='fewer than 7 valid Diwali years in any 8x8 window')
    stocks.append(row)
n_tr = sum(s['status'] == 'TRADE' for s in stocks)
plist = dict(event='Diwali 2026 (Laxmi Pujan, Sunday 08-Nov-2026)', grid='PROJECT 8x8 GRID (entry T-N, exit T+M, N and M = 1..8)', generated='2026-10-27', scope='Nifty 50 only',
             rule='Per-stock best side + window by win rate then avg return, >=7 valid years; STRONG/MODERATE conviction gate (walk-forward), same rules as the v4 scan but on the 8x8 grid',
             total_stocks=len(stocks), trade_count=n_tr, strong=sum(s.get('conviction') == 'STRONG' for s in stocks), moderate=sum(s.get('conviction') == 'MODERATE' for s in stocks),
             no_trade=sum(s['status'] == 'NO TRADE' for s in stocks), data_insufficient=sum(s['status'] == 'DATA INSUFFICIENT' for s in stocks),
             note='On the 8x8 grid the earliest entry is T-8 = 2026-10-28, so no entry date has already passed (no re-pick needed).', stocks=stocks)
json.dump(plist, open(ROOT + 'Diwali_2026_Nifty50_Position_List_8x8.json', 'w', encoding='utf-8'), indent=2, ensure_ascii=False, allow_nan=False)
exp = dict(event='Diwali 2026 - positions for 3 low-history Nifty stocks (8x8 grid)', event_date='08-Nov-2026 (Sunday)', generated='2026-10-27',
           warning='These stocks have fewer than the 7 valid Diwali years required for a call. Picks come from 3-6 data points and are indicative only.', stocks=R['EXP'])
json.dump(exp, open(ROOT + 'Diwali_2026_3_LowHistory_Stocks_Positions_8x8.json', 'w', encoding='utf-8'), indent=2, ensure_ascii=False, allow_nan=False)
print('position list:', n_tr, 'trade;', plist['strong'], 'strong;', plist['no_trade'], 'no trade;', plist['data_insufficient'], 'insufficient')
