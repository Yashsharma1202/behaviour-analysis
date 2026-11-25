"""Backtest ONLY the 7 stocks newly added to NSE F&O (Nov-26). Same T-n/T+m grid search as
compute_real_fo211_nonnifty50_data.py, but on SPOT closes (no futures series exists yet)."""
import json, sys
from datetime import datetime
import pandas as pd
sys.path.insert(0,'.')
from compute_real_fo211_nonnifty50_data import best_direction_and_window
NEW="COROMANDEL CUB DATAPATTNS ICICILIFE KARURVYSYA NAVINFLUOR SYRMA".split()
data=json.load(open('dashboard_data/fo_stocks_211.json'))
by={x['symbol']:x for x in data}
for sym in NEW:
    px=pd.read_csv(f'Stock_Data_213/new_fo_additions_20Nov2026/{sym}.csv')
    a=[[datetime.strptime(d,'%Y-%m-%d').date(),c] for d,c in zip(px.Date,px['Adj Close'])]
    res=json.load(open(f'new_fo_results_20Nov2026/{sym}_results.json'))
    dates=sorted({datetime.strptime(x['broadCastDate'][:11],'%d-%b-%Y').date() for x in res if x.get('broadCastDate')})
    dates=[d for d in dates if d<datetime(2026,11,20).date()]
    side,win,score=best_direction_and_window(a,dates,min_n=3)
    s=by[sym]
    if side is None: print(sym,'insufficient',len(dates)); continue
    bo,so=win; wr,avg=score; spot=s['spot_ltp']; lot=s['lot_size']
    s.update(taking_window_raw=f'T-{bo} to T+{so}',entry_lead_days=bo,exit_hold_days=so,direction=side,
      q_win_rate=round(wr,1),h_win_rate=round(wr,1),expected_ret=round(avg,2),h_est_pnl=round(avg/100*spot*lot,2),
      best_strategy=f'Pre-Quarterly Run-up ({side})',n_hist_quarters=len(dates),low_confidence=len(dates)<8,
      data_source=f'Spot-price backtest (no futures series yet): {len(dates)} historical result dates, {bo}/{so} T-n/T+m grid search, direction by win-rate then avg return',
      futures_data_range=None)
    print(sym,side,f'T-{bo}/T+{so}',round(wr,1),round(avg,2),len(dates))
json.dump(data,open('dashboard_data/fo_stocks_211.json','w'),indent=1)
open('dashboard_data/fo_stocks_211.js','w').write('window.FO_STOCKS_211 = '+json.dumps(data)+';\n')
