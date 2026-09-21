import pandas as pd
import numpy as np
import pathlib
import sys
import time

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
SOURCE_FUT = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'

xl_src = pd.ExcelFile(SOURCE_FUT, engine='openpyxl')
df_src_trades = xl_src.parse('All_12Q_Futures_Trades')

def round_nse_strike(price):
    if price > 5000: interval = 50
    elif price > 1000: interval = 20
    elif price > 500: interval = 10
    else: interval = 5
    return round(price / interval) * interval

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("====================================================================================================")
print("EXACT DAY-BY-DAY PEAK CONCURRENT CAPITAL / CUMULATIVE FUND TIED UP AT A SINGLE TIME")
print("====================================================================================================")

concurrent_results = []

for qtr in quarters_order:
    q_tr = df_src_trades[df_src_trades['Quarter'] == qtr].copy()
    q_tr['Entry Date DT'] = pd.to_datetime(q_tr['Entry Date'])
    q_tr['Exit Date DT']  = pd.to_datetime(q_tr['Exit Date'])
    
    # Process Options Outlay & Futures Margin per trade
    opt_outlays = []
    fut_margins = []
    
    for idx, row in q_tr.iterrows():
        s_entry  = float(row['Entry Futures Price (₹)'])
        lot_size = int(row['Lot Size (Qty)'])
        strat    = str(row['Strategy']).strip().upper()
        is_call  = 'LONG' in strat or 'BUY' in strat
        
        if is_call:
            strike = round_nse_strike(s_entry * 0.99)
            i_entry = max(0.0, s_entry - strike)
        else:
            strike = round_nse_strike(s_entry * 1.01)
            i_entry = max(0.0, strike - s_entry)
            
        p_entry = i_entry + (s_entry * 0.02)
        opt_outlay = round(p_entry * lot_size, 2)
        fut_margin = round(s_entry * lot_size * 0.20, 2)
        
        opt_outlays.append(opt_outlay)
        fut_margins.append(fut_margin)
        
    q_tr['Opt_Outlay'] = opt_outlays
    q_tr['Fut_Margin'] = fut_margins
    
    # Generate all calendar days in the quarter
    min_date = q_tr['Entry Date DT'].min()
    max_date = q_tr['Exit Date DT'].max()
    date_range = pd.date_range(min_date, max_date)
    
    peak_opt_capital = 0.0
    peak_opt_date    = None
    peak_opt_open_trades = 0
    
    peak_fut_capital = 0.0
    peak_fut_date    = None
    peak_fut_open_trades = 0
    
    for d in date_range:
        # Filter trades active on date d
        active_mask = (q_tr['Entry Date DT'] <= d) & (d <= q_tr['Exit Date DT'])
        active_trades = q_tr[active_mask]
        
        opt_tot = active_trades['Opt_Outlay'].sum()
        fut_tot = active_trades['Fut_Margin'].sum()
        num_open = len(active_trades)
        
        if opt_tot > peak_opt_capital:
            peak_opt_capital = opt_tot
            peak_opt_date = d.strftime('%Y-%m-%d')
            peak_opt_open_trades = num_open
            
        if fut_tot > peak_fut_capital:
            peak_fut_capital = fut_tot
            peak_fut_date = d.strftime('%Y-%m-%d')
            peak_fut_open_trades = num_open
            
    concurrent_results.append({
        'Quarter': qtr,
        'Peak Opt Concurrent Capital (₹)': peak_opt_capital,
        'Peak Opt Date': peak_opt_date,
        'Peak Opt Open Trades': peak_opt_open_trades,
        'Peak Fut Concurrent Margin (₹)': peak_fut_capital,
        'Peak Fut Date': peak_fut_date,
        'Peak Fut Open Trades': peak_fut_open_trades
    })

df_res = pd.DataFrame(concurrent_results)
print(df_res.to_string())

print("\n--- SUMMARY OF PEAK CONCURRENT CAPITAL ---")
print(f"Average Quarterly Peak Concurrent Option Premium Tied Up : ₹{df_res['Peak Opt Concurrent Capital (₹)'].mean():,.2f}")
print(f"Average Quarterly Peak Concurrent Futures Margin Deployed : ₹{df_res['Peak Fut Concurrent Margin (₹)'].mean():,.2f}")
print("====================================================================================================")
