import pandas as pd, numpy as np, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
SOURCE_PIT_FUT = BASE_DIR / 'Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'

xl_pit = pd.ExcelFile(SOURCE_PIT_FUT, engine='openpyxl')

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

rows = []

for qtr in quarters_order:
    q_tr = xl_pit.parse(qtr, header=6)
    q_tr['Entry Date DT'] = pd.to_datetime(q_tr['Entry Date'])
    q_tr['Exit Date DT']  = pd.to_datetime(q_tr['Exit Date'])
    
    outlays = []
    pnls = []
    
    for idx, row in q_tr.iterrows():
        s_entry  = float(row['Entry Futures Price (₹)'])
        s_exit   = float(row['Exit Futures Price (₹)'])
        lot_size = int(row['Lot Size (Qty)'])
        strat    = str(row['Strategy']).strip().upper()
        
        is_call  = 'LONG' in strat or 'BUY' in strat
        if is_call:
            strike = round_nse_strike(s_entry * 0.99)
            i_entry = max(0.0, s_entry - strike)
            i_exit  = max(0.0, s_exit - strike)
        else:
            strike = round_nse_strike(s_entry * 1.01)
            i_entry = max(0.0, strike - s_entry)
            i_exit  = max(0.0, strike - s_exit)
            
        p_entry = i_entry + (s_entry * 0.02)
        p_exit  = i_exit + (s_exit * 0.005)
        
        outlay = p_entry * lot_size
        pnl = (p_exit - p_entry) * lot_size
        
        outlays.append(outlay)
        pnls.append(pnl)
        
    q_tr['Opt_Outlay'] = outlays
    q_tr['Booked_PnL'] = pnls
    
    tot_pnl = sum(pnls)
    non_reused_capital = sum(outlays)
    
    # Peak Concurrent Capital (Reused Fund Model)
    min_date = q_tr['Entry Date DT'].min()
    max_date = q_tr['Exit Date DT'].max()
    date_range = pd.date_range(min_date, max_date)
    
    peak_concurrent_capital = 0.0
    for d in date_range:
        active_trades = q_tr[(q_tr['Entry Date DT'] <= d) & (d <= q_tr['Exit Date DT'])]
        opt_tot = active_trades['Opt_Outlay'].sum()
        if opt_tot > peak_concurrent_capital:
            peak_concurrent_capital = opt_tot
            
    # Drawdown math
    q_tr['Cum PnL'] = q_tr['Booked_PnL'].cumsum()
    q_tr['Running Max'] = np.maximum.accumulate(q_tr['Cum PnL'])
    q_tr['Drawdown'] = q_tr['Cum PnL'] - q_tr['Running Max']
    
    trough_idx = q_tr['Drawdown'].idxmin()
    max_dd_val = abs(q_tr.loc[trough_idx, 'Drawdown'])
    peak_cum_pnl = q_tr.loc[:trough_idx, 'Cum PnL'].max()
    
    # Reused Fund Drawdowns
    ret_reused = tot_pnl / peak_concurrent_capital
    dd_reused_init = -abs(max_dd_val / peak_concurrent_capital)
    dd_reused_peak = -abs(max_dd_val / (peak_concurrent_capital + max(0.0, peak_cum_pnl)))
    
    # Non-Reused Fund Drawdowns
    ret_no_reuse = tot_pnl / non_reused_capital
    dd_no_reuse_init = -abs(max_dd_val / non_reused_capital)
    dd_no_reuse_peak = -abs(max_dd_val / (non_reused_capital + max(0.0, peak_cum_pnl)))
    
    rows.append({
        'Quarter': qtr,
        'Net Option P&L (₹)': round(tot_pnl, 2),
        'Max DD (₹)': round(max_dd_val, 2),
        'Reused Peak Capital (₹)': round(peak_concurrent_capital, 2),
        'Reused Return (%)': f"{ret_reused:.2%}",
        'Reused DD % Init': f"{dd_reused_init:.2%}",
        'Reused DD % Peak': f"{dd_reused_peak:.2%}",
        'No-Reuse Total Capital (₹)': round(non_reused_capital, 2),
        'No-Reuse Return (%)': f"{ret_no_reuse:.2%}",
        'No-Reuse DD % Init': f"{dd_no_reuse_init:.2%}",
        'No-Reuse DD % Peak': f"{dd_no_reuse_peak:.2%}"
    })

df_comp = pd.DataFrame(rows)

print("========================================================================================================================")
print("SIDE-BY-SIDE COMPARISON: REUSED FUND (CONCURRENT PEAK) vs NON-REUSED FUND (SUM ALL 50 TRADES)")
print("========================================================================================================================")
print(df_comp.to_string())

print("\n--- SUMMARY OF AVERAGES ---")
print(f"REUSED FUND MODEL     — Avg Capital: ₹{df_comp['Reused Peak Capital (₹)'].mean():,.2f} | Avg Return: {df_comp['Reused Return (%)'].str.rstrip('%').astype(float).mean():.2f}% | Avg DD Init: {df_comp['Reused DD % Init'].str.rstrip('%').astype(float).mean():.2f}% | Avg DD Peak: {df_comp['Reused DD % Peak'].str.rstrip('%').astype(float).mean():.2f}%")
print(f"NON-REUSED FUND MODEL — Avg Capital: ₹{df_comp['No-Reuse Total Capital (₹)'].mean():,.2f} | Avg Return: {df_comp['No-Reuse Return (%)'].str.rstrip('%').astype(float).mean():.2f}% | Avg DD Init: {df_comp['No-Reuse DD % Init'].str.rstrip('%').astype(float).mean():.2f}% | Avg DD Peak: {df_comp['No-Reuse DD % Peak'].str.rstrip('%').astype(float).mean():.2f}%")
print("========================================================================================================================")
