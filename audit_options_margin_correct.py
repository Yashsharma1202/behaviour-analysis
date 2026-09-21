import pandas as pd, numpy as np, pathlib, sys

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

print("==========================================================================================")
print("AUDITING INDIVIDUAL STOCK 1% ITM OPTIONS PREMIUM OUTLAYS PER QUARTER")
print("==========================================================================================")

for qtr in quarters_order:
    q_tr = df_src_trades[df_src_trades['Quarter'] == qtr].copy()
    outlays = []
    pnls = []
    
    for idx, row in q_tr.iterrows():
        s_entry = float(row['Entry Futures Price (₹)'])
        s_exit  = float(row['Exit Futures Price (₹)'])
        lot_size = int(row['Lot Size (Qty)'])
        strat   = str(row['Strategy']).strip().upper()
        is_call = 'LONG' in strat or 'BUY' in strat
        
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
        
    avg_outlay_per_trade = np.mean(outlays)
    sum_all_50_outlays   = np.sum(outlays)
    sum_15_slots_outlay  = np.mean(outlays) * len(q_tr['Assigned Slot'].unique())
    tot_pnl              = np.sum(pnls)
    
    print(f"{qtr:12s} | Avg/Trade: ₹{avg_outlay_per_trade:9,.2f} | 15-Slot Pool: ₹{sum_15_slots_outlay:10,.2f} | Sum All 50: ₹{sum_all_50_outlays:10,.2f} | Net PnL: ₹{tot_pnl:11,.2f}")

print("==========================================================================================")
