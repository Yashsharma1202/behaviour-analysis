import openpyxl
import pandas as pd
import numpy as np
import pathlib
import sys
import time

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

# ------------------- Configuration & Paths -------------------
BASE_DIR = pathlib.Path('D:/behaviour analysis')
SOURCE_PIT_FUT = BASE_DIR / 'Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'
OUTPUT_EXACT_COST_MASTER = BASE_DIR / 'Nifty50_12_Quarters_Futures_0.5PCT_BuySell_Turnover_Master.xlsx'
OUTPUT_EXACT_COST_SUMMARY = BASE_DIR / 'Nifty50_12_Quarters_Futures_0.5PCT_BuySell_Turnover_Summary.xlsx'

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("Calculating Exact Buy Turnover + Sell Turnover (0.5% each leg) Cost...", flush=True)
xl_pit = pd.ExcelFile(SOURCE_PIT_FUT, engine='openpyxl')

all_cost_trades = []

for qtr in quarters_order:
    q_tr = xl_pit.parse(qtr, header=6)
    
    t_no_col = next(c for c in q_tr.columns if 'Trade' in str(c))
    q_tr['Entry Date DT'] = pd.to_datetime(q_tr['Entry Date'])
    q_tr = q_tr.sort_values(by=['Entry Date DT', t_no_col]).reset_index(drop=True)
    
    for idx, row in q_tr.iterrows():
        sym      = str(row['Symbol']).strip()
        lot_size = int(row['Lot Size (Qty)'])
        s_entry  = float(row['Entry Futures Price (₹)'])
        s_exit   = float(row['Exit Futures Price (₹)'])
        strat    = str(row['Strategy']).strip().upper()
        
        is_long = 'LONG' in strat or 'BUY' in strat
        
        raw_pnl = (s_exit - s_entry) * lot_size if is_long else (s_entry - s_exit) * lot_size
        
        buy_turnover  = s_entry * lot_size
        sell_turnover = s_exit * lot_size
        total_turnover = buy_turnover + sell_turnover
        
        # 0.5% on Buy Turnover + 0.5% on Sell Turnover = 0.5% of Total Combined Turnover
        trans_cost_buysell = round(total_turnover * 0.005, 2)
        net_pnl_buysell    = round(raw_pnl - trans_cost_buysell, 2)
        margin_outlay      = round(buy_turnover * 0.20, 2)
        
        all_cost_trades.append({
            'Quarter': qtr,
            'Symbol': sym,
            'Strategy': strat,
            'Entry Spot Price (₹)': s_entry,
            'Exit Spot Price (₹)': s_exit,
            'Lot Size': lot_size,
            'Buy Turnover (₹)': round(buy_turnover, 2),
            'Sell Turnover (₹)': round(sell_turnover, 2),
            'Total Combined Turnover (₹)': round(total_turnover, 2),
            'Raw Futures P&L (₹)': round(raw_pnl, 2),
            'Transaction Cost (0.5% Buy + Sell Turnover) (₹)': trans_cost_buysell,
            'Net Futures P&L (After Buy+Sell Cost) (₹)': net_pnl_buysell
        })

df_all_exact = pd.DataFrame(all_cost_trades)

tot_raw_pnl   = df_all_exact['Raw Futures P&L (₹)'].sum()
tot_cost_bs   = df_all_exact['Transaction Cost (0.5% Buy + Sell Turnover) (₹)'].sum()
tot_net_pnl_bs = df_all_exact['Net Futures P&L (After Buy+Sell Cost) (₹)'].sum()
net_wins      = len(df_all_exact[df_all_exact['Net Futures P&L (After Buy+Sell Cost) (₹)'] > 0])

print("==========================================================================================")
print("EXACT BUY TURNOVER + SELL TURNOVER (0.5% EACH LEG) COST RESULTS")
print("==========================================================================================")
print(f"Total Combined 12-Quarter Turnover : ₹{df_all_exact['Total Combined Turnover (₹)'].sum():,.2f}")
print(f"Total Raw Futures Realised Profit   : ₹{tot_raw_pnl:,.2f}")
print(f"Total Transaction Cost (0.5% Buy + 0.5% Sell): ₹{tot_cost_bs:,.2f}")
print(f"Net Futures Profit (After Buy+Sell Cost)    : ₹{tot_net_pnl_bs:,.2f}")
print(f"Net Global Win Rate                         : {net_wins/600:.2%} ({net_wins}/600 Wins)")
print("==========================================================================================")
