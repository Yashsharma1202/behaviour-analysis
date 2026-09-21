import pandas as pd, numpy as np, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
FUT_V13 = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'

xl = pd.ExcelFile(FUT_V13, engine='openpyxl')
df_trades = xl.parse('All_12Q_Futures_Trades')

print("==========================================================================================")
print("EXTRACTING EXACT LOT SIZES & TRADE RESULTS ACROSS ALL 12 QUARTERS")
print("==========================================================================================")

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

summary_rows = []

for qtr in quarters_order:
    q_tr = df_trades[df_trades['Quarter'] == qtr].copy()
    
    # Ensure chronological sorting
    t_no_col = next(c for c in q_tr.columns if 'Trade' in c)
    q_tr['Entry Date DT'] = pd.to_datetime(q_tr['Entry Date'])
    q_tr = q_tr.sort_values(by=['Entry Date DT', t_no_col]).reset_index(drop=True)
    
    t_count = len(q_tr)
    pnl_col = next(c for c in q_tr.columns if 'P&L' in c or 'Profit' in c)
    tot_pnl = q_tr[pnl_col].sum()
    
    # Lot size metrics
    avg_lot = q_tr['Lot Size (Qty)'].mean()
    min_lot = q_tr['Lot Size (Qty)'].min()
    max_lot = q_tr['Lot Size (Qty)'].max()
    
    # Margin calculation
    avg_contract_val = (q_tr['Entry Futures Price (₹)'] * q_tr['Lot Size (Qty)']).mean()
    avg_margin_per_lot = avg_contract_val * 0.20
    slots = len(q_tr['Assigned Slot'].unique())
    margin_deployed = avg_margin_per_lot * slots
    
    # Chronological Drawdown
    q_tr['CumPnL'] = q_tr[pnl_col].cumsum()
    q_tr['Peak'] = np.maximum.accumulate(q_tr['CumPnL'])
    q_tr['DD'] = q_tr['CumPnL'] - q_tr['Peak']
    
    max_dd_val = abs(q_tr['DD'].min())
    max_dd_pct = -abs(max_dd_val / margin_deployed) if margin_deployed > 0 else 0.0
    win_rate = len(q_tr[q_tr[pnl_col] > 0]) / t_count
    
    summary_rows.append({
        'Quarter': qtr,
        'Trades': t_count,
        'Avg Lot Size': int(round(avg_lot)),
        'Min/Max Lot': f"{int(min_lot)} / {int(max_lot)}",
        'Avg Contract Val (₹)': round(avg_contract_val, 2),
        '20% Margin Deployed (₹)': round(margin_deployed, 2),
        'Net Futures P&L (₹)': round(tot_pnl, 2),
        'Return on Margin (%)': round(tot_pnl / margin_deployed, 4) if margin_deployed > 0 else 0.0,
        'Max Drawdown (₹)': round(max_dd_val, 2),
        'Max Drawdown (%)': round(max_dd_pct, 4),
        'Win Rate (%)': round(win_rate, 4)
    })

df_res = pd.DataFrame(summary_rows)
print(df_res.to_string())

print("\n--- 12-QUARTER TOTAL & AVERAGE SUMMARY ---")
print(f"Total 12-Quarter Futures P&L   : ₹{df_res['Net Futures P&L (₹)'].sum():,.2f}")
print(f"Average Quarterly Futures P&L  : ₹{df_res['Net Futures P&L (₹)'].mean():,.2f}")
print(f"Average Margin Deployed        : ₹{df_res['20% Margin Deployed (₹)'].mean():,.2f}")
print(f"Average Booked Return on Margin: {df_res['Return on Margin (%)'].mean():.2%}")
print(f"Average Realised Max Drawdown  : -₹{df_res['Max Drawdown (₹)'].mean():,.2f}")
print(f"Peak Realised Max Drawdown     : -₹{df_res['Max Drawdown (₹)'].max():,.2f}")
print(f"Average Realised Max DD %      : {df_res['Max Drawdown (%)'].mean():.2%}")
print(f"Peak Realised Max DD %         : {df_res['Max Drawdown (%)'].min():.2%}")
print(f"Global Overall Win Rate        : {df_res['Win Rate (%)'].mean():.2%}")
print("==========================================================================================")
