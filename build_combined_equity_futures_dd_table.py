import pandas as pd, numpy as np, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
V13_PATH = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
EQ_PATH  = BASE_DIR / 'Nifty50_12_Quarters_Equity_Master_v1.xlsx'

xl_v13 = pd.ExcelFile(V13_PATH, engine='openpyxl')
df_fut_trades = xl_v13.parse('All_12Q_Futures_Trades')
df_fut_summary = xl_v13.parse('Exec_12Q_Combined_Summary')

xl_eq = pd.ExcelFile(EQ_PATH, engine='openpyxl')
eq_sheet_name = next(s for s in xl_eq.sheet_names if 'All' in s or 'Trade' in s)
df_eq_trades = xl_eq.parse(eq_sheet_name)
df_eq_summary = xl_eq.parse('Exec_12Q_Combined_Summary')

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("==========================================================================================================================================")
print("AUDITING EQUITY & FUTURES DRAWDOWN WITH NEGATIVE PERCENTAGES ACROSS ALL 12 QUARTERS")
print("==========================================================================================================================================")

combined_data = []

for qtr in quarters_order:
    # --- Futures Calculations ---
    q_fut = df_fut_trades[df_fut_trades['Quarter'] == qtr].copy()
    q_fut['Entry Date DT'] = pd.to_datetime(q_fut['Entry Date'])
    trade_no_col = next(c for c in q_fut.columns if 'Trade' in c)
    pnl_fut_col = next(c for c in q_fut.columns if 'P&L' in c or 'Profit' in c)
    q_fut = q_fut.sort_values(by=['Entry Date DT', trade_no_col]).reset_index(drop=True)
    
    q_fut['Cum PnL'] = q_fut[pnl_fut_col].cumsum()
    q_fut['Peak PnL'] = np.maximum.accumulate(q_fut['Cum PnL'])
    q_fut['DD'] = q_fut['Cum PnL'] - q_fut['Peak PnL']
    fut_real_dd_val = abs(q_fut['DD'].min())
    
    # Get Fut Summary Row
    fut_row = df_fut_summary[df_fut_summary.iloc[:, 0] == qtr].iloc[0]
    fut_fund = float(fut_row.iloc[5])
    fut_net_pnl = float(fut_row.iloc[6])
    fut_ret = float(fut_row.iloc[8])
    fut_mtm_dd_val = float(fut_row.iloc[11]) if pd.notnull(fut_row.iloc[11]) else 0.0
    
    fut_real_dd_pct = -abs(fut_real_dd_val / fut_fund) if fut_fund > 0 else 0.0
    fut_mtm_dd_pct  = -abs(fut_mtm_dd_val / fut_fund) if fut_fund > 0 else 0.0
    
    # --- Equity Calculations ---
    q_eq = df_eq_trades[df_eq_trades['Quarter'] == qtr].copy()
    q_eq['Entry Date DT'] = pd.to_datetime(q_eq['Entry Date'])
    eq_pnl_col = next(c for c in q_eq.columns if 'P&L' in c or 'Profit' in c)
    q_eq = q_eq.sort_values(by=['Entry Date DT']).reset_index(drop=True)
    
    q_eq['Cum PnL'] = q_eq[eq_pnl_col].cumsum()
    q_eq['Peak PnL'] = np.maximum.accumulate(q_eq['Cum PnL'])
    q_eq['DD'] = q_eq['Cum PnL'] - q_eq['Peak PnL']
    eq_real_dd_val = abs(q_eq['DD'].min())
    
    eq_row = df_eq_summary[df_eq_summary.iloc[:, 0] == qtr].iloc[0]
    eq_fund = float(eq_row.iloc[5])
    eq_net_pnl = float(eq_row.iloc[6])
    eq_ret = float(eq_row.iloc[8])
    
    eq_real_dd_pct = -abs(eq_real_dd_val / eq_fund) if eq_fund > 0 else 0.0
    
    combined_data.append({
        'Quarter': qtr,
        'Trades': len(q_fut),
        'Equity Fund (₹)': round(eq_fund, 2),
        'Equity Net P&L (₹)': round(eq_net_pnl, 2),
        'Equity Ret (%)': f"{eq_ret:+.2%}",
        'Equity Max DD (₹)': f"-₹{eq_real_dd_val:,.2f}",
        'Equity Max DD (%)': f"{eq_real_dd_pct:.2%}",
        'Futures Margin (₹)': round(fut_fund, 2),
        'Futures Net P&L (₹)': round(fut_net_pnl, 2),
        'Futures Ret (%)': f"{fut_ret:+.2%}",
        'Futures Realised DD (₹)': f"-₹{fut_real_dd_val:,.2f}",
        'Futures Realised DD (%)': f"{fut_real_dd_pct:.2%}",
        'Futures MTM DD (₹)': f"-₹{fut_mtm_dd_val:,.2f}",
        'Futures MTM DD (%)': f"{fut_mtm_dd_pct:.2%}"
    })

df_result = pd.DataFrame(combined_data)
print(df_result.to_string())
print("==========================================================================================================================================")
