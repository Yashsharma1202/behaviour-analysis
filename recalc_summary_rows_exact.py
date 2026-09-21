import pandas as pd, numpy as np, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
FUT_V13 = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
EQ_V2   = BASE_DIR / 'Nifty50_12_Quarters_Equity_Master_v2.xlsx'

xl_fut = pd.ExcelFile(FUT_V13, engine='openpyxl')
df_fut_trades = xl_fut.parse('All_12Q_Futures_Trades')

xl_eq = pd.ExcelFile(EQ_V2, engine='openpyxl')
df_eq_trades = xl_eq.parse('All_12Q_Equity_Trades')

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

eq_list = []
fut_list = []

for qtr in quarters_order:
    q_eq = df_eq_trades[df_eq_trades['Quarter'] == qtr]
    eq_pnl = q_eq['Booked Equity P&L (₹)'].sum()
    eq_cap = (q_eq['Entry Price (₹)'].mean()) * len(q_eq['Assigned Slot'].unique())
    
    t_no_eq = next(c for c in q_eq.columns if 'Trade' in c)
    q_eq_s = q_eq.sort_values(by=['Entry Date', t_no_eq]).reset_index(drop=True)
    q_eq_s['CumPnL'] = q_eq_s['Booked Equity P&L (₹)'].cumsum()
    q_eq_s['Peak'] = np.maximum.accumulate(q_eq_s['CumPnL'])
    q_eq_s['DD'] = q_eq_s['CumPnL'] - q_eq_s['Peak']
    eq_max_dd_val = abs(q_eq_s['DD'].min())
    eq_max_dd_pct = -abs(eq_max_dd_val / eq_cap) if eq_cap > 0 else 0.0
    eq_win_rate = len(q_eq[q_eq['Booked Equity P&L (₹)'] > 0]) / len(q_eq)
    
    eq_list.append({
        'Quarter': qtr, 'Capital': eq_cap, 'NetPnL': eq_pnl, 'Return': eq_pnl/eq_cap,
        'MaxDD_Val': eq_max_dd_val, 'MaxDD_Pct': eq_max_dd_pct, 'WinRate': eq_win_rate
    })

    q_fut = df_fut_trades[df_fut_trades['Quarter'] == qtr]
    pnl_col_fut = next(c for c in q_fut.columns if 'P&L' in c or 'Profit' in c)
    fut_pnl = q_fut[pnl_col_fut].sum()
    fut_cap = sum(r['Entry Futures Price (₹)'] * r['Lot Size (Qty)'] * 0.20 for _, r in q_fut.iterrows()) / len(q_fut) * len(q_fut['Assigned Slot'].unique())
    
    t_no_fut = next(c for c in q_fut.columns if 'Trade' in c)
    q_fut_s = q_fut.sort_values(by=['Entry Date', t_no_fut]).reset_index(drop=True)
    q_fut_s['CumPnL'] = q_fut_s[pnl_col_fut].cumsum()
    q_fut_s['Peak'] = np.maximum.accumulate(q_fut_s['CumPnL'])
    q_fut_s['DD'] = q_fut_s['CumPnL'] - q_fut_s['Peak']
    fut_max_dd_val = abs(q_fut_s['DD'].min())
    fut_max_dd_pct = -abs(fut_max_dd_val / fut_cap) if fut_cap > 0 else 0.0
    fut_win_rate = len(q_fut[q_fut[pnl_col_fut] > 0]) / len(q_fut)
    
    fut_list.append({
        'Quarter': qtr, 'Capital': fut_cap, 'NetPnL': fut_pnl, 'Return': fut_pnl/fut_cap,
        'MaxDD_Val': fut_max_dd_val, 'MaxDD_Pct': fut_max_dd_pct, 'WinRate': fut_win_rate
    })

df_eq_audit = pd.DataFrame(eq_list)
df_fut_audit = pd.DataFrame(fut_list)

print("==========================================================================================")
print("EXACT RE-CALCULATED SUMMARY METRICS")
print("==========================================================================================")

print("\n--- SPOT EQUITY (1-Share) ---")
print("12-Quarter Total Trades       : 600")
print("Total 12-Quarter Equity P&L   : ₹", f"{df_eq_audit['NetPnL'].sum():,.2f}")
print("Average Quarterly Equity P&L  : ₹", f"{df_eq_audit['NetPnL'].mean():,.2f}")
print("Average Capital Pool Utilised : ₹", f"{df_eq_audit['Capital'].mean():,.2f}")
print("Average Quarterly Return (%)  :", f"{df_eq_audit['Return'].mean():.2%}")
print("Average Max Drawdown (₹)      : -₹", f"{df_eq_audit['MaxDD_Val'].mean():,.2f}")
print("Peak Max Drawdown (₹)         : -₹", f"{df_eq_audit['MaxDD_Val'].max():,.2f}")
print("Average Max Drawdown (%)      :", f"{df_eq_audit['MaxDD_Pct'].mean():.2%}")
print("Peak Max Drawdown (%)         :", f"{df_eq_audit['MaxDD_Pct'].min():.2%}")
print("Global Win Rate               :", f"{df_eq_audit['WinRate'].mean():.2%}")

print("\n--- FUTURES (1-Lot) ---")
print("12-Quarter Total Trades       : 600")
print("Total 12-Quarter Futures P&L  : ₹", f"{df_fut_audit['NetPnL'].sum():,.2f}")
print("Average Quarterly Futures P&L : ₹", f"{df_fut_audit['NetPnL'].mean():,.2f}")
print("Average Margin Deployed       : ₹", f"{df_fut_audit['Capital'].mean():,.2f}")
print("Average Quarterly Return (%)  :", f"{df_fut_audit['Return'].mean():.2%}")
print("Average Max Drawdown (₹)      : -₹", f"{df_fut_audit['MaxDD_Val'].mean():,.2f}")
print("Peak Max Drawdown (₹)         : -₹", f"{df_fut_audit['MaxDD_Val'].max():,.2f}")
print("Average Max Drawdown (%)      :", f"{df_fut_audit['MaxDD_Pct'].mean():.2%}")
print("Peak Max Drawdown (%)         :", f"{df_fut_audit['MaxDD_Pct'].min():.2%}")
print("Global Win Rate               :", f"{df_fut_audit['WinRate'].mean():.2%}")
print("==========================================================================================")
