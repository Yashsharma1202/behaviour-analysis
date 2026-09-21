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

print("==========================================================================================================")
print("DOUBLE-CHECK AUDIT: LINE-BY-LINE VERIFICATION OF DRAWDOWN MATH (EQUITY & FUTURES)")
print("==========================================================================================================")

fut_audit_results = []
eq_audit_results  = []

for qtr in quarters_order:
    # --- 1. FUTURES DOUBLE-CHECK ---
    q_fut = df_fut_trades[df_fut_trades['Quarter'] == qtr].copy()
    q_fut['Entry Date DT'] = pd.to_datetime(q_fut['Entry Date'])
    t_no_col_fut = next(c for c in q_fut.columns if 'Trade' in c)
    pnl_col_fut  = next(c for c in q_fut.columns if 'P&L' in c or 'Profit' in c)
    
    # Verify Sorting
    q_fut_sorted = q_fut.sort_values(by=['Entry Date DT', t_no_col_fut]).reset_index(drop=True)
    
    # Recalculate Trade-by-Trade Cumulative & DD
    q_fut_sorted['Recalc_Cum_PnL'] = q_fut_sorted[pnl_col_fut].cumsum()
    q_fut_sorted['Recalc_Peak']    = np.maximum.accumulate(q_fut_sorted['Recalc_Cum_PnL'])
    q_fut_sorted['Recalc_DD']      = q_fut_sorted['Recalc_Cum_PnL'] - q_fut_sorted['Recalc_Peak']
    
    fut_max_dd_val = abs(q_fut_sorted['Recalc_DD'].min())
    fut_trough_trade = q_fut_sorted.loc[q_fut_sorted['Recalc_DD'].idxmin()]
    fut_peak_before_trough = q_fut_sorted.loc[:q_fut_sorted['Recalc_DD'].idxmin(), 'Recalc_Cum_PnL'].max()
    
    # Capital base
    lot_size = q_fut_sorted.iloc[0]['Lot Size (Qty)']
    avg_margin = sum(r['Entry Futures Price (₹)'] * r['Lot Size (Qty)'] * 0.20 for _, r in q_fut_sorted.iterrows()) / len(q_fut_sorted)
    slots = len(q_fut_sorted['Assigned Slot'].unique())
    fut_margin_deployed = round(avg_margin * slots, 2)
    fut_max_dd_pct = -abs(fut_max_dd_val / fut_margin_deployed) if fut_margin_deployed > 0 else 0.0
    
    fut_audit_results.append({
        'Quarter': qtr,
        'Trades': len(q_fut_sorted),
        'Margin Deployed (₹)': fut_margin_deployed,
        'Net Futures P&L (₹)': round(q_fut_sorted[pnl_col_fut].sum(), 2),
        'Peak Cum P&L (₹)': round(fut_peak_before_trough, 2),
        'Trough Cum P&L (₹)': round(fut_trough_trade['Recalc_Cum_PnL'], 2),
        'Max Drawdown (₹)': round(fut_max_dd_val, 2),
        'Max Drawdown (%)': f"{fut_max_dd_pct:.2%}",
        'Worst Drop Stock': fut_trough_trade['Symbol'],
        'Trough Date': fut_trough_trade['Entry Date']
    })

    # --- 2. EQUITY DOUBLE-CHECK ---
    q_eq = df_eq_trades[df_eq_trades['Quarter'] == qtr].copy()
    q_eq['Entry Date DT'] = pd.to_datetime(q_eq['Entry Date'])
    t_no_col_eq = next(c for c in q_eq.columns if 'Trade' in c)
    pnl_col_eq  = next(c for c in q_eq.columns if 'P&L' in c or 'Profit' in c)
    
    q_eq_sorted = q_eq.sort_values(by=['Entry Date DT', t_no_col_eq]).reset_index(drop=True)
    
    q_eq_sorted['Recalc_Cum_PnL'] = q_eq_sorted[pnl_col_eq].cumsum()
    q_eq_sorted['Recalc_Peak']    = np.maximum.accumulate(q_eq_sorted['Recalc_Cum_PnL'])
    q_eq_sorted['Recalc_DD']      = q_eq_sorted['Recalc_Cum_PnL'] - q_eq_sorted['Recalc_Peak']
    
    eq_max_dd_val = abs(q_eq_sorted['Recalc_DD'].min())
    eq_trough_trade = q_eq_sorted.loc[q_eq_sorted['Recalc_DD'].idxmin()]
    eq_peak_before_trough = q_eq_sorted.loc[:q_eq_sorted['Recalc_DD'].idxmin(), 'Recalc_Cum_PnL'].max()
    
    avg_eq_price = sum(r['Entry Price (₹)'] for _, r in q_eq_sorted.iterrows()) / len(q_eq_sorted)
    eq_slots = len(q_eq_sorted['Assigned Slot'].unique())
    eq_fund_utilised = round(avg_eq_price * eq_slots, 2)
    eq_max_dd_pct = -abs(eq_max_dd_val / eq_fund_utilised) if eq_fund_utilised > 0 else 0.0
    
    eq_audit_results.append({
        'Quarter': qtr,
        'Trades': len(q_eq_sorted),
        'Equity Capital (₹)': eq_fund_utilised,
        'Net Equity P&L (₹)': round(q_eq_sorted[pnl_col_eq].sum(), 2),
        'Peak Cum P&L (₹)': round(eq_peak_before_trough, 2),
        'Trough Cum P&L (₹)': round(eq_trough_trade['Recalc_Cum_PnL'], 2),
        'Max Drawdown (₹)': round(eq_max_dd_val, 2),
        'Max Drawdown (%)': f"{eq_max_dd_pct:.2%}",
        'Worst Drop Stock': eq_trough_trade['Symbol'],
        'Trough Date': eq_trough_trade['Entry Date']
    })

print("\n--- FUTURES DRAWDOWN DOUBLE-CHECK AUDIT TABLE ---")
df_fut_res = pd.DataFrame(fut_audit_results)
print(df_fut_res.to_string())

print("\n--- EQUITY DRAWDOWN DOUBLE-CHECK AUDIT TABLE ---")
df_eq_res = pd.DataFrame(eq_audit_results)
print(df_eq_res.to_string())

print("\n==========================================================================================================")
