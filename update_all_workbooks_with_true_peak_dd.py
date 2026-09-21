import openpyxl, pandas as pd, numpy as np, pathlib, sys, time

sys.stdout.reconfigure(errors='replace')

t0 = time.time()

BASE_DIR = pathlib.Path('D:/behaviour analysis')
FUT_V13 = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
EQ_V2   = BASE_DIR / 'Nifty50_12_Quarters_Equity_Master_v2.xlsx'
FULL_CONS = BASE_DIR / 'Nifty50_12_Quarters_Equity_Consolidated_Summary.xlsx'
SUMMARY_EXCEL = BASE_DIR / 'Nifty50_Consolidated_12_Quarter_Drawdown_Summary.xlsx'

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("==========================================================================================")
print("RE-CALCULATING DRAWDOWN PERCENTAGES FROM TRUE PEAKS FOR ALL WORKBOOKS")
print("==========================================================================================")

# --- 1. FUTURES WORKBOOK RE-CALCULATION ---
xl_fut = pd.ExcelFile(FUT_V13, engine='openpyxl')
df_fut_trades = xl_fut.parse('All_12Q_Futures_Trades')

fut_summary_rows = []

for qtr in quarters_order:
    q_tr = df_fut_trades[df_fut_trades['Quarter'] == qtr].copy()
    t_no = next(c for c in q_tr.columns if 'Trade' in c)
    pnl_col = next(c for c in q_tr.columns if 'P&L' in c or 'Profit' in c)
    
    q_s = q_tr.sort_values(by=['Entry Date', t_no]).reset_index(drop=True)
    q_s['CumPnL'] = q_s[pnl_col].cumsum()
    q_s['Peak']   = np.maximum.accumulate(q_s['CumPnL'])
    q_s['DD']     = q_s['CumPnL'] - q_s['Peak']
    
    t_count = len(q_s)
    tot_pnl = q_s[pnl_col].sum()
    
    avg_slot_mrg = sum(r['Entry Futures Price (₹)'] * r['Lot Size (Qty)'] * 0.20 for _, r in q_s.iterrows()) / t_count
    tot_mrg = round(avg_slot_mrg * len(q_s['Assigned Slot'].unique()), 2)
    
    trough_idx = q_s['DD'].idxmin()
    max_dd_val = abs(q_s.loc[trough_idx, 'DD']) # Positive magnitude
    peak_cum_pnl = q_s.loc[:trough_idx, 'CumPnL'].max()
    
    # Peak Portfolio Value = Capital/Margin Deployed + Peak Cum PnL prior to trough
    peak_portfolio_val = tot_mrg + max(0.0, peak_cum_pnl)
    max_dd_pct_peak = -abs(max_dd_val / peak_portfolio_val) if peak_portfolio_val > 0 else 0.0
    
    win_rate = len(q_s[q_s[pnl_col] > 0]) / t_count
    
    fut_summary_rows.append({
        'Quarter': qtr,
        'Trades': t_count,
        '20% Margin Deployed (₹)': tot_mrg,
        'Net Futures P&L (₹)': round(tot_pnl, 2),
        'Peak Cum P&L (₹)': round(peak_cum_pnl, 2),
        'Trough Cum P&L (₹)': round(q_s.loc[trough_idx, 'CumPnL'], 2),
        'Max Drawdown (₹)': round(max_dd_val, 2),
        'Max DD % (from Peak)': round(max_dd_pct_peak, 4),
        'Win Rate (%)': round(win_rate, 4),
        'Trough Stock': q_s.loc[trough_idx, 'Symbol']
    })

df_fut_peak_sum = pd.DataFrame(fut_summary_rows)

# --- 2. EQUITY WORKBOOK RE-CALCULATION ---
xl_eq = pd.ExcelFile(EQ_V2, engine='openpyxl')
df_eq_trades = xl_eq.parse('All_12Q_Equity_Trades')

eq_summary_rows = []

for qtr in quarters_order:
    q_tr = df_eq_trades[df_eq_trades['Quarter'] == qtr].copy()
    t_no = next(c for c in q_tr.columns if 'Trade' in c)
    pnl_col = next(c for c in q_tr.columns if 'P&L' in c or 'Profit' in c)
    
    q_s = q_tr.sort_values(by=['Entry Date', t_no]).reset_index(drop=True)
    q_s['CumPnL'] = q_s[pnl_col].cumsum()
    q_s['Peak']   = np.maximum.accumulate(q_s['CumPnL'])
    q_s['DD']     = q_s['CumPnL'] - q_s['Peak']
    
    t_count = len(q_s)
    tot_pnl = q_s[pnl_col].sum()
    
    avg_eq_price = q_s['Entry Price (₹)'].mean()
    tot_cap = round(avg_eq_price * len(q_s['Assigned Slot'].unique()), 2)
    
    trough_idx = q_s['DD'].idxmin()
    max_dd_val = abs(q_s.loc[trough_idx, 'DD'])
    peak_cum_pnl = q_s.loc[:trough_idx, 'CumPnL'].max()
    
    peak_portfolio_val = tot_cap + max(0.0, peak_cum_pnl)
    max_dd_pct_peak = -abs(max_dd_val / peak_portfolio_val) if peak_portfolio_val > 0 else 0.0
    
    win_rate = len(q_s[q_s[pnl_col] > 0]) / t_count
    
    eq_summary_rows.append({
        'Quarter': qtr,
        'Trades': t_count,
        'Equity Capital (₹)': tot_cap,
        'Net Equity P&L (₹)': round(tot_pnl, 2),
        'Peak Cum P&L (₹)': round(peak_cum_pnl, 2),
        'Trough Cum P&L (₹)': round(q_s.loc[trough_idx, 'CumPnL'], 2),
        'Max Drawdown (₹)': round(max_dd_val, 2),
        'Max DD % (from Peak)': round(max_dd_pct_peak, 4),
        'Win Rate (%)': round(win_rate, 4),
        'Trough Stock': q_s.loc[trough_idx, 'Symbol']
    })

df_eq_peak_sum = pd.DataFrame(eq_summary_rows)

print("\n--- RE-CALCULATED FUTURES DRAWDOWNS FROM PEAK PORTFOLIO VALUE ---")
print(df_fut_peak_sum.to_string())

print("\n--- RE-CALCULATED EQUITY DRAWDOWNS FROM PEAK PORTFOLIO VALUE ---")
print(df_eq_peak_sum.to_string())

print(f"\nCompleted Audit in {time.time() - t0:.1f}s!")
print("==========================================================================================")
