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
print("EXACT TRADE-BY-TRADE AUDIT OF DRAWDOWNS FROM PEAKS (EQUITY & FUTURES)")
print("==========================================================================================================")

fut_peak_rows = []
eq_peak_rows  = []

for qtr in quarters_order:
    # --- FUTURES ---
    q_fut = df_fut_trades[df_fut_trades['Quarter'] == qtr].copy()
    t_no_fut = next(c for c in q_fut.columns if 'Trade' in c)
    pnl_fut  = next(c for c in q_fut.columns if 'P&L' in c or 'Profit' in c)
    
    q_fut_s = q_fut.sort_values(by=['Entry Date', t_no_fut]).reset_index(drop=True)
    q_fut_s['CumPnL'] = q_fut_s[pnl_fut].cumsum()
    q_fut_s['Peak']   = np.maximum.accumulate(q_fut_s['CumPnL'])
    q_fut_s['DD']     = q_fut_s['CumPnL'] - q_fut_s['Peak']
    
    # Total Margin Deployed across slots
    avg_slot_mrg = sum(r['Entry Futures Price (₹)'] * r['Lot Size (Qty)'] * 0.20 for _, r in q_fut_s.iterrows()) / len(q_fut_s)
    tot_mrg_fut = avg_slot_mrg * len(q_fut_s['Assigned Slot'].unique())
    
    trough_idx_fut = q_fut_s['DD'].idxmin()
    max_dd_fut_val = q_fut_s.loc[trough_idx_fut, 'DD'] # Negative value
    
    # Peak prior to trough
    peak_before_trough_fut = q_fut_s.loc[:trough_idx_fut, 'CumPnL'].max()
    trough_val_fut = q_fut_s.loc[trough_idx_fut, 'CumPnL']
    
    # DD % Option A: Relative to Total Margin Deployed
    dd_pct_mrg_fut = (max_dd_fut_val / tot_mrg_fut) if tot_mrg_fut > 0 else 0.0
    
    # DD % Option B: Relative to Peak Portfolio Value (Peak Cum PnL + Total Margin)
    peak_port_val_fut = tot_mrg_fut + max(0, peak_before_trough_fut)
    dd_pct_peak_fut = (max_dd_fut_val / peak_port_val_fut) if peak_port_val_fut > 0 else 0.0

    fut_peak_rows.append({
        'Quarter': qtr,
        'Tot Margin (₹)': round(tot_mrg_fut, 2),
        'Peak Cum P&L (₹)': round(peak_before_trough_fut, 2),
        'Trough Cum P&L (₹)': round(trough_val_fut, 2),
        'Max DD (₹)': round(max_dd_fut_val, 2),
        'DD % vs Margin': f"{dd_pct_mrg_fut:.2%}",
        'DD % vs Peak Portfolio': f"{dd_pct_peak_fut:.2%}",
        'Trough Trade #': q_fut_s.loc[trough_idx_fut, t_no_fut],
        'Trough Symbol': q_fut_s.loc[trough_idx_fut, 'Symbol']
    })

    # --- EQUITY ---
    q_eq = df_eq_trades[df_eq_trades['Quarter'] == qtr].copy()
    t_no_eq = next(c for c in q_eq.columns if 'Trade' in c)
    pnl_eq  = next(c for c in q_eq.columns if 'P&L' in c or 'Profit' in c)
    
    q_eq_s = q_eq.sort_values(by=['Entry Date', t_no_eq]).reset_index(drop=True)
    q_eq_s['CumPnL'] = q_eq_s[pnl_eq].cumsum()
    q_eq_s['Peak']   = np.maximum.accumulate(q_eq_s['CumPnL'])
    q_eq_s['DD']     = q_eq_s['CumPnL'] - q_eq_s['Peak']
    
    avg_eq_price = q_eq_s['Entry Price (₹)'].mean()
    tot_cap_eq = avg_eq_price * len(q_eq_s['Assigned Slot'].unique())
    
    trough_idx_eq = q_eq_s['DD'].idxmin()
    max_dd_eq_val = q_eq_s.loc[trough_idx_eq, 'DD'] # Negative value
    
    peak_before_trough_eq = q_eq_s.loc[:trough_idx_eq, 'CumPnL'].max()
    trough_val_eq = q_eq_s.loc[trough_idx_eq, 'CumPnL']
    
    dd_pct_cap_eq = (max_dd_eq_val / tot_cap_eq) if tot_cap_eq > 0 else 0.0
    peak_port_val_eq = tot_cap_eq + max(0, peak_before_trough_eq)
    dd_pct_peak_eq = (max_dd_eq_val / peak_port_val_eq) if peak_port_val_eq > 0 else 0.0

    eq_peak_rows.append({
        'Quarter': qtr,
        'Tot Capital (₹)': round(tot_cap_eq, 2),
        'Peak Cum P&L (₹)': round(peak_before_trough_eq, 2),
        'Trough Cum P&L (₹)': round(trough_val_eq, 2),
        'Max DD (₹)': round(max_dd_eq_val, 2),
        'DD % vs Capital': f"{dd_pct_cap_eq:.2%}",
        'DD % vs Peak Portfolio': f"{dd_pct_peak_eq:.2%}",
        'Trough Trade #': q_eq_s.loc[trough_idx_eq, t_no_eq],
        'Trough Symbol': q_eq_s.loc[trough_idx_eq, 'Symbol']
    })

print("\n--- FUTURES DRAWDOWN AUDIT FROM PEAKS ---")
df_fut_peaks = pd.DataFrame(fut_peak_rows)
print(df_fut_peaks.to_string())

print("\n--- EQUITY DRAWDOWN AUDIT FROM PEAKS ---")
df_eq_peaks = pd.DataFrame(eq_peak_rows)
print(df_eq_peaks.to_string())

print("==========================================================================================")
