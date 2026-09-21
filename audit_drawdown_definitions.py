import pandas as pd, numpy as np, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
V11_PATH = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v11.xlsx'
FUTURES_MASTER_EXISTING = BASE_DIR / 'Futures_Master_Only_v2.xlsx'

xl = pd.ExcelFile(V11_PATH, engine='openpyxl')
df_trades = xl.parse('All_12Q_Futures_Trades')
df_daily  = pd.read_excel(FUTURES_MASTER_EXISTING, sheet_name='Futures_Master', engine='openpyxl')

df_daily['Date'] = pd.to_datetime(df_daily['Date']).dt.strftime('%Y-%m-%d')
df_daily['Symbol'] = df_daily['Symbol'].astype(str).str.strip()

price_lookup_map = {}
for _, r in df_daily.iterrows():
    p = r.iloc[2]
    if pd.notnull(p) and float(p) > 0:
        price_lookup_map[(r['Symbol'], str(r['Date']))] = float(p)

quarters = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("======================================================================================================================")
print("DRAWDOWN COMPREHENSIVE AUDIT ACROSS ALL 3 FINANCIAL DEFINITIONS")
print("======================================================================================================================")

trade_no_col = next(c for c in df_trades.columns if 'Trade' in c)
pnl_col = next(c for c in df_trades.columns if 'P&L' in c or 'Profit' in c)

comparison = []

for q in quarters:
    q_trades = df_trades[df_trades['Quarter'] == q].copy()
    q_trades['Entry Date'] = pd.to_datetime(q_trades['Entry Date'])
    q_trades['Exit Date']  = pd.to_datetime(q_trades['Exit Date'])
    
    # Capital base
    start_cap = sum(r['Entry Futures Price (₹)'] * r['Lot Size (Qty)'] * 0.20 for _, r in q_trades.iterrows()) / len(q_trades) * len(q_trades['Assigned Slot'].unique())
    
    # --- Definition 1: Closed-Trade Realised DD (Sorted by Exit Date) ---
    qt_exit = q_trades.sort_values(by='Exit Date').copy()
    qt_exit['Cum PnL'] = qt_exit[pnl_col].cumsum()
    qt_exit['Peak PnL'] = np.maximum.accumulate(qt_exit['Cum PnL'])
    qt_exit['DD'] = qt_exit['Cum PnL'] - qt_exit['Peak PnL']
    dd_closed_exit = abs(qt_exit['DD'].min())
    
    # --- Definition 2: Closed-Trade Realised DD (Sorted by Entry Date) ---
    qt_entry = q_trades.sort_values(by=['Entry Date', 'Exit Date']).copy()
    qt_entry['Cum PnL'] = qt_entry[pnl_col].cumsum()
    qt_entry['Peak PnL'] = np.maximum.accumulate(qt_entry['Cum PnL'])
    qt_entry['DD'] = qt_entry['Cum PnL'] - qt_entry['Peak PnL']
    dd_closed_entry = abs(qt_entry['DD'].min())

    # --- Definition 3: Closed-Trade Realised DD (Sorted by Trade No) ---
    qt_no = q_trades.sort_values(by=trade_no_col).copy()
    qt_no['Cum PnL'] = qt_no[pnl_col].cumsum()
    qt_no['Peak PnL'] = np.maximum.accumulate(qt_no['Cum PnL'])
    qt_no['DD'] = qt_no['Cum PnL'] - qt_no['Peak PnL']
    dd_closed_trade_no = abs(qt_no['DD'].min())

    # --- Definition 4: Worst Single Losing Trade (₹) ---
    worst_single_loss = abs(q_trades[pnl_col].min())
    
    # --- Definition 5: Daily MTM Equity Curve Drawdown ---
    min_date = q_trades['Entry Date'].min()
    max_date = q_trades['Exit Date'].max()
    all_b_dates = pd.date_range(min_date, max_date, freq='B')
    
    daily_equity = []
    
    for dt in all_b_dates:
        dt_str = dt.strftime('%Y-%m-%d')
        
        # Closed trades up to dt
        closed_t = q_trades[q_trades['Exit Date'] < dt]
        closed_pnl = closed_t[pnl_col].sum()
        
        # Open trades on dt (Entry Date <= dt <= Exit Date)
        open_t = q_trades[(q_trades['Entry Date'] <= dt) & (q_trades['Exit Date'] >= dt)]
        open_mtm_pnl = 0.0
        
        for _, ot in open_t.iterrows():
            sym = ot['Symbol']
            direction = 1 if 'LONG' in str(ot['Strategy']).upper() else -1
            lot = ot['Lot Size (Qty)']
            en_p = ot['Entry Futures Price (₹)']
            
            curr_p = price_lookup_map.get((sym, dt_str), en_p)
            # if unadjusted split in raw data, fallback to entry
            if curr_p > 1.8 * en_p or curr_p < 0.5 * en_p:
                curr_p = en_p
                
            open_mtm_pnl += direction * (curr_p - en_p) * lot
            
        tot_eq = start_cap + closed_pnl + open_mtm_pnl
        daily_equity.append(tot_eq)
        
    daily_equity_series = pd.Series(daily_equity)
    daily_peak = np.maximum.accumulate(daily_equity_series)
    daily_dd = daily_equity_series - daily_peak
    max_daily_mtm_dd = abs(daily_dd.min())
    
    comparison.append({
        'Quarter': q,
        'Fund Utilised (₹)': round(start_cap, 2),
        'DD (Exit Order)': round(dd_closed_exit, 2),
        'DD (Entry Order)': round(dd_closed_entry, 2),
        'DD (Trade No Order)': round(dd_closed_trade_no, 2),
        'Daily MTM DD (₹)': round(max_daily_mtm_dd, 2),
        'Daily MTM DD (%)': f"{(max_daily_mtm_dd / start_cap):.2%}",
        'Worst Single Loss (₹)': round(worst_single_loss, 2)
    })

df_comp = pd.DataFrame(comparison)
print(df_comp.to_string())
print("======================================================================================================================")
