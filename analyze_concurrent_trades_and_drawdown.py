import pandas as pd, numpy as np, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
V10_PATH = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v10.xlsx'

xl = pd.ExcelFile(V10_PATH, engine='openpyxl')
df_trades = xl.parse('All_12Q_Futures_Trades')

ret_col = next(c for c in df_trades.columns if 'Return We Get' in c or 'Booked Return' in c or 'Realised Return' in c)
pnl_col = next(c for c in df_trades.columns if 'P&L' in c or 'Profit' in c)
en_dt_col = next(c for c in df_trades.columns if 'Entry Date' in c)
ex_dt_col = next(c for c in df_trades.columns if 'Exit Date' in c)
qtr_col   = next(c for c in df_trades.columns if 'Quarter' in c)

df_trades['Entry Date'] = pd.to_datetime(df_trades[en_dt_col])
df_trades['Exit Date']  = pd.to_datetime(df_trades[ex_dt_col])

quarters = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

print("==========================================================================================")
print("EMPIRICAL CONCURRENT OPEN TRADES & MAX DRAWDOWN ANALYSIS (ALL 12 QUARTERS)")
print("==========================================================================================")

qtr_analysis = []

for q in quarters:
    q_df = df_trades[df_trades[qtr_col] == q].copy().sort_values(by='Entry Date')
    
    min_date = q_df['Entry Date'].min()
    max_date = q_df['Exit Date'].max()
    all_dates = pd.date_range(min_date, max_date, freq='B') # Business days
    
    daily_concurrent = []
    daily_pnl = []
    
    for dt in all_dates:
        # Open positions on date dt: Entry Date <= dt <= Exit Date
        open_pos = q_df[(q_df['Entry Date'] <= dt) & (q_df['Exit Date'] >= dt)]
        daily_concurrent.append({
            'Date': dt,
            'Concurrent Trades': len(open_pos),
            'Symbols Open': list(open_pos['Symbol'].unique())
        })
    
    df_daily_conc = pd.DataFrame(daily_concurrent)
    max_concurrent = df_daily_conc['Concurrent Trades'].max()
    avg_concurrent = df_daily_conc['Concurrent Trades'].mean()
    peak_dates = df_daily_conc[df_daily_conc['Concurrent Trades'] == max_concurrent]['Date'].dt.strftime('%Y-%m-%d').tolist()
    
    # Cumulative PnL & Max Drawdown Calculation for Quarter
    q_df_sorted = q_df.sort_values(by='Exit Date')
    cum_pnl = q_df_sorted[pnl_col].cumsum()
    running_max = np.maximum.accumulate(cum_pnl)
    drawdown = cum_pnl - running_max
    max_dd_val = abs(drawdown.min()) if len(drawdown) > 0 else 0.0
    
    # Capital base
    start_cap = sum(r['Entry Futures Price (₹)'] * r['Lot Size (Qty)'] * 0.20 for _, r in q_df.iterrows()) / len(q_df) * len(q_df['Assigned Slot'].unique())
    max_dd_pct = (max_dd_val / start_cap) if start_cap > 0 else 0.0
    tot_pnl = q_df[pnl_col].sum()
    
    qtr_analysis.append({
        'Quarter': q,
        'Total Trades': len(q_df),
        'Slots Deployed': len(q_df['Assigned Slot'].unique()),
        'Max Concurrent Scripts Open': max_concurrent,
        'Avg Concurrent Scripts Open': round(avg_concurrent, 1),
        'Peak Open Dates': peak_dates[:3],
        'Total Net PnL (₹)': round(tot_pnl, 2),
        'Max Drawdown (₹)': round(max_dd_val, 2),
        'Max Drawdown (%)': round(max_dd_pct, 4)
    })

df_res = pd.DataFrame(qtr_analysis)
print(df_res.to_string())
print("==========================================================================================")
