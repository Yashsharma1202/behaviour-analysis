import pandas as pd, numpy as np, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
FUT_V13 = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
FUT_PIT = BASE_DIR / 'Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'

xl_v13 = pd.ExcelFile(FUT_V13, engine='openpyxl')
df_v13 = xl_v13.parse('All_12Q_Futures_Trades')

xl_pit = pd.ExcelFile(FUT_PIT, engine='openpyxl')

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

comp_rows = []

for qtr in quarters_order:
    # Fixed Lot V13
    q13 = df_v13[df_v13['Quarter'] == qtr].copy()
    pnl13 = q13['Booked Futures P&L (₹)'].sum()
    mrg13 = (q13['Entry Futures Price (₹)'] * q13['Lot Size (Qty)'] * 0.20).mean() * len(q13['Assigned Slot'].unique())
    
    t_no_13 = next(c for c in q13.columns if 'Trade' in c)
    q13_s = q13.sort_values(by=['Entry Date', t_no_13]).reset_index(drop=True)
    q13_s['Cum'] = q13_s['Booked Futures P&L (₹)'].cumsum()
    q13_s['Peak'] = np.maximum.accumulate(q13_s['Cum'])
    q13_s['DD'] = q13_s['Cum'] - q13_s['Peak']
    idx13 = q13_s['DD'].idxmin()
    dd13_val = abs(q13_s.loc[idx13, 'DD'])
    peak13 = q13_s.loc[:idx13, 'Cum'].max()
    dd13_pct = -abs(dd13_val / (mrg13 + max(0, peak13)))
    
    # Point-in-time PIT (header at row 6)
    qpit = xl_pit.parse(qtr, header=6)
    t_no_pit = next(c for c in qpit.columns if 'Trade' in str(c))
    pnl_col_pit = next(c for c in qpit.columns if 'Booked' in str(c) or 'P&L' in str(c))
    pnl_pit = qpit[pnl_col_pit].sum()
    
    mrg_col_pit = next(c for c in qpit.columns if 'Margin' in str(c))
    mrg_pit = qpit[mrg_col_pit].mean() * len(qpit['Assigned Slot'].dropna().unique())
    
    qpit_s = qpit.sort_values(by=['Entry Date', t_no_pit]).reset_index(drop=True)
    qpit_s['Cum'] = qpit_s[pnl_col_pit].cumsum()
    qpit_s['Peak'] = np.maximum.accumulate(qpit_s['Cum'])
    qpit_s['DD'] = qpit_s['Cum'] - qpit_s['Peak']
    idx_pit = qpit_s['DD'].idxmin()
    dd_pit_val = abs(qpit_s.loc[idx_pit, 'DD'])
    peak_pit = qpit_s.loc[:idx_pit, 'Cum'].max()
    dd_pit_pct = -abs(dd_pit_val / (mrg_pit + max(0, peak_pit)))

    comp_rows.append({
        'Quarter': qtr,
        'Fixed Lot P&L (₹)': round(pnl13, 2),
        'Fixed Lot Max DD (₹)': round(dd13_val, 2),
        'Fixed Lot True DD %': f"{dd13_pct:.2%}",
        'PIT Lot P&L (₹)': round(pnl_pit, 2),
        'PIT Lot Max DD (₹)': round(dd_pit_val, 2),
        'PIT Lot True DD %': f"{dd_pit_pct:.2%}"
    })

df_comp = pd.DataFrame(comp_rows)

print("==========================================================================================")
print("SIDE-BY-SIDE COMPARISON: FIXED LOT SIZE vs POINT-IN-TIME LOT SIZE")
print("==========================================================================================")
print(df_comp.to_string())

print("\n--- TOTALS & AVERAGES COMPARISON ---")
print(f"FIXED LOT MODEL — Total 12Q Futures P&L: ₹{df_comp['Fixed Lot P&L (₹)'].sum():,.2f}")
print(f"PIT LOT MODEL   — Total 12Q Futures P&L: ₹{df_comp['PIT Lot P&L (₹)'].sum():,.2f}")
print("==========================================================================================")
