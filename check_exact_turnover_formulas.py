import openpyxl, pandas as pd, numpy as np, pathlib, sys

sys.stdout.reconfigure(errors='replace')

src = r'D:\behaviour analysis\Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'
xl = pd.ExcelFile(src, engine='openpyxl')

quarters_order = [
    'Q3 2023-24', 'Q4 2023-24', 'Q1 2024-25', 'Q2 2024-25',
    'Q3 2024-25', 'Q4 2024-25', 'Q1 2025-26', 'Q2 2025-26',
    'Q3 2025-26', 'Q4 2025-26', 'Q1 2026-27', 'Q2 2026-27'
]

rows = []
for qtr in quarters_order:
    df_q = xl.parse(qtr, header=6)
    for idx, r in df_q.iterrows():
        s_entry = float(r['Entry Futures Price (₹)'])
        s_exit  = float(r['Exit Futures Price (₹)'])
        lot     = int(r['Lot Size (Qty)'])
        strat   = str(r['Strategy']).strip().upper()
        is_long = 'LONG' in strat or 'BUY' in strat
        
        raw_pnl = (s_exit - s_entry) * lot if is_long else (s_entry - s_exit) * lot
        buy_to  = s_entry * lot
        sell_to = s_exit * lot
        comb_to = buy_to + sell_to
        
        rows.append({
            'Quarter': qtr,
            'Symbol': r['Symbol'],
            'Buy_TO': buy_to,
            'Sell_TO': sell_to,
            'Comb_TO': comb_to,
            'Raw_PnL': raw_pnl,
            'Cost_05_Comb': comb_to * 0.005,
            'Cost_025_Comb': comb_to * 0.0025,
            'Cost_05_Buy': buy_to * 0.005
        })

df_all = pd.DataFrame(rows)

print("==========================================================================================")
print("COMPARING EXACT TRANSACTION COST FORMULAS (600 TRADES)")
print("==========================================================================================")
print(f"Total Combined Turnover (Buy+Sell) : ₹{df_all['Comb_TO'].sum():,.2f}")
print(f"Total Raw Futures P&L              : ₹{df_all['Raw_PnL'].sum():,.2f}")
print("------------------------------------------------------------------------------------------")
print(f"Formula 1: 0.5% on Combined Turnover (0.005 * Comb_TO)  -> Total Cost: ₹{df_all['Cost_05_Comb'].sum():,.2f} | Net P&L: ₹{df_all['Raw_PnL'].sum() - df_all['Cost_05_Comb'].sum():,.2f}")
print(f"Formula 2: 0.5% Total (0.25% Buy + 0.25% Sell = 0.0025*Comb_TO) -> Total Cost: ₹{df_all['Cost_025_Comb'].sum():,.2f} | Net P&L: ₹{df_all['Raw_PnL'].sum() - df_all['Cost_025_Comb'].sum():,.2f}")
print(f"Formula 3: 0.5% on Entry Value (0.005 * Buy_TO)         -> Total Cost: ₹{df_all['Cost_05_Buy'].sum():,.2f} | Net P&L: ₹{df_all['Raw_PnL'].sum() - df_all['Cost_05_Buy'].sum():,.2f}")
print("==========================================================================================")

# Row 1 sample print
r1 = df_all.iloc[0]
print(f"\nRow 1 ({r1['Symbol']}):")
print(f"  Buy Turnover: ₹{r1['Buy_TO']:,.2f} | Sell Turnover: ₹{r1['Sell_TO']:,.2f} | Combined Turnover: ₹{r1['Comb_TO']:,.2f}")
print(f"  Raw P&L: ₹{r1['Raw_PnL']:,.2f}")
print(f"  Formula 1 (0.5% * Comb_TO = 0.005 * {r1['Comb_TO']:,.2f})  : ₹{r1['Cost_05_Comb']:,.2f} -> Net P&L: ₹{r1['Raw_PnL'] - r1['Cost_05_Comb']:,.2f}")
print(f"  Formula 2 (0.25% * Comb_TO = 0.0025 * {r1['Comb_TO']:,.2f}): ₹{r1['Cost_025_Comb']:,.2f} -> Net P&L: ₹{r1['Raw_PnL'] - r1['Cost_025_Comb']:,.2f}")
print(f"  Formula 3 (0.5% * Buy_TO = 0.005 * {r1['Buy_TO']:,.2f})   : ₹{r1['Cost_05_Buy']:,.2f} -> Net P&L: ₹{r1['Raw_PnL'] - r1['Cost_05_Buy']:,.2f}")
print("==========================================================================================")
