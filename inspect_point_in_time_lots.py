import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path_fut = r'D:/behaviour analysis/Futures_Master_Only_v2.xlsx'
xl_fut = pd.ExcelFile(path_fut, engine='openpyxl')

print("==========================================================================================")
print("INSPECTING POINT-IN-TIME LOT SIZES ACROSS 12 QUARTERS")
print("==========================================================================================")
print("Futures Master Sheets:", xl_fut.sheet_names[:6])

quarters = [s for s in xl_fut.sheet_names if s.startswith('Q')]

symbol_sample = ['RELIANCE', 'TCS', 'INFY', 'HDFCBANK', 'SBIN', 'TATAMOTORS', 'TATASTEEL', 'BHARTIARTL']

path_v13 = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
xl_v13 = pd.ExcelFile(path_v13, engine='openpyxl')
df_v13_trades = xl_v13.parse('All_12Q_Futures_Trades')

symbol_lot_qtr = df_v13_trades.groupby(['Quarter', 'Symbol'])['Lot Size (Qty)'].first().reset_index()
pivot_v13 = symbol_lot_qtr[symbol_lot_qtr['Symbol'].isin(symbol_sample)].pivot(index='Symbol', columns='Quarter', values='Lot Size (Qty)')
print("\nPoint-in-Time Lot Sizes in Master Trades across 12 Quarters:")
print(pivot_v13.to_string())

print("==========================================================================================")
