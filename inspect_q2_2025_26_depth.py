import pandas as pd, numpy as np, pathlib, sys

sys.stdout.reconfigure(errors='replace')

BASE_DIR = pathlib.Path('D:/behaviour analysis')
V12_PATH = BASE_DIR / 'Nifty50_12_Quarters_Futures_OI_Master_v12.xlsx'
CONSOLIDATED_PATH = BASE_DIR / 'Nifty50_12_Quarters_Consolidated.xlsx'

xl_v12 = pd.ExcelFile(V12_PATH, engine='openpyxl')
df_v12 = xl_v12.parse('Q2 2025-26')

xl_cons = pd.ExcelFile(CONSOLIDATED_PATH, engine='openpyxl')
df_cons_q2 = xl_cons.parse('Q2 2025-26')

print("==========================================")
print("GRANULAR INSPECTION OF Q2 2025-26")
print("==========================================")

print("Master Consolidated Sheet Q2 2025-26 Card Row (Row index 3):")
print(df_cons_q2.iloc[3, :13].to_string())

print("\nTrade Table Head 10 rows in v12:")
trade_rows = df_v12.iloc[6:].copy().dropna(subset=[df_v12.columns[1]])
print(trade_rows.iloc[:10, :16].to_string())

print(f"\nTotal Trades in Q2 2025-26: {len(trade_rows)}")
print("Summary Card Row in v12 (Row index 2):")
print(df_v12.iloc[2, :13].to_string())
