import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

v13_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'
eq_path  = r'D:/behaviour analysis/Nifty50_12_Quarters_Equity_Master_v1.xlsx'

xl_v13 = pd.ExcelFile(v13_path, engine='openpyxl')
df_v13 = xl_v13.parse('Exec_12Q_Combined_Summary')

xl_eq = pd.ExcelFile(eq_path, engine='openpyxl')
df_eq = xl_eq.parse('Exec_12Q_Combined_Summary')

print("==========================================================================================")
print("EXACT CAPITAL / FUND USED PER QUARTER (EQUITY VS FUTURES)")
print("==========================================================================================")

df_v13_sub = df_v13.iloc[5:17, [0, 1, 5, 6, 8]].copy()
df_v13_sub.columns = ['Quarter', 'Trades', 'Futures Margin Utilised (₹)', 'Net Futures P&L (₹)', 'Booked Return (%)']

df_eq_sub = df_eq.iloc[6:18, [0, 5, 6, 8]].copy()
df_eq_sub.columns = ['Quarter', 'Equity 1-Share Fund (₹)', 'Net Equity P&L (₹)', 'Equity Return (%)']

merged = pd.merge(df_v13_sub, df_eq_sub, on='Quarter')
print(merged.to_string())
print("==========================================================================================")
