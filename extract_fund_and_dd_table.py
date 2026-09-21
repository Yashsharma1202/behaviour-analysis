import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

v13_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master_v13.xlsx'

xl_v13 = pd.ExcelFile(v13_path, engine='openpyxl')
df_v13 = xl_v13.parse('Exec_12Q_Combined_Summary')

print("==========================================================================================")
print("EXTRACTING COMPLETE FUND USED & DRAWDOWN TABLE ACROSS ALL 12 QUARTERS")
print("==========================================================================================")

df_sub = df_v13.iloc[5:17, [0, 1, 2, 5, 6, 8, 9, 10, 11]].copy()
df_sub.columns = [
    'Quarter', 'Trades', 'Wins', 'Futures Margin Utilised (₹)',
    'Net Futures P&L (₹)', 'Return (%)', 'Realised Max DD (₹)',
    'Realised Max DD (%)', 'Daily MTM DD (₹)'
]

for col in ['Futures Margin Utilised (₹)', 'Net Futures P&L (₹)', 'Return (%)', 'Realised Max DD (₹)', 'Realised Max DD (%)', 'Daily MTM DD (₹)']:
    df_sub[col] = pd.to_numeric(df_sub[col], errors='coerce')

df_sub['Daily MTM DD (%)'] = df_sub['Daily MTM DD (₹)'] / df_sub['Futures Margin Utilised (₹)']

print(df_sub.to_string())
print("==========================================================================================")
