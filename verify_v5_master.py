import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master_v5.xlsx'

xl = pd.ExcelFile(path, engine='openpyxl')
print("==========================================")
print(f"Verifying Master Workbook: {path}")
print("Total Sheets:", len(xl.sheet_names))
print("Sheet List:", xl.sheet_names)
print("==========================================")

df_exec = xl.parse('Exec_12Q_Combined_Summary')
print("\nExecutive Summary Table:")
print(df_exec.iloc[4:17, :10].to_string())

df_q3 = xl.parse('Q3 2023-24')
print("\nSheet 'Q3 2023-24' Head 10 rows:")
print(df_q3.iloc[:10, :14].to_string())

df_all = xl.parse('All_12Q_Futures_Trades')
print(f"\nAll 12Q Futures Trades: {len(df_all)} trades")
print(f"  Zero Entry Prices: {len(df_all[df_all['Entry Futures Price (₹)'] == 0])}")
print(f"  Zero Exit Prices: {len(df_all[df_all['Exit Futures Price (₹)'] == 0])}")
