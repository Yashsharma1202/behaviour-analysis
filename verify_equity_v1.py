import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Equity_Master_v1.xlsx'

xl = pd.ExcelFile(path, engine='openpyxl')
print("==========================================")
print(f"Verifying Equity Master Workbook: {path}")
print("Total Sheets:", len(xl.sheet_names))
print("Sheet List:", xl.sheet_names)
print("==========================================")

df_exec = xl.parse('Exec_12Q_Combined_Summary')
print("\nExecutive Summary Table:")
print(df_exec.iloc[4:17, :10].to_string())

df_q3 = xl.parse('Q3 2023-24')
print("\nSheet 'Q3 2023-24' Head 10 rows:")
print(df_q3.iloc[:10, :13].to_string())

df_all = xl.parse('All_12Q_Equity_Trades')
print(f"\nAll 12Q Equity Trades: {len(df_all)} trades")
print(f"Total Net Realised Profit: ₹{df_all['Realised Profit/Loss (₹)'].sum():,.2f}")
