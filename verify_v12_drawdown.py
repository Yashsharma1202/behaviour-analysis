import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master_v12.xlsx'

xl = pd.ExcelFile(path, engine='openpyxl')
df_exec = xl.parse('Exec_12Q_Combined_Summary')
df_conc = xl.parse('Concurrent_Trades_Analysis')

print("==========================================================================================")
print("VERIFYING CORRECTED DRAWDOWN VALUES IN V12 MASTER WORKBOOK")
print("==========================================================================================")
print("\nExecutive Summary Sheet (Exec_12Q_Combined_Summary):")
print(df_exec.iloc[4:17, [0, 1, 4, 5, 6, 8, 9, 10, 11]].to_string())

print("\nConcurrent Trades Analysis Sheet (Concurrent_Trades_Analysis):")
print(df_conc.iloc[2:15, [0, 1, 2, 3, 5, 6, 7, 8, 9]].to_string())
print("==========================================================================================")
