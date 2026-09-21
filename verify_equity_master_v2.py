import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Equity_Master_v2.xlsx'

xl = pd.ExcelFile(path, engine='openpyxl')
df_exec = xl.parse('Exec_12Q_Combined_Summary')

print("==========================================================================================")
print("VERIFYING EQUITY MASTER V2 WORKBOOK")
print("==========================================================================================")
print("Sheet Names:", xl.sheet_names)
print("\nExecutive Summary Sheet (Exec_12Q_Combined_Summary):")
print(df_exec.iloc[5:18, [0, 1, 4, 5, 6, 8, 9, 10]].to_string())

print("\nTotal Equity Trades in All_12Q_Equity_Trades:", len(xl.parse('All_12Q_Equity_Trades')))
print("==========================================================================================")
