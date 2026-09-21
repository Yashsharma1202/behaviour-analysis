import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master_v6.xlsx'

xl = pd.ExcelFile(path, engine='openpyxl')
print("==========================================")
print(f"Verifying v6 Master Workbook: {path}")
print("Total Sheets:", len(xl.sheet_names))
print("Sheet List:", xl.sheet_names)
print("==========================================")

df_exec = xl.parse('Exec_12Q_Combined_Summary')
print("\nExecutive Summary Table:")
print(df_exec.iloc[4:17, :10].to_string())

df_all = xl.parse('All_12Q_Futures_Trades')
print(f"\nAll 12Q Futures Trades: {len(df_all)} trades")
ret_col = next(c for c in df_all.columns if 'Return' in c)
wins = len(df_all[df_all[ret_col] > 0])
losses = len(df_all[df_all[ret_col] <= 0])
print(f"  Total Wins: {wins} ({wins/len(df_all):.1%})")
print(f"  Total Losses: {losses} ({losses/len(df_all):.1%})")
print(f"  Total 12Q Futures Net P&L: ₹{df_all['Booked Futures P&L (₹)'].sum():,.2f}")
