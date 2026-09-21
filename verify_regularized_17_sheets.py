import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master_v4_v2.xlsx'

xl = pd.ExcelFile(path, engine='openpyxl')
print(f"==========================================")
print(f"Verifying Regularized Futures Master: {path}")
print(f"Total Sheets: {len(xl.sheet_names)}")
print(f"Sheet List: {xl.sheet_names}")
print(f"==========================================")

df_exec = xl.parse('Exec_12Q_Combined_Summary')
print("\nExecutive Summary Table:")
print(df_exec.iloc[4:17, :9].to_string())

df_all = xl.parse('All_12Q_Futures_Trades')
print(f"\nAll 12Q Futures Trades: {len(df_all)} trades")
print(f"  Total Wins: {len(df_all[df_all['Booked Futures P&L (₹)'] > 0])} ({len(df_all[df_all['Booked Futures P&L (₹)'] > 0])/len(df_all):.2%})")
print(f"  Total Losses: {len(df_all[df_all['Booked Futures P&L (₹)'] <= 0])} ({len(df_all[df_all['Booked Futures P&L (₹)'] <= 0])/len(df_all):.2%})")
print(f"  Zero Entry Prices: {len(df_all[df_all['Entry Futures Price (₹)'] == 0])}")
print(f"  Zero Exit Prices: {len(df_all[df_all['Exit Futures Price (₹)'] == 0])}")
