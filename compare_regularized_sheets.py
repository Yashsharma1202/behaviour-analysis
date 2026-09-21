import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Regularized_Master.xlsx'
xl = pd.ExcelFile(path, engine='openpyxl')

df_all = xl.parse('All_12_Quarters_Master_Trades')

print("--- Inspecting sheet Q3 2023-24 ---")
df_q3_raw = xl.parse('Q3 2023-24')
print("Shape:", df_q3_raw.shape)
print("Columns:", df_q3_raw.columns.tolist()[:10])
print("\nFirst 12 rows of Q3 2023-24:")
print(df_q3_raw.head(12).to_string())

# Compare Q3 trades in All_12_Quarters_Master_Trades
df_all_q3 = df_all[df_all['Quarter'] == 'Q3 2023-24'].reset_index(drop=True)
print("\n--- Q3 trades from All_12_Quarters_Master_Trades ---")
print(df_all_q3[['Trade No', 'Symbol', 'Strategy', 'Entry Date', 'Exit Date', 'Entry Price (?)', 'Exit Price (?)', 'Booked Profit/Loss (?)']].head(10).to_string())
