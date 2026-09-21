import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Regularized_Master.xlsx'
xl = pd.ExcelFile(path, engine='openpyxl')
df_q3 = xl.parse('Q3 2023-24')

print("--- Q3 2023-24 Rows 0 to 7 ---")
for r in range(8):
    row_vals = [str(x) if pd.notnull(x) else '' for x in df_q3.iloc[r].values]
    print(f"Row {r+1:2d}: {row_vals[:13]}")
