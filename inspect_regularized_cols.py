import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Regularized_Master.xlsx'
xl = pd.ExcelFile(path, engine='openpyxl')

df_all = xl.parse('All_12_Quarters_Master_Trades')

print("All_12_Quarters_Master_Trades columns:")
for i, c in enumerate(df_all.columns):
    print(f"  {i}: {c}")

print("\nFirst 5 rows of All_12_Quarters_Master_Trades:")
print(df_all.head(5).to_string())
