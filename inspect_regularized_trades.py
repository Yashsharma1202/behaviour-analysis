import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Regularized_Master.xlsx'
df_reg = pd.read_excel(path, sheet_name='All_12_Quarters_Master_Trades', engine='openpyxl')

print("Total trades in Regularized Master:", len(df_reg))
print("\nStrategy distribution:")
print(df_reg['Strategy'].value_counts())

print("\nQuarter distribution:")
print(df_reg['Quarter'].value_counts())

print("\nSample row 0:")
print(df_reg.iloc[0].to_dict())

print("\nSample row 1:")
print(df_reg.iloc[1].to_dict())
