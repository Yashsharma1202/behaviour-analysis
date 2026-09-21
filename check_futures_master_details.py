import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Futures_Master_Only.xlsx'

df = pd.read_excel(path, engine='openpyxl')
print("Non-null counts:")
print(df.notnull().sum())

print("\nSample row 0:")
print(df.iloc[0].to_dict())

print("\nSample row 1000:")
print(df.iloc[1000].to_dict())
