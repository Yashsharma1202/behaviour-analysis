import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'
df = pd.read_excel(path, sheet_name='All_12Q_Options_Trades', engine='openpyxl')
print("Total rows in All_12Q_Options_Trades:", len(df))
print("\nUnique Quarters:", df['Quarter'].unique())
print("\nUnique Strategies:", df['Strategy'].unique())
print("\nSample columns:")
for col in df.columns:
    print(" -", col)
