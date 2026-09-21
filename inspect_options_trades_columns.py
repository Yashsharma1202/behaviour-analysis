import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'

try:
    df = pd.read_excel(path, sheet_name='All_12Q_Options_Trades', engine='openpyxl')
    print('--- All_12Q_Options_Trades Columns ---')
    for i, col in enumerate(df.columns):
        print(f"{i}: {col}")
    print("\nHead 2 rows:")
    print(df.head(2).to_string())
except Exception as e:
    print(f"Error reading All_12Q_Options_Trades: {e}")
