import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'
xl = pd.ExcelFile(path, engine='openpyxl')

for sheet in xl.sheet_names:
    print(f"\n==========================================")
    print(f"SHEET: {sheet}")
    print(f"==========================================")
    df = xl.parse(sheet)
    print("Shape:", df.shape)
    print("Head 10 rows:")
    print(df.head(10).to_string())
