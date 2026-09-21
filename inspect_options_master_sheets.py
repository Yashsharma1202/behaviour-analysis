import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Options_OI_Master_v4.xlsx'

try:
    xl = pd.ExcelFile(path, engine='openpyxl')
    print('--- Sheet names in Options Master ---')
    for sheet in xl.sheet_names:
        print(f"Sheet: {sheet}")
        df = xl.parse(sheet, nrows=5)
        print(f"   Shape: {df.shape}")
        print(f"   Columns: {df.columns.tolist()[:5]}")
except Exception as e:
    print(f"Error inspecting options master: {e}")
