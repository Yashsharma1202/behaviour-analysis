import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Futures_Master_Only.xlsx'

try:
    xl = pd.ExcelFile(path, engine='openpyxl')
    print('--- Sheet names in Futures_Master_Only.xlsx ---')
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        print(f"Sheet: '{sheet}', Shape: {df.shape}")
        print(f"  Columns: {df.columns.tolist()}")
except Exception as e:
    print(f"Error inspecting futures master: {e}")
