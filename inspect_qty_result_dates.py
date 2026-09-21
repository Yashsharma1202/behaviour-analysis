import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50Stocks_QtyResultDates.xlsx'

try:
    xl = pd.ExcelFile(path, engine='openpyxl')
    print('--- Sheet names in Nifty50Stocks_QtyResultDates.xlsx ---')
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        print(f"Sheet: '{sheet}', Shape: {df.shape}")
        print(f"  Columns: {df.columns.tolist()[:10]}")
        print("  Head:")
        print(df.head(2).to_string())
except Exception as e:
    print(f"Error inspecting result dates Excel: {e}")
