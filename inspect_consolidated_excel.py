import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Consolidated.xlsx'

try:
    xl = pd.ExcelFile(path, engine='openpyxl')
    print("==========================================")
    print(f"File: {path}")
    print("Total Sheets:", len(xl.sheet_names))
    print("Sheets:", xl.sheet_names)
    print("==========================================")
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        print(f"\nSheet: '{sheet:<30}' | Shape: {str(df.shape):<10}")
        print("  Columns:", df.columns.tolist()[:10])
        if not df.empty:
            print("  Head 5 rows:")
            print(df.head(5).to_string())
except Exception as e:
    print(f"Error inspecting consolidated file: {e}")
