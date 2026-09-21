import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Regularized_Master.xlsx'

try:
    xl = pd.ExcelFile(path, engine='openpyxl')
    print("==========================================")
    print(f"File: {path}")
    print("Sheets:", xl.sheet_names)
    print("==========================================")
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        print(f"\nSheet: '{sheet:<30}' | Rows: {len(df)} | Cols: {len(df.columns)}")
        print("  Columns:", df.columns.tolist()[:12])
        if not df.empty:
            print("  Head row 0:", df.iloc[0].to_dict())
except Exception as e:
    print(f"Error inspecting regularized master: {e}")
