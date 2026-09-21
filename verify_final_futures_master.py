import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

p1 = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master.xlsx'
p2 = r'D:/behaviour analysis/Futures_Master_Only_v2.xlsx'

for path in [p1, p2]:
    print(f"\n==========================================")
    print(f"Verifying workbook: {path}")
    print(f"==========================================")
    try:
        xl = pd.ExcelFile(path, engine='openpyxl')
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            print(f"\nSheet: '{sheet}' | Total Rows: {len(df)} | Columns: {len(df.columns)}")
            for i, col in enumerate(df.columns):
                non_nulls = df[col].notnull().sum()
                print(f"  Col {i:2d}: {str(col):<28} | Non-Null: {non_nulls}")
            print("\n  Sample row 0:")
            print("  ", df.iloc[0].to_dict())
    except Exception as e:
        print(f"Error inspecting {path}: {e}")
