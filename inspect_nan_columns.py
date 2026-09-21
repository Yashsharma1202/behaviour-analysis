import openpyxl, pandas as pd, pathlib, sys

sys.stdout.reconfigure(errors='replace')

src = r'D:\behaviour analysis\Nifty50_12_Quarters_Point_In_Time_Futures_Master.xlsx'
xl = pd.ExcelFile(src, engine='openpyxl')

for sheet in xl.sheet_names[1:4]:
    df = xl.parse(sheet, header=6)
    print(f"\n--- Sheet: {sheet} ---")
    print("Columns:", list(df.columns))
    cols = [c for c in ['Symbol', 'Entry Date', 'Assigned Slot', 'Re-entry Type', 'Quarterly Result Date', 'Slot', 'Reentry'] if c in df.columns]
    print(df[cols].head(10))
