import pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_OI_Master_v4.xlsx'

xl = pd.ExcelFile(path, engine='openpyxl')
print(f"==========================================")
print(f"Verifying Master Workbook: {path}")
print(f"Total Sheets: {len(xl.sheet_names)}")
print(f"Sheet List: {xl.sheet_names}")
print(f"==========================================")

for sheet in xl.sheet_names:
    df = xl.parse(sheet)
    print(f"\nSheet: '{sheet:<28}' | Shape: {str(df.shape):<10}")
    if sheet == 'Exec_12Q_Combined_Summary':
        print("Summary Head:")
        print(df.iloc[4:15, :8].to_string())
    elif sheet == 'All_12Q_Futures_Trades':
        zero_entries = len(df[df['Entry Futures Price (₹)'] == 0])
        zero_exits = len(df[df['Exit Futures Price (₹)'] == 0])
        print(f"  Zero Entry Prices: {zero_entries} | Zero Exit Prices: {zero_exits}")
