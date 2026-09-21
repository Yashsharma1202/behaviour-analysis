import openpyxl, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Equity_Master_v2.xlsx'
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb['Exec_12Q_Combined_Summary']

print("Row-by-Row Values in Exec_12Q_Combined_Summary:")
for r in range(1, 25):
    vals = [str(ws.cell(row=r, column=c).value or '') for c in range(1, 12)]
    if any(vals):
        print(f"  Row {r:2d}: {vals}")
