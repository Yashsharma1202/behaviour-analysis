import openpyxl
import sys
sys.stdout.reconfigure(errors='replace')
import pathlib

ROOT = pathlib.Path('D:/behaviour analysis')
fpath = ROOT / 'Nifty211_Past_4_Quarters_Combined_Best_Futures_Master_v6.xlsx'

wb = openpyxl.load_workbook(fpath, data_only=True)
ws_sum = wb['Performance Summary']

print("==========================================================================================")
print("AUDITING PERFORMANCE SUMMARY SHEET WIN RATE CELLS")
print("==========================================================================================")

for r in range(1, 40):
    row_vals = [ws_sum.cell(row=r, column=c).value for c in range(1, 12)]
    if any(row_vals):
        print(f"Row {r:2d}: {row_vals}")

print("==========================================================================================")
