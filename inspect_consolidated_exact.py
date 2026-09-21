import pandas as pd, openpyxl, sys

sys.stdout.reconfigure(errors='replace')

path = r'D:/behaviour analysis/Nifty50_12_Quarters_Consolidated.xlsx'
wb = openpyxl.load_workbook(path, data_only=True)

print("==========================================")
print("INSPECTING Nifty50_12_Quarters_Consolidated.xlsx")
print("==========================================")
print("Total Sheets:", len(wb.sheetnames))
print("Sheet Names :", wb.sheetnames)

for sheet in wb.sheetnames[:4]:
    ws = wb[sheet]
    print(f"\nSheet '{sheet}': max_row={ws.max_row}, max_col={ws.max_column}")
    print("Row 1-8 sample:")
    for r in range(1, min(ws.max_row + 1, 9)):
        row_vals = [str(ws.cell(row=r, column=c).value or '') for c in range(1, min(ws.max_column + 1, 15))]
        print(f"  Row {r}: {row_vals[:8]}")

print("==========================================")
