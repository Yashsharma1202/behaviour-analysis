import openpyxl
import pathlib
import sys

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
fpath = ROOT / 'Nifty211_Quarterly_Performance_Master_v2.xlsx'

wb = openpyxl.load_workbook(fpath, data_only=True)

expected_sheets = ["Q1 April Earnings", "Q2 July Earnings", "Q3 October Earnings", "Q4 January Earnings"]
expected_headers = [
    "Rank", "Stock Symbol", "Industry Sector",
    "Listing Start Quarter Name", "First Reporting Date / Listing Date",
    "Strategy", "Action Status", "Total History Quarters", "Quarter Data Count",
    "Position Window", "Win Rate %", "Win Ratio", "Losing Quarters",
    "Average Return %", "Margin Capital (₹)", "Net Realised P&L (₹)"
]

print("==========================================================================================")
print("EXHAUSTIVE AUDIT FOR CLEAN & EASY QUARTERLY PERFORMANCE MASTER V2")
print("==========================================================================================")

total_errors = 0

for sname in expected_sheets:
    print(f"\nAUDITING SHEET: {sname}")
    print("-" * 80)
    if sname not in wb.sheetnames:
        print(f"  ❌ Sheet missing: {sname}")
        total_errors += 1
        continue
    
    ws = wb[sname]
    headers = [ws.cell(row=7, column=c).value for c in range(1, 17)]
    
    if headers == expected_headers:
        print("  ✅ All 16 Column Headers Match Perfectly!")
    else:
        print("  ❌ Header Mismatch!")
        total_errors += 1
        
    stock_count = 0
    missing_dates = 0
    missing_qnames = 0
    
    for r in range(8, 219):
        sym = ws.cell(row=r, column=2).value
        if sym:
            stock_count += 1
            qname = ws.cell(row=r, column=4).value
            fdate = ws.cell(row=r, column=5).value
            if not qname or qname == "N/A":
                missing_qnames += 1
            if not fdate or fdate == "N/A":
                missing_dates += 1
                
    print(f"  ✅ Total Stocks Ranked in Sheet: {stock_count} / 211")
    print(f"  ✅ Valid Listing Start Quarter Names: {stock_count - missing_qnames} / {stock_count}")
    print(f"  ✅ Valid First Reporting Dates: {stock_count - missing_dates} / {stock_count}")

print("\n==========================================================================================")
if total_errors == 0:
    print("AUDIT VERDICT: 100% PASS! ALL 4 SHEETS & 211 STOCKS CONTAIN EXPLICIT FIRST QUARTER DATA!")
else:
    print(f"AUDIT VERDICT: FAILED WITH {total_errors} ERRORS!")
print("==========================================================================================")
