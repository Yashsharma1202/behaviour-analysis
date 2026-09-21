import openpyxl
import pathlib
import sys

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
fpath = ROOT / 'Nifty211_Quarterly_Performance_Master.xlsx'

wb = openpyxl.load_workbook(fpath, data_only=True)

for q_sheet in ["Q1 April Earnings", "Q2 July Earnings", "Q3 October Earnings", "Q4 January Earnings"]:
    ws = wb[q_sheet]
    print(f"\n==========================================================================================")
    print(f"TOP 5 STOCKS IN {q_sheet.upper()}")
    print(f"==========================================================================================")
    for r in range(8, 13):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 15)]
        rank, sym, sec, strat, status, tot_q, spec_q, win_str, wr, ratio_str, l_c, avg_r, margin, pnl = row_vals
        print(f"Rank #{rank:<2} | {sym:<12} | Sector: {sec:<28} | {strat:<13} | Total Hist: {tot_q:<12} | Qtr Count: {spec_q:<11} | Win Ratio: {ratio_str:<20} | Win Rate: {wr:<7} | Avg Ret: {avg_r:<7} | Net P&L: {pnl}")

print("==========================================================================================")
