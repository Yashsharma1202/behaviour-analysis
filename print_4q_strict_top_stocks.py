import openpyxl
import pathlib
import sys

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
fpath = ROOT / 'Nifty211_Four_Quarters_Strict_Ranked_Performance_Master.xlsx'

wb = openpyxl.load_workbook(fpath, data_only=True)

for q_sheet in ["Q1 (April Earnings)", "Q2 (July Earnings)", "Q3 (October Earnings)", "Q4 (January Earnings)"]:
    ws = wb[q_sheet]
    print(f"\n==========================================================================================")
    print(f"TOP 10 STRICTLY RANKED STOCKS IN {q_sheet.upper()}")
    print(f"==========================================================================================")
    for r in range(8, 18):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, 13)]
        rank, sym, strat, status, n_q, win_str, wr, w_c, l_c, avg_r, margin, pnl = row_vals
        print(f"Rank #{rank:<2} | {sym:<12} | {strat:<13} | {status:<20} | Q-Cycles: {n_q:<10} | Win Rate: {wr:<7} | Avg Ret: {avg_r:<7} | Net P&L: {pnl}")

print("==========================================================================================")
