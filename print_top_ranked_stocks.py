import openpyxl
import pathlib
import sys

sys.stdout.reconfigure(errors='replace')

ROOT = pathlib.Path('D:/behaviour analysis')
fpath = ROOT / 'Nifty211_All_Stocks_Historical_WinRate_Ranking_Master.xlsx'

wb = openpyxl.load_workbook(fpath, data_only=True)
ws = wb['All 211 Stocks Ranking']

print("==========================================================================================")
print("TOP 25 STOCKS RANKED BY EMPIRICAL WIN RATE %")
print("==========================================================================================")

for r in range(17, 42):
    row_vals = [ws.cell(row=r, column=c).value for c in range(1, 13)]
    rank, sym, strat, status, n_q, win_str, wr, w_c, l_c, avg_r, margin, pnl = row_vals
    print(f"Rank #{rank:<2} | {sym:<12} | {strat:<13} | {status:<20} | Win Rate: {wr:<7} | Wins: {w_c}/{w_c+l_c} | 12-Qtr Net P&L: {pnl}")

print("==========================================================================================")
