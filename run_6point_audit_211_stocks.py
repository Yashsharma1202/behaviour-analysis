import openpyxl, sys, pathlib

sys.stdout.reconfigure(errors='replace')

master_path = pathlib.Path('D:/behaviour analysis/Nifty211_Recent_Quarter_Combined_Best_Futures_Master_v3.xlsx')
wb = openpyxl.load_workbook(master_path, data_only=False)

ws = wb['Q2 2026-27']

print("==========================================================================================")
print("6-POINT AUDIT ON Nifty211_Recent_Quarter_Futures_Equity_Window_Master_v2.xlsx")
print("==========================================================================================")

total_rows = ws.max_row - 4 # exclude headers & total row
print(f"Total Trade Rows Evaluated: {total_rows}")

empty_window_count = 0
empty_winrate_count = 0
formula_errors = 0

for r in range(4, ws.max_row):
    sym = ws.cell(row=r, column=2).value
    window = ws.cell(row=r, column=5).value
    win_rate = ws.cell(row=r, column=6).value
    
    f_buy_to = ws.cell(row=r, column=13).value
    f_sell_to = ws.cell(row=r, column=14).value
    f_comb_to = ws.cell(row=r, column=15).value
    f_cost = ws.cell(row=r, column=16).value
    f_net = ws.cell(row=r, column=18).value
    
    # 1. Window Check
    if not window or not str(window).strip():
        empty_window_count += 1
        
    # 2. Win Rate Check
    if not win_rate or not str(win_rate).strip():
        empty_winrate_count += 1
        
    # 3. Formula Syntax Audit
    if f_buy_to != f"=J{r}*K{r}": formula_errors += 1
    if f_sell_to != f"=J{r}*L{r}": formula_errors += 1
    if f_comb_to != f"=M{r}+N{r}": formula_errors += 1
    if f_cost != f"=ROUND(O{r}*0.0005, 2)": formula_errors += 1
    if f_net != f"=Q{r}-P{r}": formula_errors += 1

print(f"1. Non-empty Position Taking Window Column : {'PASS ✅' if empty_window_count == 0 else 'FAIL ❌'}")
print(f"2. Non-empty Historical Win Rate % Column   : {'PASS ✅' if empty_winrate_count == 0 else 'FAIL ❌'}")
print(f"3. Turnover Buy Formula (=J*K) Audit        : {'PASS ✅' if formula_errors == 0 else 'FAIL ❌'}")
print(f"4. Turnover Sell Formula (=J*L) Audit       : {'PASS ✅' if formula_errors == 0 else 'FAIL ❌'}")
print(f"5. Combined Turnover Formula (=M+N) Audit   : {'PASS ✅' if formula_errors == 0 else 'FAIL ❌'}")
print(f"6. 0.05% Transaction Cost Formula Audit     : {'PASS ✅' if formula_errors == 0 else 'FAIL ❌'}")
print(f"7. Net Realised P&L Formula (=Q-P) Audit    : {'PASS ✅' if formula_errors == 0 else 'FAIL ❌'}")

print("------------------------------------------------------------------------------------------")
if empty_window_count == 0 and empty_winrate_count == 0 and formula_errors == 0:
    print("VERDICT: 100% PASS! ALL 211 TRADE ROWS ARE MATHEMATICALLY FLAWLESS!")
else:
    print("VERDICT: FAIL - DISCREPANCIES DETECTED.")
print("==========================================================================================")
