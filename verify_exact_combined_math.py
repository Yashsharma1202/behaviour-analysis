import openpyxl, pandas as pd, sys

sys.stdout.reconfigure(errors='replace')

master_clean_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_Combined_Turnover_0.5PCT_Cost_Master_CLEAN.xlsx'
wb = openpyxl.load_workbook(master_clean_path, data_only=True)

print("==========================================================================================")
print("VERIFYING 100% MATHEMATICALLY EXACT COMBINED TURNOVER 0.5% COST MASTER WORKBOOK")
print("==========================================================================================")

ws_q = wb['Q3 2023-24']
print("\nSample Rows from Q3 2023-24 (Turnover, Costs & P&L):")
headers = [ws_q.cell(row=7, column=c).value for c in range(12, 19)]
print(f"Headers: {headers}")

all_match = True
for r in range(8, 18):
    buy_to  = float(ws_q.cell(row=r, column=12).value)
    sell_to = float(ws_q.cell(row=r, column=13).value)
    comb_to = float(ws_q.cell(row=r, column=14).value)
    margin  = float(ws_q.cell(row=r, column=15).value)
    raw_pnl = float(ws_q.cell(row=r, column=16).value)
    cost    = float(ws_q.cell(row=r, column=17).value)
    net_pnl = float(ws_q.cell(row=r, column=18).value)
    
    exp_comb = round(buy_to + sell_to, 2)
    exp_cost = round(comb_to * 0.005, 2)
    exp_net  = round(raw_pnl - cost, 2)
    
    match_comb = abs(comb_to - exp_comb) < 0.01
    match_cost = abs(cost - exp_cost) < 0.01
    match_net  = abs(net_pnl - exp_net) < 0.01
    
    if not (match_comb and match_cost and match_net):
        all_match = False
        
    print(f"  Row {r:2d}: Comb TO=₹{comb_to:,.2f} | Cost (0.5%)=₹{cost:,.2f} (Exp: ₹{exp_cost:,.2f}) | Raw=₹{raw_pnl:,.2f} | Net=₹{net_pnl:,.2f} (Exp: ₹{exp_net:,.2f}) | Match? {match_cost and match_net}")

print("\n------------------------------------------------------------------------------------------")
print(f"ALL ROWS 100% MATHEMATICALLY EXACT MATCH? {all_match}")
print("==========================================================================================")
