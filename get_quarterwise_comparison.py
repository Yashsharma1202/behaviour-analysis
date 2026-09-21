import openpyxl
import pandas as pd
import sys

sys.stdout.reconfigure(errors='replace')

comp_path = r'D:/behaviour analysis/Nifty50_12_Quarters_Futures_vs_Options_Comparison.xlsx'
xl_comp = pd.ExcelFile(comp_path, engine='openpyxl')

df_sum = xl_comp.parse('Summary_&_Win_Rate_Compare', header=6)

print("==========================================================================================")
print("QUARTER-BY-QUARTER COMPARISON METRICS (ALL 12 QUARTERS)")
print("==========================================================================================")

for idx, r in df_sum.iterrows():
    qtr = r['Quarter']
    if 'TOTAL' in str(qtr) or 'AVERAGE' in str(qtr): continue
    
    f_wins = r['Futures Wins']
    f_wr   = r['Futures Win Rate (%)']
    f_pnl  = r['Futures P&L (₹)']
    
    o_wins = r['Options Wins']
    o_wr   = r['Options Win Rate (%)']
    o_pnl  = r['Options P&L (₹)']
    
    disc   = r['Futures Win / Options Loss Trades']
    
    print(f"\n📌 {qtr}:")
    print(f"  • Futures Model : P&L = +₹{f_pnl:,.2f} | Win Rate = {f_wr:.2%} ({f_wins})")
    print(f"  • Options Model : P&L = +₹{o_pnl:,.2f} | Win Rate = {o_wr:.2%} ({o_wins})")
    print(f"  • Discrepancy   : {disc} trades made profit in Futures but lost in Options (due to IV crush).")

print("==========================================================================================")
